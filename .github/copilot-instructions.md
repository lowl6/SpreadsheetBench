# Copilot Instructions for SpreadsheetBench

## Project Overview
**SpreadsheetBench** is a NeurIPS 2024 benchmark for real-world spreadsheet manipulation, featuring 912 questions from Excel forums with diverse, non-standard spreadsheet formats. It evaluates LLMs/agents on realistic tasks with OJ-style multi-test-case evaluation.

### Core Architecture (3-Layer Design)
1. **Inference Layer** (`inference/`): LLM-based code generation with multi-round reasoning
   - Single-round baseline, Multi-round ReAct, SheetCopilot v1/v2 agents
   - All approaches generate Python code that manipulates Excel files via `openpyxl`
   - **SheetCopilot v2** uses 6-stage pipeline: Observation → Understanding → Planning → Implementation → **Validation with Execution** → Final Execution & Revision
2. **Execution Layer** (`code_exec_docker/`): Isolated Docker-based Python runtime
   - Jupyter kernel server running in container at `http://localhost:8080/execute`
   - Conversation-scoped kernels tracked by `conv_id` for stateful execution
   - **Stage 5 Validation** now executes code early to verify results reasonableness before final execution
3. **Evaluation Layer** (`evaluation/`): Windows-only Excel automation for ground truth comparison
   - Uses `win32com` to open/recalculate spreadsheets before comparing with `openpyxl`
   - Compares cell values, formulas, and formatting (fill colors, font colors)

## Critical Workflows & Commands

### Environment Setup (Python 3.11/3.12 Recommended)
```powershell
# Python 3.13 has pandas compatibility issues - use Conda:
conda create -n sheetbench_py312 python=3.12 -y
conda activate sheetbench_py312
pip install -r requirements.txt
```

### Code Execution Environment (REQUIRED before inference)
```bash
cd code_exec_docker
# Edit config.json: set "volumes_path" to absolute path of dataset folder
docker build -t xingyaoww/codeact-execute-api -f Dockerfile.api .
docker build -t xingyaoww/codeact-executor -f Dockerfile.executor .
bash start_jupyter_server.sh 8080  # Starts API on port 8080
```

### Inference Workflows
```powershell
cd inference

# Single-round (baseline): Direct code generation
.\scripts\inference_single.ps1  # Edit model/api_key/base_url inside

# Multi-round with execution feedback:
.\scripts\inference_multiple_row_exec.ps1  # 5-row preview + error feedback
.\scripts\inference_multiple_react_exec.ps1  # ReAct reasoning + error feedback

# SheetCopilot v2 (6-stage enhanced agent with intelligent validation):
.\scripts\sheetcopilot_v2.ps1  # Observation→Understanding→Planning→Implementation→Validation(Execute+Verify)→Final Execution
# Key innovation: Stage 5 EXECUTES code and reads answer_position to verify reasonableness before final execution
```

**Inference Outputs:**
- Conversation logs: `data/{dataset}/outputs/conv_{setting}_{model}.jsonl`
- Generated spreadsheets: `data/{dataset}/outputs/{setting}_{model}/*.xlsx`
- Execution logs: `inference/log/` or `../log/`

### Evaluation (Windows ONLY)
```powershell
cd evaluation
.\scripts\evaluation.ps1  # Edit --setting and --model to match inference run
# Results: outputs/eval_{setting}_{model}.json with per-test-case pass/fail
```

## Project-Specific Conventions

### Data Structure Pattern
```json
{
  "id": "59196",
  "instruction": "Natural language task from real Excel forum",
  "spreadsheet_path": "spreadsheet/59196",  // Contains 1_{id}_input.xlsx, 1_{id}_answer.xlsx, 2_...
  "instruction_type": "Cell-Level Manipulation" | "Sheet-Level Manipulation",
  "answer_position": "'Sheet1'!H3:H5"  // Target cells to modify (single cell or range)
}
```

### Critical Code Generation Patterns (See `SHEETCOPILOT_V2_DESIGN.md`)

**❌ ANTI-PATTERN: Hard-coded cell references**
```python
data = ws['A1:D10']  # Assumes data starts at A1!
```

**✅ CORRECT: Dynamic cell references**
```python
# Based on observation stage, data actually starts at D3:G5
data_start_row, data_start_col = 3, 4  # From observation
for row in range(data_start_row, data_end_row + 1):
    cell = ws.cell(row, data_start_col)
```

**❌ ANTI-PATTERN: Ignoring empty cells**
```python
max_val = max([ws.cell(r, c).value for c in range(4, 8)])  # Crashes on None
```

**✅ CORRECT: Null checks**
```python
values = [ws.cell(r, c).value for c in range(4, 8) 
          if ws.cell(r, c).value is not None]
max_val = max(values) if values else 0
```

**Critical: answer_position parsing**
```python
# Handles multi-sheet references like 'Sheet1'!H3:H5 or A1:B2
import re
sheet_match = re.match(r"'([^']+)'!(.+)", answer_position)
if sheet_match:
    target_sheet = wb[sheet_match.group(1)]
    target_range = sheet_match.group(2)
else:
    target_sheet = wb.active
    target_range = answer_position
```

### Code Execution Client Pattern
```python
from jupyter_kernel_cli import ClientJupyterKernel
from code_exec import extract_code, exec_code

client = ClientJupyterKernel('http://localhost:8080/execute', conv_id='UNIQUE_ID')
code = extract_code(llm_response)  # Extracts from ```python blocks
result = exec_code(client, code)  # Returns output or error traceback

# Error detection:
has_error = 'Error' in result or 'Traceback' in result
```

### LLM API Call Pattern with Retry
```python
from llm_api import get_llm_response

messages = ['user prompt 1', 'assistant response 1', 'user prompt 2', ...]
response = get_llm_response(messages, opt, max_retries=3, timeout=120)
# Handles exponential backoff (2s, 4s, 8s) on failures
```

### Excel Recalculation for Dynamic Arrays (Windows Only)
```python
from excel_recalc import recalc_workbook

# After code execution, force Excel to recalculate formulas and materialize spills
recalc_workbook(output_path, materialize_dynamic=True, strip_dynamic_formula=True)
# Use flags: --excel-recalc --materialize-dynamic --strip-dynamic-formula in CLI
```

## Integration Points & Dependencies

### Docker Container Communication
- **API**: `code_exec_docker/api.py` runs Tornado server exposing `/execute` endpoint
- **Executor**: Manages per-conversation Jupyter kernels via Docker SDK
- **Volume Mounting**: `config.json` specifies host path mapped to `/mnt/data` in container
- **Timeout**: 30s for HTTP requests; kernel lifetime = conversation lifetime

### Evaluation Cell Comparison Logic
```python
# evaluation/evaluation.py uses multi-stage comparison:
# 1. Value equality (with type coercion for numbers)
# 2. Fill color matching (ignoring alpha channel, last 6 RGB chars)
# 3. Font color matching
# Cell-Level: All cells in answer_position must match
# Sheet-Level: Entire sheet(s) must match
```

### Common Pitfalls & Solutions

**Pitfall 1: Non-standard table layouts**
- Real spreadsheets don't start at A1; have merged cells, multiple tables per sheet
- Solution: SheetCopilot v2's Stage 1 Deep Observation analyzes actual data boundaries

**Pitfall 2: Dynamic array formulas not spilling**
- `openpyxl` writes formulas but Excel doesn't auto-recalculate on load
- Solution: Use `excel_recalc.py` with `--materialize-dynamic` to convert spills to static values

**Pitfall 3: Pandas import errors on Python 3.13**
- Pre-built wheels unavailable; Meson build fails
- Solution: Use Python 3.11 or 3.12 via Conda (documented in README)

**Pitfall 4: Evaluation fails silently on Linux/Mac**
- `win32com` not available; Excel COM automation required
- Solution: Run evaluation only on Windows; inference is cross-platform

## Key Files for Understanding Architecture

- `inference/sheetcopilot_v2.py`: 6-stage agent implementation (lines 1-100 show structure)
- `inference/prompt_format.py`: Prompt templates for different inference modes
- `code_exec_docker/jupyter.py`: Kernel lifecycle management (lines 1-61)
- `evaluation/evaluation.py`: Cell comparison logic (lines 1-81)
- `SHEETCOPILOT_V2_DESIGN.md`: Comprehensive architectural evolution documentation

## Testing & Debugging

**Quick System Test:**
```powershell
cd inference
python test_sheetcopilot.py  # Validates tools, logging, prompts (if exists)
```

**View Detailed Logs:**
```powershell
Get-Content ../log/sheetcopilot_v2_*.log -Tail 100
# Contains all 6 stage prompts, responses, code, errors with timestamps
```

**Inspect Failed Tasks:**
```python
import json
with open('data/test1/outputs/conv_sheetcopilot_glm-4.5-air.jsonl') as f:
    results = [json.loads(line) for line in f]
failed = [r for r in results if not r.get('success', False)]
# Analyze: error messages, revision_count, stage_history
```

---
**Last Updated:** 2024-11 | **Datasets:** test1 (10), sample_data_200 (200), all_data_912_v0.1 (912)
