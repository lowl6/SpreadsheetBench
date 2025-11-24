"""
SheetCopilot v2: Enhanced multi-stage reasoning for real-world spreadsheets
Specifically designed for SpreadsheetBench's two key characteristics:
1. Complex Instructions from Real World - Natural language understanding
2. Spreadsheet in Diverse Formats - Non-standard tables, multiple tables, rich formats

Architecture: Observing → Understanding → Planning → Implementing → Validating → Executing
"""
# 修改日志记录脚本，只保留最关键对优化模型有用的信息，比如生成的脚本与测试记录，保持记录和代码的简洁
import os
import re
import json
import logging
import argparse
from tqdm import tqdm
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from llm_api import get_llm_response
from code_exec import get_exec_client, extract_code, exec_code
from excel_calculator import calculate_formulas
from excel_recalc import recalc_workbook
from stage_prompts import (
    build_stage1_summary,
    build_stage2_prompt,
    build_stage3_prompt,
    build_stage4_prompt,
    build_stage5_failure_prompt,
    build_stage5_success_prompt,
    build_stage6_revision_prompt,
)


def setup_logger(log_dir: str, model_name: str, dataset_name: str = None) -> logging.Logger:
    """Setup comprehensive logging system"""
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Format: timestamp_dataset_model.log
    dataset_part = f"_{dataset_name}" if dataset_name else ""
    log_file = f"{log_dir}/{timestamp}{dataset_part}_{model_name}.log"
    
    logger = logging.getLogger('SheetCopilot_v2')
    logger.setLevel(logging.DEBUG)
    
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    
    formatter = logging.Formatter(
        '[%(asctime)s] [%(levelname)s] [%(funcName)s]\n%(message)s\n',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    fh.setFormatter(formatter)
    ch.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


class SheetCopilotV2:
    """
    Enhanced SheetCopilot with improved understanding of:
    - Non-standard table layouts (data not starting at A1)
    - Multiple tables in single/multiple sheets
    - Complex natural language instructions from real users
    - Rich formatting and non-textual elements
    """
    
    def __init__(self, opt, logger: logging.Logger):
        self.opt = opt
        self.logger = logger
        self.client = get_exec_client(opt.code_exec_url, opt.conv_id)
        self.stage_history = []
        self.max_revisions = getattr(opt, 'max_revisions', 3)
        self.enable_timing = getattr(opt, 'enable_timing', True)  # 控制是否启用计时
        self.debug = getattr(opt, 'debug', False)  # 控制是否输出详细调试信息
        self.stage_timings = {}  # 存储每个阶段的运行时间
        
    def log_stage(self, stage: str, content: str, stage_time: float = None):
        """Track stage execution (minimal logging)"""
        stage_record = {
            'stage': stage,
            'content': content[:200],  # Truncate for storage
            'timestamp': datetime.now().isoformat()
        }
        if self.enable_timing and stage_time is not None:
            stage_record['duration'] = stage_time
        self.stage_history.append(stage_record)
    
    def stage_1_deep_observation(self, file_path: str, instruction: str, 
                                 answer_position: str, instruction_type: str) -> Dict[str, Any]:
        """
        Stage 1: Deep Observation - Understanding NON-STANDARD spreadsheet structures
        
        Key focus areas for diverse formats:
        1. Identify actual data region (not assuming A1 start)
        2. Detect multiple tables in single sheet
        3. Find headers and their positions
        4. Detect merged cells, empty regions, formatting
        5. Understand multi-sheet relationships
        
        NOTE: This stage uses pre-defined observation code instead of LLM generation
        to ensure stability and avoid common errors (e.g., incorrect imports).
        """
        import time
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 1: Deep Observation")
        
        # Pre-defined observation code - no LLM generation needed
        # Use triple quotes without f-string to avoid escaping issues
        observation_code = """import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import re

# Phase 1: Global Structure Analysis
wb = openpyxl.load_workbook('""" + file_path + """')

print("📊 WORKBOOK STRUCTURE:")
print(f"All sheets: {wb.sheetnames}")
print(f"Active sheet: {wb.active.title}")

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"\\n--- Sheet: {sheet_name} ---")
    print(f"Dimensions: {ws.max_row} rows × {ws.max_column} cols")
    
    # Find actual data boundaries
    min_row, max_row = None, None
    min_col, max_col = None, None
    for row in range(1, ws.max_row + 1):
        if any(ws.cell(row, col).value is not None for col in range(1, ws.max_column + 1)):
            if min_row is None:
                min_row = row
            max_row = row
    for col in range(1, ws.max_column + 1):
        if any(ws.cell(row, col).value is not None for row in range(1, ws.max_row + 1)):
            if min_col is None:
                min_col = col
            max_col = col
    
    if min_row and min_col:
        print(f"Actual data region: Row {min_row}-{max_row}, Col {min_col}-{max_col}")
        print(f"Column letters: {get_column_letter(min_col)}-{get_column_letter(max_col)}")

# Phase 2: Target Position Analysis
target_str = \"""" + answer_position + """\"
sheet_match = re.match(r"'([^']+)'!(.+)", target_str)
if sheet_match:
    target_sheet = sheet_match.group(1)
    target_range = sheet_match.group(2)
    print(f"\\n🎯 TARGET: Sheet '{target_sheet}', Range '{target_range}'")
    ws = wb[target_sheet]
else:
    target_range = target_str
    ws = wb.active
    print(f"\\n🎯 TARGET: Active sheet, Range '{target_range}'")

print(f"\\n📍 TARGET CELL ANALYSIS:")
try:
    min_col, min_row, max_col, max_row = range_boundaries(target_range)
    print(f"Target range: {target_range}, min_row={min_row}, max_row={max_row}, min_col={min_col}, max_col={max_col}")
    total_rows = max_row - min_row + 1
    # For large ranges (>20 rows), show sample only
    if total_rows > 20:
        print(f"Large range detected ({total_rows} rows). Showing first 10 and last 5 rows as sample:")
        sample_rows = list(range(min_row, min(min_row + 10, max_row + 1))) + list(range(max(max_row - 4, min_row + 10), max_row + 1))
    else:
        sample_rows = range(min_row, max_row + 1)
    
    for row in sample_rows:
        values = []
        coords = []
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            values.append(cell.value)
            coords.append(cell.coordinate)
        print(f"Row {row}: {coords} = {values}")
    
    if total_rows > 20:
        print(f"... ({total_rows - 15} middle rows omitted) ...")
except Exception as e:
    print(f"⚠️ Could not analyze target range in detail: {str(e)}")
    print(f"Attempting to access as single cell or use fallback...")
    try:
        cell = ws[target_range]
        print(f"Single cell {cell.coordinate} = {cell.value}")
    except:
        print(f"Target range is complex. Will handle dynamically in code.")

# Phase 3: Context & Merged Cells
print(f"\\n🔗 MERGED CELLS:")
merged_ranges = ws.merged_cells.ranges
for merged in merged_ranges:
    print(f"Merged: {str(merged)}")

# Phase 4: Answer Position Format Analysis (as reference example)
print(f"\\n📝 ANSWER POSITION CURRENT CONTENT (Format Reference):")
try:
    # Read current content in answer position to understand expected format
    answer_content = []
    answer_has_data = False
    min_col, min_row, max_col, max_row = range_boundaries(target_range)
    
    # Sample up to 10 cells to understand format
    sample_limit = min(10, (max_row - min_row + 1) * (max_col - min_col + 1))
    cell_count = 0
    
    for row in range(min_row, max_row + 1):
        if cell_count >= sample_limit:
            break
        for col in range(min_col, max_col + 1):
            if cell_count >= sample_limit:
                break
            cell = ws.cell(row=row, column=col)
            cell_value = cell.value
            cell_type = type(cell_value).__name__
            cell_format = cell.number_format if hasattr(cell, 'number_format') else 'General'
            
            # Check if cell has formula (correct detection)
            if hasattr(cell, 'data_type') and cell.data_type == 'f':
                formula = cell.value  # This IS the formula string
                print(f"  {cell.coordinate}: FORMULA = {formula}")
            else:
                print(f"  {cell.coordinate}: VALUE = {cell_value} (type={cell_type}, format={cell_format})")
            
            if cell_value is not None and cell_value != '':
                answer_has_data = True
                answer_content.append({
                    'coord': cell.coordinate,
                    'value': str(cell_value)[:50],  # Truncate long values
                    'type': cell_type,
                    'format': cell_format
                })
            cell_count += 1
    
    if answer_has_data:
        print(f"\\n✅ Answer position contains data - USE AS FORMAT REFERENCE")
        print(f"   Total non-empty cells sampled: {len(answer_content)}")
        
        # Check if answer cells contain formulas
        formula_cells = []
        for row in range(min_row, min(min_row + 5, max_row + 1)):  # Check first 5 rows
            for col in range(min_col, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if hasattr(cell, 'data_type') and cell.data_type == 'f':
                    formula_cells.append((cell.coordinate, cell.value))
        
        if formula_cells:
            print(f"\\n   🎯 FORMULA PATTERN DETECTED:")
            print(f"   - {len(formula_cells)} cells contain formulas")
            print(f"   - First formula example: {formula_cells[0][1]}")
            if len(formula_cells) > 1:
                print(f"   - Second formula example: {formula_cells[1][1]}")
            print(f"   ⚠️ CRITICAL: Must generate FORMULAS, not static values!")
            print(f"   💡 Analyze the formula pattern and apply it with correct row references")
            
            # NEW: Extract referenced cells/ranges from formulas to understand data dependencies
            print(f"\\n   📊 DATA TYPE & DEPENDENCY ANALYSIS:")
            referenced_cells = set()
            for coord, formula in formula_cells[:3]:  # Analyze first 3 formulas
                # Extract cell references (A1, $B$2, etc.)
                import re
                refs = re.findall(r'\\$?[A-Z]+\\$?[0-9]+', formula)
                referenced_cells.update(refs)
            
            if referenced_cells:
                print(f"   - Referenced cells in formulas: {', '.join(sorted(referenced_cells)[:10])}")
                print(f"\\n   📋 DATA TYPE OF REFERENCED CELLS:")
                for ref in sorted(referenced_cells)[:5]:  # Show first 5
                    try:
                        ref_clean = ref.replace('$', '')
                        ref_cell = ws[ref_clean]
                        ref_value = ref_cell.value
                        ref_type = type(ref_value).__name__
                        ref_format = ref_cell.number_format if hasattr(ref_cell, 'number_format') else 'General'
                        print(f"   - {ref_clean}: {ref_value} (type={ref_type}, format={ref_format})")
                    except:
                        pass
        else:
            print(f"   Format patterns detected:")
            # Summarize types
            types_found = set(item['type'] for item in answer_content)
            formats_found = set(item['format'] for item in answer_content if item['format'] != 'General')
            print(f"   - Data types: {', '.join(types_found)}")
            if formats_found:
                print(f"   - Number formats: {', '.join(formats_found)}")
            
            # NEW: Answer Source Analysis - discover where answer values come from
            print(f"\\n   🔍 ANSWER SOURCE ANALYSIS (comparing with nearby columns):")
            # Get answer column letter
            answer_col_idx = min_col
            from openpyxl.utils import get_column_letter
            answer_col_letter = get_column_letter(answer_col_idx)
            
            # Check 3 rows before answer column (likely input columns)
            nearby_cols = range(max(1, answer_col_idx - 3), answer_col_idx)
            sample_rows_for_analysis = list(range(min_row, min(min_row + 5, max_row + 1)))
            
            print(f"   Comparing answer column {answer_col_letter} with nearby columns:")
            for row_idx in sample_rows_for_analysis:
                answer_val = ws.cell(row=row_idx, column=answer_col_idx).value
                if answer_val in (None, ""):
                    continue
                    
                row_info = f"   Row {row_idx}: Answer={answer_val}"
                
                # Check each nearby column to find potential source
                for col_idx in nearby_cols:
                    col_letter = get_column_letter(col_idx)
                    col_val = ws.cell(row=row_idx, column=col_idx).value
                    
                    # Check if answer is substring or exact match
                    if col_val and str(answer_val) in str(col_val):
                        row_info += f" | {col_letter}={col_val}"
                        if str(col_val) == str(answer_val):
                            row_info += " ✓EXACT"
                        else:
                            row_info += " ✓CONTAINS"
                
                print(row_info)
            
            print(f"\\n   💡 INSIGHT: Analyze which column contains the answer values!")
            print(f"   - If answer is EXACT match of a column → Simple copy")
            print(f"   - If answer is SUBSTRING of a column → Extract/parse logic needed")
            print(f"   - Check if other columns act as INDEX/CONDITION for extraction")
            
            # NEW: For non-formula cells, analyze actual data type patterns
            print(f"\\n   📊 REFERENCE DATA TYPE ANALYSIS:")
            for item in answer_content[:5]:  # Show first 5 cells
                print(f"   - {item['coord']}: type={item['type']}, value_sample={item['value'][:30]}")
    else:
        print(f"⚠️ Answer position is empty - no format reference available")
        
except Exception as e:
    print(f"⚠️ Could not analyze answer position format: {str(e)}")

# Phase 5: Pattern Recognition
print("\\n🎯 TASK PATTERN RECOGNITION:")
# Instruction and type are shown in logs, no need to print in code

wb.close()
"""
        
        # Execute the observation code (no logging to reduce verbosity)
        try:
            result = exec_code(self.client, observation_code)
            # Check for fatal errors only (warnings ⚠️ are OK)
            has_fatal_error = 'Traceback' in result or 'JSON_DECODE_ERROR' in result or 'EXECUTION REQUEST ERROR' in result
            
            # Check if result is actually SOURCE CODE (not executed)
            is_source_code = ('import openpyxl' in result and 'wb = openpyxl.load_workbook' in result and result.count('\n') < 5)
            
            # Check if we got minimal required info from execution
            has_basic_info = ('Target range:' in result or 'WORKBOOK STRUCTURE:' in result or 'All sheets:' in result)
            
            # Success: no fatal error, not source code, has basic info
            success = (not has_fatal_error) and (not is_source_code) and has_basic_info
            
            # Only log critical errors
        except Exception as e:
            result = f"Observation error: {str(e)}"
            success = False
        
        # Create a summary prompt for context (used in later stages)
        observation_summary = build_stage1_summary(
            instruction=instruction,
            instruction_type=instruction_type,
            answer_position=answer_position,
            file_path=file_path,
            observation_result=result,
        )

        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_1_observation'] = stage_time
            if self.debug:
                print(f"[DEBUG] ✅ Stage 1 completed in {stage_time:.2f}s")
            self.log_stage("1️⃣ DEEP OBSERVATION", 
                          f"Analyzing non-standard spreadsheet structure\nTask: {instruction}", 
                          stage_time)

        return {
            'prompt': observation_summary,
            'response': result,
            'code': observation_code,
            'result': result,
            'success': success
        }
    
    def stage_2_instruction_understanding(self, observation: Dict[str, Any], 
                                         instruction: str, instruction_type: str) -> Dict[str, Any]:
        """
        Stage 2: Instruction Understanding - Parse COMPLEX REAL-WORLD natural language
        
        Real-world instructions characteristics:
        - Long, descriptive, informal language
        - May contain background context and explanations
        - Multiple requirements in one sentence
        - Implicit assumptions and domain knowledge
        - References to attached files and examples
        """
        import time
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 2: Instruction Understanding")
        
        understanding_prompt = build_stage2_prompt(
            instruction=instruction,
            instruction_type=instruction_type,
            observation_result=observation['result'][:1200],  # 已截断
        )
        
        # Generate understanding (logging disabled for brevity)
        response = get_llm_response(
            [observation['prompt'], observation['response'], understanding_prompt], 
            self.opt
        )
        
        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_2_understanding'] = stage_time
            if self.debug:
                print(f"[DEBUG] ✅ Stage 2 completed in {stage_time:.2f}s")
            self.log_stage("2️⃣ INSTRUCTION UNDERSTANDING", 
                          "Parsing complex natural language from real user", 
                          stage_time)
        
        return {
            'prompt': understanding_prompt,
            'response': response,
            'structured_requirements': self._parse_requirements(response)
        }
    
    def stage_3_solution_planning(self, observation: Dict[str, Any], 
                                 understanding: Dict[str, Any],
                                 file_path: str, output_path: str,
                                 answer_position: str) -> Dict[str, Any]:
        """
        Stage 3: Solution Planning - Design robust solution for non-standard formats
        
        Planning must account for:
        - Dynamic cell references (not hardcoded A1)
        - Multiple table navigation
        - Empty cell handling
        - Format preservation
        - Edge cases from real-world data
        """
        import time
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 3: Solution Planning")
        
        planning_prompt = build_stage3_prompt(
            observation_result=observation['result'][:1000],
            understanding_result=understanding['response'][:800],
            file_path=file_path,
            output_path=output_path,
            answer_position=answer_position,
        )
        
        # Generate plan (logging disabled for brevity)
        messages = [
            observation['prompt'], observation['response'],
            understanding['prompt'], understanding['response'],
            planning_prompt
        ]
        response = get_llm_response(messages, self.opt)
        
        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_3_planning'] = stage_time
            if self.debug:
                print(f"[DEBUG] ✅ Stage 3 completed in {stage_time:.2f}s")
            self.log_stage("3️⃣ SOLUTION PLANNING", 
                          "Designing robust implementation for non-standard structure", 
                          stage_time)
        
        return {
            'prompt': planning_prompt,
            'response': response,
            'plan_steps': self._extract_steps(response)
        }
    
    def stage_4_code_implementation(self, observation: Dict[str, Any],
                                   understanding: Dict[str, Any],
                                   plan: Dict[str, Any],
                                   file_path: str, output_path: str,
                                   answer_position: str) -> Dict[str, Any]:
        """
        Stage 4: Code Implementation - Generate robust Python code
        
        Implementation principles:
        - Use dynamic cell references from observation
        - Handle all edge cases identified in planning
        - Include comprehensive error handling
        - Support diverse table formats
        """
        import time
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 4: Code Implementation")
        
        implementation_prompt = build_stage4_prompt(
            observation_result=observation['result'][:1000],
            understanding_result=understanding['response'][:800],
            planning_result=plan['response'],
            file_path=file_path,
            output_path=output_path,
            answer_position=answer_position,
        )
        
        # Generate implementation (prompt logging disabled)
        messages = [
            observation['prompt'], observation['response'],
            understanding['prompt'], understanding['response'],
            plan['prompt'], plan['response'],
            implementation_prompt
        ]
        response = get_llm_response(messages, self.opt)
        
        code = extract_code(response)
        # Replace placeholders with actual paths for current test case
        code = code.replace('{file_path}', file_path)
        code = code.replace('{output_path}', output_path)
        # Log generated code (KEY INFO for debugging)
        self.logger.info(f"[GENERATED CODE]\n{code}")
        
        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_4_implementation'] = stage_time
            if self.debug:
                print(f"[DEBUG] ✅ Stage 4 completed in {stage_time:.2f}s")
            self.log_stage("4️⃣ CODE IMPLEMENTATION", 
                          "Generating production-ready Python code", 
                          stage_time)
        
        return {
            'prompt': implementation_prompt,
            'response': response,
            'code': code
        }
    
    def stage_5_code_validation(self, implementation: Dict[str, Any],
                               observation: Dict[str, Any],
                               plan: Dict[str, Any],
                               file_path: str, output_path: str,
                               instruction: str, answer_position: str) -> Dict[str, Any]:
        """
        Stage 5: Code Validation - Execute and verify results
        
        NEW APPROACH:
        1. Execute the generated code first
        2. If successful, read the answer_position content from output
        3. Let AI judge if results are reasonable given the instruction
        4. If unreasonable or errors, provide corrected code
        """
        import time
        import openpyxl
        import re
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 5: Code Validation (Execute & Verify)")
        
        # Replace paths before execution
        code_to_execute = self._replace_hardcoded_paths(implementation['code'], file_path, output_path)
        
        # Execute code for validation
        try:
            exec_result = exec_code(self.client, code_to_execute)
            has_error = 'Error' in exec_result or 'Traceback' in exec_result or '❌' in exec_result
            # Only log if there's an error
            if has_error:
                self.logger.info(f"[VALIDATION ERROR]\n{exec_result}")
            
            if has_error:
                # Execution failed - do static validation only
                validation_prompt = build_stage5_failure_prompt(
                    instruction=instruction,
                    observation_result=observation['response'][:500],
                    planning_result=plan['response'][:600],
                    generated_code=implementation['code'],
                    execution_error=exec_result,
                )
                
            else:
                # Execution succeeded - verify result content and build rich feedback
                output_local = output_path.replace('/mnt/data/', '../data/')
                input_local = file_path.replace('/mnt/data/', '../data/')
                answer_content = "Could not read output file"
                input_answer_content = "Could not read input file"
                answer_values_raw = []
                input_answer_values_raw = []
                summary_json = {}
                input_summary_json = {}
                neighbor_alert = None
                
                # FIRST: Read input file's answer column to detect pattern
                try:
                    if os.path.exists(input_local):
                        wb_input = openpyxl.load_workbook(input_local, data_only=False)
                        sheet_match = re.match(r"'([^']+)'!(.+)", answer_position)
                        if sheet_match:
                            sheet_name = sheet_match.group(1)
                            cell_range = sheet_match.group(2)
                            ws_input = wb_input[sheet_name]
                        else:
                            cell_range = answer_position
                            ws_input = wb_input.active
                        
                        # Collect input answer values - handle both single cell and range
                        input_coords_lines = []
                        input_has_formulas = False
                        input_cell_or_range = ws_input[cell_range]
                        
                        if hasattr(input_cell_or_range, 'coordinate'):
                            # Single cell
                            if hasattr(input_cell_or_range, 'data_type') and input_cell_or_range.data_type == 'f':
                                input_coords_lines.append(f"{input_cell_or_range.coordinate}: FORMULA = {input_cell_or_range.value}")
                                input_has_formulas = True
                            else:
                                input_coords_lines.append(f"{input_cell_or_range.coordinate}: {input_cell_or_range.value}")
                            input_answer_values_raw.append(input_cell_or_range.value)
                        else:
                            # Range
                            for row in input_cell_or_range:
                                for cell in row:
                                    if hasattr(cell, 'data_type') and cell.data_type == 'f':
                                        input_coords_lines.append(f"{cell.coordinate}: FORMULA = {cell.value}")
                                        input_has_formulas = True
                                    else:
                                        input_coords_lines.append(f"{cell.coordinate}: {cell.value}")
                                    input_answer_values_raw.append(cell.value)
                        
                        input_answer_content = '\n'.join(input_coords_lines)
                        input_non_empty = [v for v in input_answer_values_raw if v not in (None, "")]
                        input_unique_vals = set(input_non_empty)
                        input_numeric_vals = [float(v) for v in input_non_empty if isinstance(v, (int,float))]
                        input_summary_json = {
                            "total_cells": len(input_answer_values_raw),
                            "non_empty_count": len(input_non_empty),
                            "unique_count": len(input_unique_vals),
                            "all_same": len(input_unique_vals) == 1,
                            "has_formulas": input_has_formulas,
                            "sample_values": list(input_non_empty[:10]),
                            "numeric_min": min(input_numeric_vals) if input_numeric_vals else None,
                            "numeric_max": max(input_numeric_vals) if input_numeric_vals else None,
                        }
                        wb_input.close()
                except Exception as e:
                    input_answer_content = f"Error reading input: {str(e)}"
                
                # SECOND: Read output file's answer column
                try:
                    if os.path.exists(output_local):
                        wb = openpyxl.load_workbook(output_local, data_only=False)
                        sheet_match = re.match(r"'([^']+)'!(.+)", answer_position)
                        if sheet_match:
                            sheet_name = sheet_match.group(1)
                            cell_range = sheet_match.group(2)
                            ws = wb[sheet_name]
                        else:
                            cell_range = answer_position
                            ws = wb.active
                        # Extract raw cells
                        def _col_letter_to_index(col_letters: str) -> int:
                            from openpyxl.utils import column_index_from_string
                            return column_index_from_string(col_letters)
                        def _parse_range(r: str):
                            if ':' not in r:
                                return r, r
                            return r.split(':', 1)
                        start_ref, end_ref = _parse_range(cell_range)
                        # Derive column letters
                        import string
                        import itertools
                        def _split_ref(ref: str):
                            m = re.match(r"([A-Z]+)([0-9]+)", ref)
                            return m.group(1), int(m.group(2)) if m else (None, None)
                        start_col_letters, start_row_num = _split_ref(start_ref)
                        end_col_letters, end_row_num = _split_ref(end_ref)
                        start_col_idx = _col_letter_to_index(start_col_letters)
                        end_col_idx = _col_letter_to_index(end_col_letters)
                        # Collect values - handle both single cell and range
                        cell_or_range = ws[cell_range]
                        coords_lines = []
                        has_formulas = False
                        
                        # Check if it's a single cell (not iterable) or a range
                        if hasattr(cell_or_range, 'coordinate'):
                            # Single cell
                            answer_values_raw.append(cell_or_range.value)
                            if hasattr(cell_or_range, 'data_type') and cell_or_range.data_type == 'f':
                                coords_lines.append(f"{cell_or_range.coordinate}: FORMULA = {cell_or_range.value}")
                                has_formulas = True
                            else:
                                coords_lines.append(f"{cell_or_range.coordinate}: {cell_or_range.value}")
                        else:
                            # Range - iterate through rows
                            for row in cell_or_range:
                                for cell in row:
                                    answer_values_raw.append(cell.value)
                                    if hasattr(cell, 'data_type') and cell.data_type == 'f':
                                        coords_lines.append(f"{cell.coordinate}: FORMULA = {cell.value}")
                                        has_formulas = True
                                    else:
                                        coords_lines.append(f"{cell.coordinate}: {cell.value}")
                        answer_content = '\n'.join(coords_lines)
                        non_empty = [v for v in answer_values_raw if v not in (None, "")]
                        unique_vals = set(non_empty)
                        numeric_vals = [float(v) for v in non_empty if isinstance(v, (int,float))]
                        
                        # NEW: Calculate formulas and verify results if formulas detected
                        calculated_values = []
                        calculated_coords_lines = []
                        suspicious_patterns = []
                        
                        if has_formulas:
                            try:
                                # Close current workbook before Excel calculation
                                wb.close()
                                
                                # Calculate formulas using Excel COM
                                calculate_formulas(output_local)
                                
                                # Reopen and read calculated values with data_only=True
                                wb_calc = openpyxl.load_workbook(output_local, data_only=True)
                                if sheet_match:
                                    ws_calc = wb_calc[sheet_name]
                                else:
                                    ws_calc = wb_calc.active
                                
                                # Handle both single cell and range
                                calc_cell_or_range = ws_calc[cell_range]
                                if hasattr(calc_cell_or_range, 'coordinate'):
                                    # Single cell
                                    calc_val = calc_cell_or_range.value
                                    calculated_values.append(calc_val)
                                    calculated_coords_lines.append(f"{calc_cell_or_range.coordinate}: {calc_val}")
                                else:
                                    # Range
                                    for row in calc_cell_or_range:
                                        for cell in row:
                                            calc_val = cell.value
                                            calculated_values.append(calc_val)
                                            calculated_coords_lines.append(f"{cell.coordinate}: {calc_val}")
                                
                                wb_calc.close()
                                
                                # Detect suspicious patterns in calculated results
                                calc_numeric = [float(v) for v in calculated_values if isinstance(v, (int, float)) and v is not None]
                                
                                # Pattern 1: All zeros (likely wrong column reference)
                                if calc_numeric and all(v == 0 for v in calc_numeric):
                                    suspicious_patterns.append("⚠️ ALL_ZEROS: All calculated numeric values are 0 - formula may reference wrong column or empty data")
                                
                                # Pattern 2: All identical non-zero values
                                if calc_numeric and len(set(calc_numeric)) == 1 and calc_numeric[0] != 0 and len(calc_numeric) > 1:
                                    suspicious_patterns.append(f"⚠️ ALL_SAME: All {len(calc_numeric)} values are identical ({calc_numeric[0]}) - may indicate formula copy error")
                                
                                # Pattern 3: Expected non-empty but got empty/zero
                                if not calc_numeric and non_empty:
                                    suspicious_patterns.append("⚠️ EMPTY_RESULT: Formula exists but calculated to empty/None - check formula validity")
                                
                                # Pattern 4: Check for invalid @ symbol in formula
                                if has_formulas:
                                    for formula_line in coords_lines:
                                        if 'FORMULA =' in formula_line:
                                            formula_text = formula_line.split('FORMULA =')[1].strip()
                                            if '@' in formula_text:
                                                suspicious_patterns.append(
                                                    f"⚠️ INVALID_AT_SYMBOL: Formula contains '@' (implicit intersection operator) which is invalid: {formula_text[:100]}. "
                                                    f"Remove all @ symbols. If comparing two ranges (e.g., A=C), use SUMPRODUCT(--(A3:A57=C3:C57), B3:B57) instead of SUMIFS."
                                                )
                                
                                # Pattern 5: Check if sum_range points to text column (CRITICAL for lookup/sum tasks)
                                if calc_numeric and all(v == 0 for v in calc_numeric) and has_formulas:
                                    # Extract formula text to analyze sum_range
                                    for formula_line in coords_lines:
                                        if 'FORMULA =' in formula_line:
                                            formula_text = formula_line.split('FORMULA =')[1].strip()
                                            # Check for SUM/SUMIF/SUMIFS patterns - extract first range (sum_range)
                                            sum_match = re.search(r'SUM(?:IF|IFS)?\s*\(\s*([A-Z]+\d+:[A-Z]+\d+)', formula_text, re.IGNORECASE)
                                            if sum_match:
                                                sum_range = sum_match.group(1)
                                                # Check data type in sum_range by sampling
                                                try:
                                                    range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', sum_range)
                                                    if range_match:
                                                        col_letter = range_match.group(1)
                                                        start_row = int(range_match.group(2))
                                                        # Sample first 3 non-empty cells to check type
                                                        sample_vals = []
                                                        for r in range(start_row, min(start_row + 10, start_row + 50)):
                                                            try:
                                                                val = ws_calc[f'{col_letter}{r}'].value
                                                                if val is not None:
                                                                    sample_vals.append(val)
                                                                    if len(sample_vals) >= 3:
                                                                        break
                                                            except:
                                                                pass
                                                        # Check if mostly text (>50% text values)
                                                        if sample_vals:
                                                            text_count = sum(1 for v in sample_vals if isinstance(v, str))
                                                            if text_count / len(sample_vals) > 0.5:
                                                                suspicious_patterns.append(
                                                                    f"⚠️ TEXT_COLUMN_SUM: Formula tries to sum TEXT column {sum_range}! "
                                                                    f"Sample values: {sample_vals[:3]}. Result is 0 because text cannot be summed. "
                                                                    f"For lookup/sum tasks, verify you're summing the NUMERIC VALUE column, not the TEXT LABEL column."
                                                                )
                                                except Exception as e:
                                                    pass
                                            break
                                
                                # Reopen for further processing (without data_only)
                                wb = openpyxl.load_workbook(output_local, data_only=False)
                                if sheet_match:
                                    ws = wb[sheet_name]
                                else:
                                    ws = wb.active
                                    
                            except Exception as calc_error:
                                suspicious_patterns.append(f"⚠️ CALC_ERROR: Formula calculation failed: {str(calc_error)[:100]}")
                                calculated_values = ["[Calculation Error]"]
                                calculated_coords_lines = [f"Error: {str(calc_error)[:100]}"]
                        
                        # Check for suspicious patterns even if no formulas (static values)
                        if not has_formulas and numeric_vals:
                            # Pattern: Static zero value for summation tasks
                            if all(v == 0 for v in numeric_vals) and ('sum' in instruction.lower() or 'total' in instruction.lower()):
                                suspicious_patterns.append("⚠️ STATIC_ZERO: Wrote static value 0 for summation task - may have summed wrong column or text data")
                        
                        summary_json = {
                            "total_cells": len(answer_values_raw),
                            "non_empty_count": len(non_empty),
                            "unique_count": len(unique_vals),
                            "all_same": len(unique_vals) == 1,
                            "has_formulas": has_formulas,
                            "sample_values": list(non_empty[:10]),
                            "numeric_min": min(numeric_vals) if numeric_vals else None,
                            "numeric_max": max(numeric_vals) if numeric_vals else None,
                            "numeric_mean": (sum(numeric_vals)/len(numeric_vals)) if numeric_vals else None,
                            "calculated_values": calculated_values[:10] if calculated_values else None,
                            "calculated_content": '\n'.join(calculated_coords_lines[:10]) if calculated_coords_lines else None,
                            "suspicious_patterns": suspicious_patterns if suspicious_patterns else None,
                        }
                        # Neighbor column leak detection (only if single column range)
                        if start_col_idx == end_col_idx:
                            right_col_idx = end_col_idx + 1
                            # Only check if within sheet bounds
                            if right_col_idx <= ws.max_column:
                                from openpyxl.utils import get_column_letter
                                right_letter = get_column_letter(right_col_idx)
                                leak_values = []
                                for r in range(start_row_num, end_row_num + 1):
                                    cv = ws.cell(row=r, column=right_col_idx).value
                                    if cv not in (None, ""):
                                        leak_values.append(cv)
                                if leak_values:
                                    neighbor_alert = {
                                        "right_column": right_letter,
                                        "non_empty_count": len(leak_values),
                                        "sample": leak_values[:10]
                                    }
                        wb.close()
                except Exception as e:
                    answer_content = f"Error reading output: {str(e)}"
                import json as _json
                answer_summary_block = _json.dumps(summary_json, ensure_ascii=False, indent=2)
                input_summary_block = _json.dumps(input_summary_json, ensure_ascii=False, indent=2)
                neighbor_alert_block = _json.dumps(neighbor_alert, ensure_ascii=False, indent=2) if neighbor_alert else "None"
                validation_prompt = build_stage5_success_prompt(
                    instruction=instruction,
                    observation_result=observation['response'][:600],
                    planning_result=plan['response'][:600],
                    generated_code=implementation['code'],
                    execution_stdout=exec_result,
                    answer_position=answer_position,
                    input_answer_content=input_answer_content[:800],
                    input_summary_json=input_summary_block,
                    output_answer_content=answer_content[:800],
                    output_summary_json=answer_summary_block,
                    neighbor_alert_json=neighbor_alert_block,
                )
        
        except Exception as e:
            exec_result = f"Exception during validation execution: {str(e)}"
            validation_prompt = build_stage5_failure_prompt(
                instruction=instruction,
                observation_result=observation['response'][:500],
                planning_result=plan['response'][:600],
                generated_code=implementation['code'],
                execution_error=exec_result,
            )
        
        # Generate validation response (logging disabled)
        messages = [
            observation['prompt'], observation['response'],
            plan['prompt'], plan['response'],
            implementation['prompt'], implementation['response'],
            validation_prompt
        ]
        response = get_llm_response(messages, self.opt)
        
        # Check if validation passed or code was corrected
        validation_passed = "VALIDATION PASSED" in response.upper()
        corrected_code = extract_code(response) if not validation_passed else None
        
        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_5_validation'] = stage_time
            if self.debug:
                print(f"[DEBUG] ✅ Stage 5 completed in {stage_time:.2f}s")
            self.log_stage("5️⃣ CODE VALIDATION", 
                          "Execute and verify results reasonableness", 
                          stage_time)
        
        return {
            'prompt': validation_prompt,
            'response': response,
            'passed': validation_passed,
            'corrected_code': corrected_code,
            'execution_result': exec_result,
            'issues_found': self._extract_issues(response),
            'answer_values_summary': summary_json if 'summary_json' in locals() else {},
            'neighbor_alert': neighbor_alert if 'neighbor_alert' in locals() else None
        }
    
    def stage_6_execution_and_revision(self, validation: Dict[str, Any],
                                      implementation: Dict[str, Any],
                                      observation: Dict[str, Any],
                                      plan: Dict[str, Any],
                                      file_path: str, output_path: str,
                                      instruction: str) -> Dict[str, Any]:
        """
        Stage 6: Execution and Revision - Execute code with retry mechanism
        
        Execution with smart revision:
        - Try validated/corrected code first
        - If errors, analyze and fix up to max_revisions times
        - Learn from execution feedback
        - Adapt to runtime errors
        """
        import time
        stage_start = time.time() if self.enable_timing else None
        
        if self.debug:
            print(f"\n[DEBUG] 🚀 Starting Stage 6: Final Execution & Revision")
        
        # Print stage header at the beginning
        if self.enable_timing:
            self.log_stage("6️⃣ EXECUTION & REVISION", 
                          "Running code with intelligent error recovery", 
                          None)  # Time will be updated at the end
        
        # Use corrected code from validation if available (handle None case)
        if validation is not None:
            code_to_execute = validation.get('corrected_code') or implementation['code']
        else:
            code_to_execute = implementation['code']
        
        # Replace hardcoded paths before execution (important for code reuse)
        code_to_execute = self._replace_hardcoded_paths(code_to_execute, file_path, output_path)
        
        for revision_num in range(self.max_revisions + 1):
            try:
                result = exec_code(self.client, code_to_execute)
                
                # Check for errors in output
                has_error = 'Error' in result or 'Traceback' in result or '❌' in result
                
                # Only log errors or final success
                if has_error:
                    self.logger.info(f"[EXECUTION ERROR - Attempt {revision_num + 1}]\n{result}")
                
                if not has_error:
                    if self.debug and self.enable_timing:
                        elapsed = time.time() - stage_start
                        print(f"[DEBUG] ✅ Stage 6 execution successful (attempt {revision_num + 1}, {elapsed:.2f}s)")
                    # Calculate formulas if output file exists (Windows only)
                    output_local = output_path.replace('/mnt/data/', '../data/')
                    if os.path.exists(output_local):
                        calculate_formulas(output_local, self.logger)
                        if getattr(self.opt, 'excel_recalc', False):
                            recalc_workbook(
                                input_path=output_local,
                                output_path=output_local,
                                materialize_dynamic=getattr(self.opt, 'materialize_dynamic', False),
                                strip_formula=getattr(self.opt, 'strip_dynamic_formula', False),
                                logger=self.logger,
                            )
                    if self.enable_timing:
                        stage_time = time.time() - stage_start
                        self.stage_timings['stage_6_execution'] = stage_time
                        if self.debug:
                            print(f"[DEBUG] ✅ Stage 6 completed in {stage_time:.2f}s (success after {revision_num} revisions)")
                    return {
                        'success': True,
                        'result': result,
                        'final_code': code_to_execute,
                        'revisions_needed': revision_num
                    }
                
                # Error occurred, need revision
                if revision_num < self.max_revisions:
                    if self.debug:
                        print(f"[DEBUG] ⚠️  Stage 6 error detected, revising code (attempt {revision_num + 1}/{self.max_revisions})")
                    code_to_execute = self._revise_code(
                        code_to_execute, result, observation, plan, instruction, 
                        file_path, output_path
                    )
                else:
                    self.logger.info(f"❌ Max revisions reached")
                    if self.enable_timing:
                        stage_time = time.time() - stage_start
                        self.stage_timings['stage_6_execution'] = stage_time
                    return {
                        'success': False,
                        'result': result,
                        'final_code': code_to_execute,
                        'revisions_needed': revision_num
                    }
                    
            except Exception as e:
                error_msg = f"Exception during execution: {str(e)}"
                self.logger.info(f"[EXECUTION EXCEPTION]\n{error_msg}")
                
                if revision_num < self.max_revisions:
                    code_to_execute = self._revise_code(
                        code_to_execute, error_msg, observation, plan, instruction,
                        file_path, output_path
                    )
                else:
                    if self.enable_timing:
                        stage_time = time.time() - stage_start
                        self.stage_timings['stage_6_execution'] = stage_time
                    return {
                        'success': False,
                        'result': error_msg,
                        'final_code': code_to_execute,
                        'revisions_needed': revision_num
                    }
        
        if self.enable_timing:
            stage_time = time.time() - stage_start
            self.stage_timings['stage_6_execution'] = stage_time
        return {
            'success': False,
            'result': 'Unexpected execution flow',
            'final_code': code_to_execute,
            'revisions_needed': self.max_revisions
        }
    
    def _revise_code(self, current_code: str, error_output: str,
                    observation: Dict[str, Any], plan: Dict[str, Any],
                    instruction: str, file_path: str, output_path: str) -> str:
        """Internal method to revise code based on execution errors"""

        revision_prompt = build_stage6_revision_prompt(
            instruction=instruction,
            observation_result=observation['result'][:600],
            planning_result=plan['response'][:600],
            current_code=current_code,
            execution_error=error_output,
        )
        
        # Generate revised code (prompt logging disabled)
        messages = [
            observation['prompt'], observation['response'],
            plan['prompt'], plan['response'],
            revision_prompt
        ]
        response = get_llm_response(messages, self.opt)
        
        revised_code = extract_code(response)
        # Replace with current test case paths (in case LLM hardcoded old paths)
        revised_code = self._replace_hardcoded_paths(revised_code, file_path, output_path)
        # Log revised code (KEY INFO for debugging)
        self.logger.info(f"[REVISED CODE]\n{revised_code}")
        
        return revised_code
    
    # Helper methods for parsing LLM responses
    def _parse_requirements(self, response: str) -> Dict[str, str]:
        """Extract structured requirements from understanding response"""
        sections = {}
        current_section = None
        current_content = []
        
        for line in response.split('\n'):
            if line.startswith('##'):
                if current_section:
                    sections[current_section] = '\n'.join(current_content).strip()
                current_section = line.strip('# ').strip()
                current_content = []
            elif current_section:
                current_content.append(line)
        
        if current_section:
            sections[current_section] = '\n'.join(current_content).strip()
        
        return sections
    
    def _extract_steps(self, response: str) -> List[str]:
        """Extract step-by-step plan from planning response"""
        steps = []
        for line in response.split('\n'):
            if line.strip().startswith('###'):
                steps.append(line.strip('# ').strip())
        return steps
    
    def _extract_issues(self, response: str) -> List[str]:
        """Extract validation issues from validation response"""
        issues = []
        for line in response.split('\n'):
            if '✗' in line or '[✗]' in line or '- [ ]' in line:
                issues.append(line.strip())
        return issues
    
    def _replace_hardcoded_paths(self, code: str, file_path: str, output_path: str) -> str:
        """Replace hardcoded file paths in generated code with current test case paths"""
        import re
        
        # Pattern 1: openpyxl.load_workbook('...')
        code = re.sub(
            r"openpyxl\.load_workbook\(['\"]([^'\"]+)['\"]\)",
            f"openpyxl.load_workbook('{file_path}')",
            code
        )
        
        # Pattern 2: wb.save('...')
        code = re.sub(
            r"wb\.save\(['\"]([^'\"]+)['\"]\)",
            f"wb.save('{output_path}')",
            code
        )
        
        # Pattern 3: Variable assignment like input_path = '...'
        code = re.sub(
            r"input_path\s*=\s*['\"]([^'\"]+)['\"]",
            f"input_path = '{file_path}'",
            code
        )
        code = re.sub(
            r"output_path\s*=\s*['\"]([^'\"]+)['\"]",
            f"output_path = '{output_path}'",
            code
        )
        
        return code
    
    def solve_task(self, task_data: Dict[str, Any], dataset_path: str) -> Dict[str, Any]:
        """
        Main pipeline: Process one task through all stages
        
        Enhanced 6-stage pipeline specifically for SpreadsheetBench:
        1. Deep Observation - Understand non-standard structure
        2. Instruction Understanding - Parse complex natural language
        3. Solution Planning - Design robust approach
        4. Code Implementation - Generate Python code
        5. Code Validation - Static analysis
        6. Execution & Revision - Run with retry mechanism
        """
        task_id = task_data['id']
        self.logger.info(f"\n{'='*60}\n🚀 Task {task_id}\n{'='*60}")
        
        if self.debug:
            print(f"\n{'='*80}")
            print(f"[DEBUG] 🎯 Starting Task {task_id}: {task_data['instruction'][:80]}...")
            print(f"{'='*80}")
        
        import time
        task_start = time.time() if self.enable_timing else None

        all_case_results = []
        aggregate_conversation = []
        total_revisions = 0
        overall_success = True
        final_code_last = None
        
        # Reset stage timings for this task
        self.stage_timings = {}
        
        # Shared code generation results (only generate once for test case 1)
        shared_understanding = None
        shared_plan = None
        shared_implementation = None
        shared_validation = None

        for test_case_idx in range(1, 4):
            file_name = f"{test_case_idx}_{task_data['spreadsheet_path'].lstrip('spreadsheet/')}_input.xlsx"
            input_path = f"/mnt/data/{self.opt.dataset}/{task_data['spreadsheet_path']}/{file_name}"
            output_path = f"/mnt/data/{self.opt.dataset}/outputs/sheetcopilot_{self.opt.model}/{file_name.replace('_input.xlsx', '_output.xlsx')}"
            if self.debug:
                print(f"\n[DEBUG] 📝 Processing test case {test_case_idx}/3")
            # Processing test case (logging reduced)

            conversation = []
            # Reset stage history per test case (must be list for append)
            self.stage_history = []
            try:
                # Stage 1: Deep Observation (always run - each test case has different input file)
                observation = self.stage_1_deep_observation(
                    input_path,
                    task_data['instruction'],
                    task_data['answer_position'],
                    task_data['instruction_type']
                )
                # Only add prompt summary to conversation, not full response (too verbose)
                conversation.append(f"[Stage 1 Observation completed for {input_path}]")
                if not observation['success']:
                    self.logger.info(f"⚠️ Observation failed: test {test_case_idx}")
                    all_case_results.append({
                        'test_case_index': test_case_idx,
                        'success': False,
                        'revisions_needed': 0,
                        'final_code': ''
                    })
                    overall_success = False
                    aggregate_conversation.extend(conversation)
                    continue

                # Stage 2-5: Only generate for test case 1, reuse for 2 and 3
                if test_case_idx == 1:
                    # Generate code with LLM for test case 1
                    
                    try:
                        # Stage 2: Instruction Understanding
                        shared_understanding = self.stage_2_instruction_understanding(
                            observation,
                            task_data['instruction'],
                            task_data['instruction_type']
                        )
                        conversation.append("[Stage 2 Understanding completed]")

                        # Stage 3: Solution Planning
                        shared_plan = self.stage_3_solution_planning(
                            observation, shared_understanding,
                            input_path, output_path,
                            task_data['answer_position']
                        )
                        conversation.append("[Stage 3 Planning completed]")

                        # Stage 4: Code Implementation
                        shared_implementation = self.stage_4_code_implementation(
                            observation, shared_understanding, shared_plan,
                            input_path, output_path,
                            task_data['answer_position']
                        )
                        conversation.append(f"[Stage 4 Implementation completed - Code length: {len(shared_implementation['code'])} chars]")

                        # Stage 5: Code Validation (with execution and result verification)
                        shared_validation = self.stage_5_code_validation(
                            shared_implementation, observation, shared_plan,
                            input_path, output_path,
                            task_data['instruction'], task_data['answer_position']
                        )
                        conversation.append(f"[Stage 5 Validation: {'PASSED' if shared_validation['passed'] else 'ISSUES FOUND'}]")
                    
                    except Exception as llm_error:
                        self.logger.error(f"❌ LLM stages (2-5) failed for test case 1: {str(llm_error)}")
                        # Mark all test cases as failed since we can't generate code
                        for tc_idx in range(1, 4):
                            all_case_results.append({
                                'test_case_index': tc_idx,
                                'success': False,
                                'revisions_needed': 0,
                                'final_code': ''
                            })
                        overall_success = False
                        aggregate_conversation.extend(conversation)
                        break  # Exit the test case loop
                else:
                    # Reuse code from test case 1
                    conversation.append(f"[Reusing code generated for test case 1]")

                # Stage 6: Execution & Revision (always run - different input/output paths)
                execution = self.stage_6_execution_and_revision(
                    shared_validation, shared_implementation, observation, shared_plan,
                    input_path, output_path, task_data['instruction']
                )
                conversation.append(execution['result'])
                self.logger.info(f"Test {test_case_idx}: {'✅' if execution['success'] else '❌'} (revisions: {execution['revisions_needed']})")
                final_code_last = execution['final_code']
                total_revisions += execution['revisions_needed']
                
                # Update shared code with revised version for next test case
                if execution['success'] and execution['final_code']:
                    # Store the successful code, but remove test-case-specific paths
                    # Next test case will have paths replaced via _replace_hardcoded_paths
                    shared_implementation['code'] = execution['final_code']
                    if shared_validation is not None and shared_validation.get('corrected_code'):
                        shared_validation['corrected_code'] = execution['final_code']
                if not execution['success']:
                    overall_success = False
                all_case_results.append({
                    'test_case_index': test_case_idx,
                    'success': execution['success'],
                    'revisions_needed': execution['revisions_needed'],
                    'final_code': execution['final_code'] or ''
                })
            except Exception as e:
                self.logger.info(f"❌ Exception test {test_case_idx}: {str(e)}")
                import traceback
                traceback.print_exc()
                overall_success = False
                all_case_results.append({
                    'test_case_index': test_case_idx,
                    'success': False,
                    'revisions_needed': 0,
                    'final_code': ''
                })
            aggregate_conversation.extend(conversation)

        if self.enable_timing:
            task_time = time.time() - task_start
            self.logger.info(f"Task time: {task_time:.2f}s")
            if self.debug:
                print(f"\n[DEBUG] ⏱️  Task {task_id} total time: {task_time:.2f}s")
                print(f"[DEBUG] Stage timings:")
                for stage_name, stage_time in self.stage_timings.items():
                    print(f"[DEBUG]   - {stage_name}: {stage_time:.2f}s")
        
        self.logger.info(f"Task {task_id}: {'✅ SUCCESS' if overall_success else '❌ FAILED'}")
        if self.debug:
            print(f"[DEBUG] {'✅ Task SUCCEEDED' if overall_success else '❌ Task FAILED'}\n{'='*60}")
        # Return combined result; final_code uses last successful code if any
        combined_result = self._create_result(task_data, aggregate_conversation, final_code_last, overall_success, total_revisions)
        combined_result['per_test_case'] = all_case_results
        if self.enable_timing:
            combined_result['timing'] = {
                'total_time': task_time,
                'stage_timings': self.stage_timings.copy()
            }
        return combined_result
    
    def _create_result(self, task_data: Dict[str, Any], conversation: List[str],
                      final_code: Optional[str], success: bool, 
                      revisions: int) -> Dict[str, Any]:
        """Create standardized result dictionary"""
        return {
            'id': task_data['id'],
            'instruction_type': task_data['instruction_type'],
            'conversation': conversation,
            'solution': final_code or '',
            'success': success,
            'revisions_needed': revisions,
            'stage_history': self.stage_history.copy()
        }


def parse_option():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser("SheetCopilot v2 - Enhanced for Real-World Spreadsheets")
    
    parser.add_argument('--model', type=str, required=True, help='LLM model name')
    parser.add_argument('--api_key', type=str, default="", help='API key for model')
    parser.add_argument('--base_url', type=str, default="", help='Base URL for model API')
    parser.add_argument('--dataset', type=str, default="test1", help='Dataset name')
    parser.add_argument('--code_exec_url', type=str, 
                       default="http://localhost:8080/execute", 
                       help='Code execution Docker URL')
    parser.add_argument('--conv_id', type=str, default="EVAL", 
                       help='Conversation ID for code execution')
    parser.add_argument('--max_revisions', type=int, default=3,
                       help='Maximum number of code revision attempts')
    parser.add_argument('--log_dir', type=str, default='../log',
                       help='Directory for log files')
    parser.add_argument('--enable_timing', action='store_true', default=True,
                       help='Enable stage timing for performance analysis')
    parser.add_argument('--disable_timing', dest='enable_timing', action='store_false',
                       help='Disable stage timing')
    parser.add_argument('--debug', action='store_true', default=True,
                       help='Enable debug output showing stage progress and timing in terminal')
    parser.add_argument('--excel-recalc', action='store_true', default=False,
                       help='After Docker execution, open workbook in real Excel to force full recalc (dynamic arrays spill)')
    parser.add_argument('--materialize-dynamic', action='store_true', default=False,
                       help='With --excel-recalc: convert dynamic spill ranges to static values for openpyxl compatibility')
    parser.add_argument('--strip-dynamic-formula', action='store_true', default=False,
                       help='With --materialize-dynamic: replace source spill formula cell with its calculated value')
    
    return parser.parse_args()


def main():
    """Main execution function"""
    opt = parse_option()
    print(f"SheetCopilot v2 Configuration:\n{opt}\n")
    
    # Setup logging
    logger = setup_logger(opt.log_dir, opt.model, opt.dataset)
    logger.info(f"Starting SheetCopilot v2 with config: {opt}")
    
    # Load dataset
    dataset_path = os.path.abspath(f'../data/{opt.dataset}')
    with open(f'{dataset_path}/dataset.json', 'r') as fp:
        dataset = json.load(fp)
    
    logger.info(f"Loaded {len(dataset)} tasks from {opt.dataset}")
    
    # Prepare output directories
    output_dir = f'{dataset_path}/outputs/sheetcopilot_{opt.model}'
    os.makedirs(output_dir, exist_ok=True)
    os.chmod(output_dir, 0o777)
    
    conv_file = f'{dataset_path}/outputs/conv_sheetcopilot_{opt.model}.jsonl'
    
    # Initialize copilot
    copilot = SheetCopilotV2(opt, logger)
    
    # Process tasks
    success_count = 0
    total_revisions = 0
    
    for task_data in tqdm(dataset, desc="Processing tasks"):
        result = copilot.solve_task(task_data, dataset_path)
        
        # Save conversation
        with open(conv_file, 'a+', encoding='utf-8') as fp:
            fp.write(json.dumps(result, ensure_ascii=False) + '\n')
        
        if result['success']:
            success_count += 1
        total_revisions += result['revisions_needed']
    
    # Summary
    success_rate = success_count / len(dataset) * 100
    avg_revisions = total_revisions / len(dataset)
    
    logger.info(f"\n{'='*100}")
    logger.info(f"FINAL RESULTS:")
    logger.info(f"Total tasks: {len(dataset)}")
    logger.info(f"Successful: {success_count}/{len(dataset)} ({success_rate:.1f}%)")
    logger.info(f"Average revisions: {avg_revisions:.2f}")
    logger.info(f"='*100")
    
    print(f"\n✅ SheetCopilot v2 completed!")
    print(f"Success rate: {success_rate:.1f}%")
    print(f"Results saved to: {conv_file}")


if __name__ == '__main__':
    main()
