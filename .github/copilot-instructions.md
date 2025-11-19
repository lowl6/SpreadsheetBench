# Copilot Instructions for SpreadsheetBench

## Project Overview
- **SpreadsheetBench** is a benchmark for real-world spreadsheet manipulation, featuring 912 real user questions and diverse spreadsheet files. It evaluates LLMs and agents on realistic, complex spreadsheet tasks.
- The project is organized into major components:
  - `data/`: All benchmark data, including JSON datasets and spreadsheet test cases.
  - `inference/`: Scripts and code for running LLM inference (single/multi-round, ReAct, execution feedback).
  - `evaluation/`: Evaluation scripts, including Windows-only tools using `win32com` for spreadsheet inspection.
  - `code_exec_docker/`: Docker-based code execution environment for running generated Python code securely.

## Key Workflows
- **Environment Setup:**
  - Install Python dependencies: `pip install -r requirements.txt`
  - For code execution, build Docker images in `code_exec_docker/`:
    - `docker build -t xingyaoww/codeact-execute-api -f Dockerfile.api .`
    - `docker build -t xingyaoww/codeact-executor -f Dockerfile.executor .`
  - Start the Jupyter server for code execution: `bash start_jupyter_server.sh PORT`
- **Inference:**
  - Single-round: `cd inference && bash scripts/inference_single.sh`
  - Multi-round: Use scripts in `inference/scripts/` (see README for details). Edit model/api_key/base_url in scripts as needed.
  - Outputs are saved in `inference/outputs/` and result spreadsheets in `data/sample_data_200/outputs/`.
- **Evaluation:**
  - Only supported on Windows (uses `win32com`). Run: `cd evaluation && bash scripts/evaluation.sh`
  - Edit settings and model in `evaluation/scripts/evaluation.sh`.

## Project Conventions & Patterns
- **Data Structure:**
  - Each data point: `{id, instruction, spreadsheet_path, instruction_type, answer_position}`
  - Test cases: Each question has multiple input/answer spreadsheet pairs in subfolders.
- **Execution Isolation:**
  - All generated code is executed in Docker containers for safety and reproducibility.
- **Cross-Component Communication:**
  - Inference scripts call the Docker API for code execution.
  - Evaluation scripts read outputs from inference and compare with ground truth spreadsheets.
- **Platform-Specific Notes:**
  - Evaluation requires Windows due to Excel automation.
  - Inference and code execution are cross-platform (Linux/Mac/Windows).

## Examples & References
- See `README.md` (root) for full workflow and troubleshooting.
- See `code_exec_docker/README.md` for Docker setup and config details.
- Key scripts: `inference/scripts/inference_single.sh`, `evaluation/scripts/evaluation.sh`, `code_exec_docker/start_jupyter_server.sh`

## Contribution & Maintenance
- All data and code are versioned. Contact maintainers for questions or contributions (see README for emails).

---
For any unclear workflow or missing convention, please ask for clarification or check the latest `README.md` files.
