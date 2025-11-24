"""stage_prompts.py
集中管理 SheetCopilot v2 六个阶段的提示词模板。

设计目的：
1. 统一维护，方便后续优化与版本迭代（例如针对模型差异做适配）。
2. 避免在核心管线代码中混杂超长 f-string，提升可读性与可维护性。
3. 每个阶段提供清晰中文注释，说明其意图、输入、输出关注点。

使用方式：
在 `sheetcopilot_v2.py` 中导入后，调用对应的 `build_...` 函数/模板，将已截断的上下文内容传入。

占位符命名规范： `{instruction}`, `{instruction_type}`, `{answer_position}`, `{file_path}`,
`{output_path}`, `{observation_result}`, `{understanding_result}`, `{planning_result}`,
`{implementation_plan}` 等。调用方应保证传入的文本已做长度截断（避免 prompt 过长）。
"""

# =========================
# Stage 1: 观察阶段总结模板
# 说明：Stage1 实际不向 LLM 生成代码，仅运行预定义的观察脚本。这里仅定义其“总结上下文”模板，
#       供后续阶段拼接到消息列表里。
# =========================
STAGE1_OBSERVATION_SUMMARY_TEMPLATE = """📊 SPREADSHEET OBSERVATION COMPLETED\n\n🎯 Task: {instruction}\n📋 Type: {instruction_type}\n🎯 Target: {answer_position}\n📂 File: {file_path}\n\nObservation Results:\n{observation_result}\n"""

# =========================
# Stage 2: 指令理解阶段 (Instruction Understanding)
# 目标：将真实论坛的自然语言指令解析为结构化需求（核心目标 / 输入位置 / 输出格式 / 业务逻辑）。
# 输入：原始指令 + 观察阶段摘要 + 指令类型。
# 输出：结构化分段文本，后续用于规划阶段。要求模型避免臆测不存在的数据。
# 关键点：识别是否存在“ANSWER POSITION CURRENT CONTENT”作为格式参考。
# =========================
STAGE2_UNDERSTANDING_PROMPT_TEMPLATE = """You are SheetCopilot v2 in INSTRUCTION UNDERSTANDING stage.\n\nThis is a REAL-WORLD user question from Excel forums. Your task is to extract the CORE requirements.\n\n📝 **ORIGINAL INSTRUCTION** (may be long and informal):\n{instruction}\n\n📊 **SPREADSHEET STRUCTURE** (from observation):\n{observation_result}\n\n🎯 **TASK TYPE**: {instruction_type}\n\n💡 **IMPORTANT**: Check if \"ANSWER POSITION CURRENT CONTENT\" section shows existing data - if yes, this is a FORMAT REFERENCE showing the expected output format (data type, number format, formula style, etc.). Your solution MUST preserve this format.\n\n**YOUR ANALYSIS TASK**:\nBreak down this real-world instruction into structured requirements:\n\n## 1. Core Objective\nWhat is the PRIMARY goal? (in one clear sentence)\n\n## 2. Input Data Location\n- Which cells/ranges contain the INPUT data?\n- Are there multiple source locations?\n- What format is the input data? (numbers, text, formulas, etc.)\n\n## 3. Output Requirements\n- Where should results be written? (target cells)\n- What format should output be? (formula, value, formatting, etc.)\n- Any specific output constraints?\n\n## 4. Business Logic\n- What calculation/operation is needed?\n- Any conditions or criteria to apply?\n- Special cases or edge cases mentioned?\n\nProvide your structured analysis:\n"""

# =========================
# Stage 3: 解决方案规划阶段 (Solution Planning)
# 目标：输出分步骤计划，强调动态引用、格式保留、空单元格处理、边界与风险规避。
# 输入：观察摘要 + 指令理解结果 + 路径信息。
# 输出：计划步骤列表与风险缓解策略。
# =========================
STAGE3_PLANNING_PROMPT_TEMPLATE = """You are SheetCopilot v2 in SOLUTION PLANNING stage.\n\n📊 **SPREADSHEET FACTS** (non-standard structure):\n{observation_result}\n\n🎯 **UNDERSTOOD REQUIREMENTS**:\n{understanding_result}\n\n📂 **FILE PATHS**:\n- Input: {file_path}\n- Output: {output_path}\n- Target cells: {answer_position}\n\n💡 **FORMAT REFERENCE**: If observation shows existing data in answer position, PRESERVE that format (data type, number format, formula vs value). This is critical for correctness!\n\n**YOUR PLANNING TASK**:\nDesign a step-by-step implementation plan that handles NON-STANDARD spreadsheet formats.\n\n## Implementation Plan Template:\n\n### Step 1: Load and Validate\n- Load workbook from {file_path}\n- Identify target sheet (handle multi-sheet case)\n- Validate target range {answer_position} exists\n- Check for merged cells or formatting in target area\n\n### Step 2: Locate Input Data (DYNAMIC, not hardcoded!)\n- Based on observation, input data is at: [SPECIFY ACTUAL LOCATION]\n- NOT assuming A1 start!\n- Handle empty cells: [STRATEGY]\n- Account for non-standard table boundaries\n\n### Step 3: Extract and Process\n- Read input data using dynamic references\n- Data type conversions needed: [SPECIFY]\n- Handle edge cases: empty cells, merged cells, formulas vs values\n- Validation checks before processing\n\n### Step 4: Apply Business Logic\n- Core operation: [DESCRIBE CLEARLY]\n- Formula structure (if applicable): [FORMULA]\n- Calculation steps: [ENUMERATE]\n- Condition handling: [IF ANY]\n\n### Step 5: Write Results\n- Target cells: {answer_position}\n- Write as: [FORMULA or VALUE or FORMATTED_VALUE]\n- Preserve existing formatting: [YES/NO]\n- Handle multiple target cells: [STRATEGY]\n\n### Step 6: Save and Verify\n- Save to {output_path}\n- Verify write succeeded\n- Close workbook properly\n\n## Risk Mitigation:\n- ❌ AVOID: Hardcoding cell references like A1, B2\n- ✅ USE: Dynamic references based on observation results\n- ❌ AVOID: Assuming headers in row 1\n- ✅ USE: Actual header locations from analysis\n- ❌ AVOID: Ignoring empty cells\n- ✅ USE: Explicit null/empty checks\n\nProvide your COMPLETE plan with SPECIFIC cell references based on the observation:\n"""

# =========================
# Stage 4: 代码实现阶段 (Code Implementation)
# 目标：基于规划与观察结果生成鲁棒 Python + openpyxl 代码，强调：动态定位、格式保留、避免硬编码、避免循环引用。
# 输入：观察、理解、规划结果 + 路径信息 + 目标区域。
# 输出：完整代码（含错误处理）。
# 关键：公式写入必须写入字符串形式的公式，不要写入计算后结果（除非要求值）。
# =========================
STAGE4_IMPLEMENTATION_PROMPT_TEMPLATE = """You are SheetCopilot v2 in CODE IMPLEMENTATION stage.\n\n📊 **OBSERVED STRUCTURE**:\n{observation_result}\n\n🎯 **REQUIREMENTS SUMMARY**:\n{understanding_result}\n\n📋 **IMPLEMENTATION PLAN**:\n{planning_result}\n\n**YOUR CODING TASK**:\nWrite COMPLETE, PRODUCTION-READY Python code following the plan above.\n\n**🎯 FORMAT & DATA TYPE PRESERVATION (CRITICAL)**:\nRefer to any \"ANSWER POSITION CURRENT CONTENT\" block in observation: replicate formula vs value pattern EXACTLY.\n\n⚠️ Avoid circular references; do NOT reference target cells inside formulas for those same cells.\n\n🚫 Structural Prohibitions:\n- Do NOT create helper columns only to delete them.\n- Do NOT delete columns unless explicitly required.\n- Prefer reading original value into Python variable if needed.\n\nFormula Syntax Reminders:\n- No leading @ before function/sheet names.\n- Concatenate strings with & outside quotes: ="*"&A1&"*"\n\nPaths:\n- Input workbook: {file_path}\n- Output workbook: {output_path}\n- Target range: {answer_position}\n\nGenerate the full implementation now (with try/except, dynamic references, null checks):\n"""

# =========================
# Stage 5: 验证阶段 (Validation - Execute & Verify)
# 目标：执行生成代码，读取输出 answer_position 内容，与输入格式模式对比，判断是否匹配语义/数据类型要求。
# 输入：执行结果 + 输入/输出的抽取内容与统计摘要。
# 输出：两种可能：PASSED 或 FAILED（含修复代码）。
# 提示模板包含决策说明与修复结构化输出格式。
# =========================
STAGE5_VALIDATION_FAILURE_TEMPLATE = """You are SheetCopilot v2 in CODE VALIDATION stage.\n\nThe code execution FAILED. Please identify and fix the errors.\n\n📋 **TASK**: {instruction}\n\n📊 **OBSERVED DATA (truncated)**:\n{observation_result}\n\n📋 **IMPLEMENTATION PLAN (truncated)**:\n{planning_result}\n\n💻 **GENERATED CODE (has errors)**:\n```python\n{generated_code}\n```\n\n❌ **EXECUTION ERROR**:\n```\n{execution_error}\n```\n\n**YOUR TASK**:\n1. Root cause analysis (traceback).\n2. Provide CORRECTED code (entire script).\n\nCORRECTED CODE:\n"""

STAGE5_VALIDATION_SUCCESS_TEMPLATE = """You are SheetCopilot v2 in CODE VALIDATION stage.\n\nThe code executed SUCCESSFULLY. Evaluate the semantic correctness of results.\n\n📋 **ORIGINAL TASK**: {instruction}\n\n📊 **OBSERVED INPUT (truncated)**:\n{observation_result}\n\n📋 **IMPLEMENTATION PLAN (truncated)**:\n{planning_result}\n\n💻 **EXECUTED CODE**:\n```python\n{generated_code}\n```\n\n✅ **RAW EXECUTION STDOUT**:\n```\n{execution_stdout}\n```\n\n🎯 **INPUT ANSWER COLUMN PATTERN (reference in {answer_position})**:\n```\n{input_answer_content}\n```\n📊 **INPUT ANSWER SUMMARY**:\n```json\n{input_summary_json}\n```\n\n📌 **OUTPUT RESULT CELLS (generated in {answer_position})**:\n```\n{output_answer_content}\n```\n📊 **OUTPUT RESULT SUMMARY**:\n```json\n{output_summary_json}\n```\n\n🛑 **NEIGHBOR COLUMN LEAK CHECK**:\n```json\n{neighbor_alert_json}\n```\n\nFollow validation instructions (pattern match: formula vs value, data types, casing). Return EXACTLY one of: VALIDATION PASSED block or VALIDATION FAILED block with corrected code.\n"""

# =========================
# Stage 6: 执行与修订阶段 (Execution & Revision)
# 目标：根据错误输出进行迭代修复，直到成功或达到最大次数。此处仅抽离修复提示模板。
# 输入：当前代码 + 错误输出 + 观察 + 规划 + 指令。
# 输出：新的修复后代码。
# =========================
STAGE6_REVISION_PROMPT_TEMPLATE = """You are SheetCopilot v2 in ERROR RECOVERY mode.\n\n🎯 **TASK**: {instruction}\n\n📊 **SPREADSHEET STRUCTURE (observed)**:\n{observation_result}\n\n📋 **ORIGINAL PLAN (truncated)**:\n{planning_result}\n\n💻 **CURRENT CODE (has errors)**:\n```python\n{current_code}\n```\n\n❌ **EXECUTION ERROR**:\n{execution_error}\n\nDebug & fix root cause (not superficial patch). Typical issues: wrong range, None cell, sheet name mismatch, formula syntax (@ prefix / string concat), type conversion, circular reference. Provide COMPLETE corrected code only.\n\nCORRECTED CODE:\n"""

# =========================
# 构建型函数：方便后续灵活插入截断后的上下文
# =========================
def build_stage1_summary(instruction: str, instruction_type: str, answer_position: str, file_path: str, observation_result: str) -> str:
    return STAGE1_OBSERVATION_SUMMARY_TEMPLATE.format(
        instruction=instruction,
        instruction_type=instruction_type,
        answer_position=answer_position,
        file_path=file_path,
        observation_result=observation_result,
    )

def build_stage2_prompt(instruction: str, instruction_type: str, observation_result: str) -> str:
    return STAGE2_UNDERSTANDING_PROMPT_TEMPLATE.format(
        instruction=instruction,
        instruction_type=instruction_type,
        observation_result=observation_result,
    )

def build_stage3_prompt(observation_result: str, understanding_result: str, file_path: str, output_path: str, answer_position: str) -> str:
    return STAGE3_PLANNING_PROMPT_TEMPLATE.format(
        observation_result=observation_result,
        understanding_result=understanding_result,
        file_path=file_path,
        output_path=output_path,
        answer_position=answer_position,
    )

def build_stage4_prompt(observation_result: str, understanding_result: str, planning_result: str, file_path: str, output_path: str, answer_position: str) -> str:
    return STAGE4_IMPLEMENTATION_PROMPT_TEMPLATE.format(
        observation_result=observation_result,
        understanding_result=understanding_result,
        planning_result=planning_result,
        file_path=file_path,
        output_path=output_path,
        answer_position=answer_position,
    )

def build_stage5_failure_prompt(instruction: str, observation_result: str, planning_result: str, generated_code: str, execution_error: str) -> str:
    return STAGE5_VALIDATION_FAILURE_TEMPLATE.format(
        instruction=instruction,
        observation_result=observation_result,
        planning_result=planning_result,
        generated_code=generated_code,
        execution_error=execution_error,
    )

def build_stage5_success_prompt(instruction: str, observation_result: str, planning_result: str, generated_code: str,
                                execution_stdout: str, answer_position: str,
                                input_answer_content: str, input_summary_json: str,
                                output_answer_content: str, output_summary_json: str,
                                neighbor_alert_json: str) -> str:
    return STAGE5_VALIDATION_SUCCESS_TEMPLATE.format(
        instruction=instruction,
        observation_result=observation_result,
        planning_result=planning_result,
        generated_code=generated_code,
        execution_stdout=execution_stdout,
        answer_position=answer_position,
        input_answer_content=input_answer_content,
        input_summary_json=input_summary_json,
        output_answer_content=output_answer_content,
        output_summary_json=output_summary_json,
        neighbor_alert_json=neighbor_alert_json,
    )

def build_stage6_revision_prompt(instruction: str, observation_result: str, planning_result: str, current_code: str, execution_error: str) -> str:
    return STAGE6_REVISION_PROMPT_TEMPLATE.format(
        instruction=instruction,
        observation_result=observation_result,
        planning_result=planning_result,
        current_code=current_code,
        execution_error=execution_error,
    )
