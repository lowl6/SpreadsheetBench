# ID 325-44 Failure Root Cause Analysis

## 问题总结
- **任务**: Split filter data from column into specific columns (SQL-like filter parsing with Cartesian product expansion)
- **结果**: 0/3 test cases passed, 36-43 cell mismatches
- **根本原因**: 生成的代码使用简单分隔符拆分而非正则表达式,没有处理IN LIST的笛卡尔积展开

## 错误现象

### 预期输出 (Answer file)
```
Row 1: Headers (wtype_id, name, status, orgId, careType, specialtyId, contextId)
Row 2-10: 10 rows of split data
  - Row 2: 1 row (Input row 2, specialtyId=empty, contextId=1 value)
  - Row 3-8: 6 rows (Input row 3, specialtyId=2 × contextId=3 = 6 combinations)
  - Row 9-10: 2 rows (Input row 4, specialtyId=2 × contextId=1 = 2 combinations)
```

### 实际输出 (Generated file)
```
Row 1-6: 6 rows of data (部分正确但careType是"AMSTER"而非"careType")
Row 7-10: Input sheet原始数据(包含headers和filter列) ← 错误!
```

## 生成代码的3个关键错误

### 错误1: 使用简单分隔符而非正则表达式
```python
# ❌ 生成的错误代码
delimiters = ['|', ',', ';', ' ']
parts = str(filter_val).split(delimiter)
if len(parts) >= 4:  # 期望4个部分
    delimiter = delim

# 实际filter格式
("orgId" = "LIM") AND ("careType" = "AMSTER") AND ("contextId" IN LIST 98);
```
→ 根本无法用简单分隔符解析SQL-like格式!

**正确做法**:
```python
import re
# 提取单值字段
single_values = re.findall(r'\("([^"]+)"\s*=\s*"([^"]+)"\)', filter_val)
# 提取多值字段 (IN LIST)
multi_values = re.findall(r'\("([^"]+)"\s*IN\s*LIST\s*([^\)]+)\)', filter_val)
```

### 错误2: 没有处理笛卡尔积展开
```python
# ❌ 生成的错误代码
for row_idx in range(4):  # 只处理4行输入
    # ...简单split后创建1行输出
    output_data.append(output_row)

# 实际需求
Input Row 3: specialtyId IN LIST 66,77 × contextId IN LIST 55,689,213
→ 需要展开成 2 × 3 = 6 rows!
```

**正确做法**:
```python
from itertools import product

# 提取多值字段
specialty_values = ['66', '77']
context_values = ['55', '689', '213']

# 生成笛卡尔积
for specialty, context in product(specialty_values, context_values):
    output_row = [wtype_id, name, status, orgId, careType, specialty, context]
    output_data.append(output_row)
```

### 错误3: 只写入4行而非10行
```python
# ❌ 生成的错误代码
for row_idx in range(4):  # 固定4行
    for col_idx in range(7):
        ws.cell(row=row_idx + 2, column=col_idx + 1).value = output_data[row_idx][col_idx]
```
→ output_data应该有10个元素,但代码只写入了前4个!

**正确做法**:
```python
# 写入所有输出行
for row_idx, row_data in enumerate(output_data, start=2):  # 从row 2开始(row 1是header)
    for col_idx, value in enumerate(row_data, start=1):
        ws.cell(row=row_idx, column=col_idx).value = value
```

## 为什么会产生row 7-10的错误数据?

可能的原因:
1. **Output sheet初始化问题**: 代码可能从Input sheet复制了数据到Output,然后只更新了前几行
2. **Sheet处理逻辑**: 可能创建Output sheet时直接复制了Input的前N行,导致row 7-10残留原始数据

生成代码中的sheet处理:
```python
ws = wb['Output']  # 直接使用已存在的Output sheet
```
→ 没有清空已有数据!

**正确做法**:
```python
# 如果Output sheet已存在,先删除
if 'Output' in wb.sheetnames:
    del wb['Output']
# 创建新的Output sheet
ws = wb.create_sheet('Output')
```

或者:
```python
# 清空Output sheet的所有数据
ws = wb['Output']
for row in ws.iter_rows():
    for cell in row:
        cell.value = None
```

## Stage 2 Understanding阶段缺失的分析

当前Stage 2只有"LOOKUP OPERATIONS"指导,完全没有"FILTER SPLITTING"的指导,导致LLM:
1. 没有识别出filter是SQL-like格式
2. 没有计算需要多少行输出 (1+6+2+1=10)
3. 没有规划笛卡尔积生成逻辑

## 修复方案

### 1. 增强Stage 2 Prompt - 添加FILTER SPLITTING章节

在`STAGE2_UNDERSTANDING_PROMPT_TEMPLATE`的"LOOKUP OPERATIONS"章节后添加:

```python
🔍 **SPECIAL ATTENTION - FILTER SPLITTING OPERATIONS** (CRITICAL FOR DATA EXPANSION):
If instruction mentions "split filter", "parse filter", "extract from filter column":

⚠️ **STEP 1: IDENTIFY FILTER FORMAT**:
Examine sample filter values from observation to determine the pattern:
- SQL-like: ("key1" = "value1") AND ("key2" = "value2") AND ("key3" IN LIST a,b,c);
- Delimited: key1=value1|key2=value2|key3=a,b,c
- JSON-like: {"key1": "value1", "key2": "value2"}

📋 **STEP 2: UNDERSTAND EXPANSION REQUIREMENT**:
When filter contains "IN LIST" or multiple values:
  - **Single value field**: ("orgId" = "LIM") → 1 output row
  - **Multi-value field**: ("specialtyId" IN LIST 66,77) → 2 output rows
  - **Cartesian product**: specialtyId[66,77] × contextId[55,689,213] → 6 rows (2×3)
  
  → **CRITICAL**: Count expected output rows for EACH input row
  → Example: Input 4 rows → Output 10 rows (1+6+2+1 after expansion)

🎯 **STEP 3: DEFINE PARSING LOGIC**:
1. **Key extraction**: Use regex for SQL-like format
   - Single values: r'\("([^"]+)"\s*=\s*"([^"]+)"\)'
   - IN LIST: r'\("([^"]+)"\s*IN\s*LIST\s*([^\)]+)\)'
   
2. **Value splitting**: Handle comma-separated multi-values
   - Example: "66,77" → ["66", "77"] using split(',')
   
3. **Cartesian product**: Use itertools.product() for multiple IN LIST fields
   - Generate all combinations: (s1,c1), (s1,c2), (s1,c3), (s2,c1)...

⚠️ **COMMON MISTAKES TO AVOID**:
- Using simple split(delimiter) instead of regex
- Only processing N rows instead of calculating exact output count
- Not handling Cartesian product (assuming 1-to-1 mapping)
- Writing to wrong sheet or not clearing Output sheet first
```

### 2. 增强Stage 3 Planning - 添加多行展开计划模板

在`STAGE3_PLANNING_PROMPT_TEMPLATE`中添加针对filter splitting的特殊步骤:

```python
### Step 3.5: Data Expansion Planning (for filter splitting tasks)
If task requires splitting filter column with IN LIST:
- Calculate expected output row count per input row
- Plan Cartesian product generation using itertools.product()
- Ensure Output sheet is cleared before writing
- Write ALL expanded rows (not just original row count)

Example:
  Input Row 1: orgId="A", specialtyId=empty, contextId="1" → 1 output row
  Input Row 2: orgId="B", specialtyId IN LIST 66,77, contextId IN LIST 55,689 → 4 output rows (2×2)
  TOTAL: 5 output rows (not 2!)
```

### 3. 增强Stage 5 Validation - 检查行数匹配

在`STAGE5_VALIDATION_SUCCESS_TEMPLATE`中添加:

```python
4. **Row Count Validation** (for data transformation tasks):
   - Compare input row count vs output row count
   - For filter splitting with Cartesian product:
     * Output should have MORE rows than input (expansion)
     * If output_rows == input_rows, likely missing expansion logic
   - Suspicious patterns:
     * INPUT_ROW_MISMATCH: Output row count doesn't match expected expansion
     * MISSING_DATA: Some input rows not processed
```

### 4. 创建测试文件验证修复

创建`test_325_44_fix.py`验证:
1. 正则表达式能否正确解析filter
2. 笛卡尔积计算是否正确 (1+6+2+1=10)
3. 生成的10行数据是否符合答案格式

## 影响范围评估

这个修复可能影响所有涉及以下特征的任务:
- **Sheet-Level Manipulation**: 大规模数据转换任务
- **Complex filter parsing**: SQL-like, JSON-like格式的文本解析
- **Data expansion**: 一行输入需要展开成多行输出的场景 (如pivot展开)

建议在修复后重新运行全部test1数据集(10个任务),检查是否有regression。

## 下一步行动

1. ✅ 分析完成 - 找到3个核心错误
2. ⬜ 修改`stage_prompts.py` - 添加FILTER SPLITTING章节
3. ⬜ 创建`test_325_44_fix.py` - 验证修复逻辑
4. ⬜ 重新运行ID 325-44 inference
5. ⬜ 运行evaluation验证是否通过
6. ⬜ 检查其他任务是否有regression

---
生成时间: 2024-11 | 文档版本: v1.0
