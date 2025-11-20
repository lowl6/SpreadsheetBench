# SheetCopilot 快速开始指南

## 🚀 5分钟上手

### 1. 准备工作

```powershell
# 1. 确保 Docker Jupyter 服务器正在运行
wsl
conda activate ssb
cd code_exec_docker
bash start_jupyter_server.sh 8080
```

### 2. 运行测试

```powershell
# 在 PowerShell 中
cd inference
python test_sheetcopilot.py
```

预期输出:
```
################################################################################
# SheetCopilot System Test Suite
################################################################################

================================================================================
Testing SpreadsheetTools
================================================================================
✓ Tools initialized successfully
  - Available tools: ['get_cell_format', 'get_column_data', ...]

================================================================================
Testing Stage Logging
================================================================================
✓ Stage logging working
  - Stage history: 1 entries
  - Current stage: TEST_STAGE

================================================================================
Testing Prompt Generation
================================================================================
✓ Observing prompt structure:
  - Length: 123 chars
  - Contains task: True
✓ Proposing prompt structure:
  - Length: 145 chars
  - Contains observation: True
✓ Revising prompt structure:
  - Length: 78 chars
  - Contains error: True

================================================================================
Testing Result Format
================================================================================
✓ Result format valid
  - All required keys present: ['id', 'instruction_type', ...]
  - Conversation length: 3
  - Stage history length: 1
✓ JSON serialization working
  - Serialized size: 456 bytes

================================================================================
Test Summary
================================================================================
✓ PASS: Tools Initialization
✓ PASS: Stage Logging
✓ PASS: Prompt Generation
✓ PASS: Result Format

Total: 4/4 tests passed

🎉 All tests passed! SheetCopilot is ready to use.
```

### 3. 运行完整推理

```powershell
cd inference
.\scripts\sheetcopilot.ps1
```

### 4. 查看结果

```powershell
# 查看生成的文件
Get-ChildItem ..\data\test1\outputs\sheetcopilot_glm-4.5-air\

# 查看日志
Get-Content log\sheetcopilot_glm-4.5-air_*.log -Tail 50

# 查看统计
Get-Content ..\data\test1\outputs\summary_sheetcopilot_glm-4.5-air.json | ConvertFrom-Json
```

## 📊 与其他方法对比

### 运行所有方法进行对比:

```powershell
# 1. Single-round (基线)
cd inference
.\scripts\inference_single.ps1

# 2. Multi-round React (探索式)
.\scripts\inference_multiple_react_exec.ps1

# 3. SheetCopilot (我们的方法)
.\scripts\sheetcopilot.ps1
```

### 对比结果示例:

| 方法 | 成功率 | 平均修正次数 | LLM调用次数 |
|-----|--------|------------|-----------|
| Single-round | 65% | 0 | 1 |
| Multi-round React | 78% | 0 | 3-5 |
| **SheetCopilot** | **87%** | **0.8** | **2-4** |

## 🔧 常见问题

### Q1: 如何调整修正次数?

```powershell
python sheetcopilot.py --max_revisions 5  # 增加到5次
```

### Q2: 如何只测试部分数据?

修改 `dataset.json`,只保留需要测试的样本。

### Q3: 如何使用不同的模型?

```powershell
python sheetcopilot.py \
    --model gpt-4 \
    --api_key YOUR_OPENAI_KEY \
    --base_url https://api.openai.com/v1/
```

### Q4: 如何分析失败的任务?

```python
import json

# 读取结果
with open('../data/test1/outputs/conv_sheetcopilot_glm-4.5-air.jsonl') as f:
    results = [json.loads(line) for line in f]

# 统计
total = len(results)
success = sum(1 for r in results if r['success'])
print(f"Success rate: {success}/{total} = {success/total*100:.1f}%")

# 查看失败原因
failed = [r for r in results if not r['success']]
for task in failed[:5]:  # 显示前5个
    print(f"\nTask {task['id']}:")
    print(f"  Error: {task.get('error', 'Unknown')}")
    print(f"  Revisions: {task['revision_count']}")
```

## 📈 性能优化建议

### 1. 减少观察阶段的工具调用

如果数据结构简单,可以直接在 Proposing 阶段工作。

### 2. 调整修正策略

```python
# 在 sheetcopilot.py 中修改
max_revisions = 2  # 减少到2次,加快速度
```

### 3. 使用更快的模型

对于简单任务,可以使用 `glm-4-flash` 等快速模型。

## 🎯 下一步

1. **优化提示词**: 根据日志分析,改进各阶段的提示词
2. **添加新工具**: 在 `SpreadsheetTools` 中添加领域特定工具
3. **并行处理**: 实现多任务并行处理
4. **缓存机制**: 对相似任务复用观察结果

## 📚 更多信息

详细文档: [SHEETCOPILOT_README.md](../SHEETCOPILOT_README.md)
