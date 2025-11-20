# PowerShell script for running SheetCopilot
# SheetCopilot: Multi-stage reasoning system (Observing → Proposing → Revising → Executing)

python sheetcopilot.py `
    --model glm-4.5-air `
    --api_key a3965f9fb7c14f6f8fac15bd076ee71b.omaVemiLXaga5JXg `
    --base_url https://open.bigmodel.cn/api/paas/v4/ `
    --dataset test1 `
    --code_exec_url http://localhost:8080/execute `
    --conv_id COPILOT `
    --max_revisions 3

# Note: Logs will be saved to inference/log/sheetcopilot_<model>_<timestamp>.log
# Results will be saved to data/<dataset>/outputs/sheetcopilot_<model>/
