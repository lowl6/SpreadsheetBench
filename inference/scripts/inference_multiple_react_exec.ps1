# PowerShell script for running inference_multiple.py with react_exec setting
# react_exec: ReAct mode without file content preview - LLM explores data by itself

python inference_multiple.py `
    --setting react_exec `
    --model glm-4.5-air `
    --api_key a3965f9fb7c14f6f8fac15bd076ee71b.omaVemiLXaga5JXg `
    --base_url https://open.bigmodel.cn/api/paas/v4/ `
    --dataset test1 `
    --max_turn_num 5 `
    --code_exec_url http://localhost:8080/execute `
    --conv_id EVAL `
    --row 5
