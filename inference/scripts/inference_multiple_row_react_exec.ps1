# PowerShell script for running inference_multiple.py with row_react_exec setting
# row_react_exec: ReAct mode with file content preview - hybrid approach

python inference_multiple.py `
    --setting row_react_exec `
    --model glm-4.5-air `
    --api_key a3965f9fb7c14f6f8fac15bd076ee71b.omaVemiLXaga5JXg `
    --base_url https://open.bigmodel.cn/api/paas/v4/ `
    --dataset test1 `
    --max_turn_num 5 `
    --code_exec_url http://localhost:8080/execute `
    --conv_id EVAL `
    --row 5
