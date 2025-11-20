# PowerShell script for SheetCopilot v2
# Enhanced version specifically designed for SpreadsheetBench's characteristics

Write-Host "🚀 Starting SheetCopilot v2..." -ForegroundColor Cyan
Write-Host "Enhanced for:" -ForegroundColor Yellow
Write-Host "  ✓ Complex real-world instructions" -ForegroundColor Green
Write-Host "  ✓ Non-standard table layouts" -ForegroundColor Green
Write-Host "  ✓ Multiple tables & sheets" -ForegroundColor Green
Write-Host "  ✓ Rich formatting & diverse formats" -ForegroundColor Green
Write-Host ""

python sheetcopilot_v2.py `
    --model glm-4.5-air `
    --api_key a3965f9fb7c14f6f8fac15bd076ee71b.omaVemiLXaga5JXg `
    --base_url https://open.bigmodel.cn/api/paas/v4/ `
    --dataset test1 `
    --code_exec_url http://localhost:8080/execute `
    --conv_id COPILOT `
    --log_dir "../log"

Write-Host ""
Write-Host "✅ SheetCopilot v2 execution completed!" -ForegroundColor Green
