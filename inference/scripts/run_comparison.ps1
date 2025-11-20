# Comparison Script: Run all methods and generate comparison report
# 对比 Single-round, Multi-round, 和 SheetCopilot 的性能

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "  SpreadsheetBench: Method Comparison Experiment" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""

$MODEL = "glm-4.5-air"
$DATASET = "test1"
$API_KEY = "a3965f9fb7c14f6f8fac15bd076ee71b.omaVemiLXaga5JXg"
$BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"

# Function to run a method and capture results
function Run-Method {
    param(
        [string]$MethodName,
        [string]$Script
    )
    
    Write-Host ""
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "  Running: $MethodName" -ForegroundColor Yellow
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    
    $StartTime = Get-Date
    
    & $Script
    
    $EndTime = Get-Date
    $Duration = ($EndTime - $StartTime).TotalMinutes
    
    Write-Host ""
    Write-Host "  ✓ $MethodName completed in $([math]::Round($Duration, 2)) minutes" -ForegroundColor Green
    Write-Host ""
    
    return @{
        Method = $MethodName
        Duration = $Duration
        StartTime = $StartTime
        EndTime = $EndTime
    }
}

# Array to store results
$Results = @()

# 1. Single-round baseline
Write-Host "Method 1/3: Single-round Inference (Baseline)" -ForegroundColor Cyan
$Results += Run-Method -MethodName "Single-round" -Script ".\scripts\inference_single.ps1"

# 2. Multi-round ReAct
Write-Host "Method 2/3: Multi-round ReAct Execution" -ForegroundColor Cyan
$Results += Run-Method -MethodName "Multi-round ReAct" -Script ".\scripts\inference_multiple_react_exec.ps1"

# 3. SheetCopilot (our method)
Write-Host "Method 3/3: SheetCopilot (Multi-stage Reasoning)" -ForegroundColor Cyan
$Results += Run-Method -MethodName "SheetCopilot" -Script ".\scripts\sheetcopilot.ps1"

# Generate comparison report
Write-Host ""
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "  Generating Comparison Report..." -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""

python -c @"
import json
import os
from datetime import datetime

# Load results from each method
dataset_path = f'../data/$DATASET/outputs'

methods = {
    'Single-round': f'{dataset_path}/conv_single_$MODEL.jsonl',
    'Multi-round ReAct': f'{dataset_path}/conv_multi_react_exec_$MODEL.jsonl',
    'SheetCopilot': f'{dataset_path}/conv_sheetcopilot_$MODEL.jsonl'
}

report = {
    'experiment_date': datetime.now().isoformat(),
    'model': '$MODEL',
    'dataset': '$DATASET',
    'methods': {}
}

for method_name, file_path in methods.items():
    if not os.path.exists(file_path):
        print(f'⚠️  Warning: {file_path} not found, skipping {method_name}')
        continue
    
    with open(file_path, 'r', encoding='utf-8') as f:
        results = [json.loads(line) for line in f if line.strip()]
    
    total = len(results)
    successful = sum(1 for r in results if r.get('success', False))
    failed = total - successful
    success_rate = (successful / total * 100) if total > 0 else 0
    
    # Count revisions (only for SheetCopilot)
    avg_revisions = 0
    if method_name == 'SheetCopilot':
        total_revisions = sum(r.get('revision_count', 0) for r in results)
        avg_revisions = total_revisions / total if total > 0 else 0
    
    # Count LLM calls
    avg_llm_calls = 0
    if method_name == 'Single-round':
        avg_llm_calls = 1
    elif method_name == 'Multi-round ReAct':
        # Estimate from conversation length
        avg_conv_len = sum(len(r.get('conversation', [])) for r in results) / total if total > 0 else 0
        avg_llm_calls = avg_conv_len / 2  # Each LLM call = prompt + response
    elif method_name == 'SheetCopilot':
        # Count from conversation
        avg_conv_len = sum(len(r.get('conversation', [])) for r in results) / total if total > 0 else 0
        avg_llm_calls = (avg_conv_len - 1) / 2  # -1 for initial result
    
    report['methods'][method_name] = {
        'total_tasks': total,
        'successful': successful,
        'failed': failed,
        'success_rate': round(success_rate, 2),
        'avg_revisions': round(avg_revisions, 2),
        'avg_llm_calls': round(avg_llm_calls, 2)
    }
    
    print(f'✓ Analyzed: {method_name}')
    print(f'  Total: {total}, Success: {successful}, Failed: {failed}')
    print(f'  Success Rate: {success_rate:.2f}%')
    if avg_revisions > 0:
        print(f'  Avg Revisions: {avg_revisions:.2f}')
    print(f'  Avg LLM Calls: {avg_llm_calls:.2f}')
    print()

# Save report
report_file = f'{dataset_path}/comparison_report_$MODEL.json'
with open(report_file, 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print(f'📊 Report saved to: {report_file}')

# Print comparison table
print('\n' + '='*80)
print('COMPARISON TABLE')
print('='*80)
print(f'{'Method':<25} {'Success Rate':<15} {'Avg Revisions':<15} {'Avg LLM Calls':<15}')
print('-'*80)

for method_name, stats in report['methods'].items():
    print(f'{method_name:<25} {stats['success_rate']:<14.2f}% {stats['avg_revisions']:<15.2f} {stats['avg_llm_calls']:<15.2f}')

print('='*80)
"@

Write-Host ""
Write-Host "================================================================================" -ForegroundColor Green
Write-Host "  ✓ All experiments completed!" -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Results summary:" -ForegroundColor Cyan
foreach ($result in $Results) {
    Write-Host "  - $($result.Method): $([math]::Round($result.Duration, 2)) minutes" -ForegroundColor White
}
Write-Host ""
Write-Host "Check the comparison report at: data/$DATASET/outputs/comparison_report_$MODEL.json" -ForegroundColor Yellow
Write-Host ""
