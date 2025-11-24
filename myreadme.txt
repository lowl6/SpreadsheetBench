# 使用清华镜像安装 pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple 
conda create -n spreadsheetbench python=3.11

wsl
conda activate ssb

cd inference&&.\scripts\sheetcopilot_v2.ps1

conda activate ssb
cd code_exec_docker&&bash start_jupyter_server.sh 8080

cd evaluation&&.\scripts\evaluation.ps1
