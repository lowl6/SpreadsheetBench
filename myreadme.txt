# 使用清华镜像安装 pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple 
conda create -n spreadsheetbench python=3.11

wsl
conda activate ssb
cd inference&bash scripts/inference_single.sh


cd code_exec_docker 
bash start_jupyter_server.sh 8080