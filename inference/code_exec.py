from jupyter_kernel_cli import ClientJupyterKernel
import re

def get_exec_client(url, conv_id):
    client = ClientJupyterKernel(url, conv_id)
    return client

def extract_code(response):
    # Prefer ```python fenced blocks
    if '```python' in response:
        segment = response.split('```python', 1)[1]
        if '```' in segment:
            code = segment.split('```', 1)[0]
            return code.strip('\n')
    # Generic triple backticks fallback
    generic = re.search(r"```(.*?)```", response, re.DOTALL)
    if generic:
        inner = generic.group(1)
        # Remove possible leading language token
        inner = inner.split('\n', 1)[-1] if '\n' in inner else inner
        return inner.strip('\n')
    # Fallback: treat whole response as code (may be invalid)
    return response

def exec_code(client, code):
    res = client.execute(code)
    if '-----' in res:
        tracebacks = res.split('\n\n\n\n')
        error_feedback = ''
        for t in tracebacks:
            if 'Error' in t:
                error_feedback += t + '\n'
                break
        for t in tracebacks:
            if t.startswith('Cell'):
                error_feedback += t
                break
        error_feedback += tracebacks[-1]
        return error_feedback
    return res
