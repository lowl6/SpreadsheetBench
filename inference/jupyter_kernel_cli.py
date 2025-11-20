import json
import requests

class ClientJupyterKernel:
    def __init__(self, url, conv_id):
        self.url = url
        self.conv_id = conv_id
        print(f"ClientJupyterKernel initialized with url={url} and conv_id={conv_id}")

    def execute(self, code):
        payload = {"convid": self.conv_id, "code": code}
        try:
            response = requests.post(self.url, data=json.dumps(payload), timeout=30)
        except Exception as e:
            return f"EXECUTION REQUEST ERROR: {e}"
        raw_text = response.text
        try:
            response_data = response.json()
        except Exception as e:
            return f"JSON_DECODE_ERROR: {e}\nRAW_RESPONSE_START\n{raw_text[:1000]}\nRAW_RESPONSE_END"
        try:
            if response_data.get("new_kernel_created"):
                print(f"New kernel created for conversation {self.conv_id}")
            return response_data.get("result", "NO_RESULT_FIELD")
        except Exception as e:
            return f"EXECUTION_PARSE_ERROR: {e}\nRAW_JSON: {json.dumps(response_data)[:1000]}"
