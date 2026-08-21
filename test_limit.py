import os, json
from google import genai

config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
with open(config_path, 'r') as f:
    config = json.load(f)
api_key = config.get('GEMINI_API_KEY')

client = genai.Client(api_key=api_key)

files = []
for i in range(14):
    f = client.files.upload(file="/Volumes/davidlls/SOPs/Combined_SOPs.pdf", config={'display_name': f'test_{i}.pdf'})
    files.append(f)

print("Uploaded 14 files. Waiting for them to become ACTIVE...")
active = []
import time
for f in files:
    while True:
        info = client.files.get(name=f.name)
        if 'ACTIVE' in str(info.state):
            active.append(f)
            break
        elif 'FAILED' in str(info.state):
            print("Failed")
            break
        time.sleep(1)

print("Calling generate_content...")
try:
    response = client.models.generate_content(
        model='gemini-3.6-flash',
        contents=active + ["what are the lookback time frames?"],
        config={'system_instruction': 'You are a bot', 'temperature': 0.1}
    )
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {e}")

for f in files:
    client.files.delete(name=f.name)
