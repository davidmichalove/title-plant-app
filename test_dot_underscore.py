import os, json, time
from google import genai

config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
with open(config_path, 'r') as f:
    config = json.load(f)
api_key = config.get('GEMINI_API_KEY')

client = genai.Client(api_key=api_key)

f = client.files.upload(file="/Volumes/davidlls/SOPs/._Combined_SOPs.pdf", config={'display_name': 'test.pdf'})
while True:
    info = client.files.get(name=f.name)
    if 'ACTIVE' in str(info.state):
        print("File ACTIVE!")
        break
    elif 'FAILED' in str(info.state):
        print("File FAILED!")
        break
    time.sleep(1)

try:
    response = client.models.generate_content(
        model='gemini-3.6-flash',
        contents=[f, "hello"],
    )
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {e}")
finally:
    client.files.delete(name=f.name)
