import os, json
from google import genai

config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
with open(config_path, 'r') as f:
    config = json.load(f)
api_key = config.get('GEMINI_API_KEY')

client = genai.Client(api_key=api_key)
try:
    response = client.models.generate_content(
        model='gemini-3.6-flash',
        contents="Say hello",
        config={'system_instruction': 'You are a helpful assistant', 'temperature': 0.1}
    )
    print("SUCCESS")
except Exception as e:
    print(f"FAILED: {e}")
