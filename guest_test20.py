import sys
import subprocess

with open('download_deeds.py', 'r') as f:
    content = f.read()
    
# Replace the data loading with hardcoded for test
content = content.replace("deeds_to_fetch = list(unique_deeds.itertuples(index=False, name=None))", "deeds_to_fetch = [('345', '2')]")
with open('test_download_deeds.py', 'w') as f:
    f.write(content)
