import requests
import re
from bs4 import BeautifulSoup

# Let's check belmontcountyrecorder.com or org for the direct link to Kofile
urls = ["https://www.belmontcountyrecorder.com/", "https://belmontcountyrecorder.org/"]

for u in urls:
    try:
        r = requests.get(u, timeout=10)
        print("Testing", u)
        for link in BeautifulSoup(r.text, 'html.parser').find_all('a'):
            href = link.get('href', '')
            if 'kofile' in href.lower() or 'countyfusion' in href.lower():
                print("FOUND KOFILE LINK:", href)
    except Exception as e:
        print(e)
