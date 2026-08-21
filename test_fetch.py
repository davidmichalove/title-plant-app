import requests
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

def test_url(url):
    print("Testing:", url)
    try:
        r = requests.get(url, headers=headers, timeout=20)
        content_type = r.headers.get('content-type', '')
        print("Status:", r.status_code)
        print("Content-Type:", content_type)
        if 'html' in content_type:
            print("HTML Snippet:", r.text[:200].replace('\n', ' '))
            embeds = re.search(r'data:application/pdf;base64,([^"\']+)', r.text)
            if embeds:
                print("Found base64 PDF in HTML")
    except Exception as e:
        print("Error:", e)
    print("-" * 40)

test_url("https://belcogis.com/php/taxmapview.php?TableAndName=taxmaps2026:130806:War")
