import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)

try:
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    for b in soup.find_all(['button', 'input']):
        if 'guest' in b.get('value', '').lower() or 'guest' in b.text.lower():
            print("GUEST BUTTON HTML:", str(b))
finally:
    driver.quit()
