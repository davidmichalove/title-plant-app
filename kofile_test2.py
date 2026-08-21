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
    driver.get("https://countyfusion13.govos.com/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    forms = soup.find_all('form')
    for i, form in enumerate(forms):
        print(f"--- Form {i} ---")
        for inpt in form.find_all('input'):
            print(inpt)
        for btn in form.find_all('button'):
            print(btn)
        for a in form.find_all('a'):
            print(a)
    
finally:
    driver.quit()
