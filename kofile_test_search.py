import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('window-size=1920x1080')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
driver = webdriver.Chrome(options=options)

try:
    print("Navigating to Kofile login...")
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(2)
    
    # Login
    user_input = driver.find_element(By.ID, "usernameInput")
    pass_input = driver.find_element(By.ID, "passwordInput")
    user_input.send_keys("davidtitle")
    pass_input.send_keys("test")
    
    driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    print("Clicked Login")
    time.sleep(5)
    
    # Kofile usually has a frame named 'bodyframe'
    driver.switch_to.frame("bodyframe")
    print("Switched to bodyframe")
    
    # Find Search Public Records
    search_links = driver.find_elements(By.XPATH, "//a[contains(text(), 'Search Public Records')]|//span[contains(text(), 'Search Public Records')]")
    if search_links:
        print("Found Search Public Records, clicking...")
        search_links[0].click()
        time.sleep(5)
    else:
        print("Could not find Search Public Records link.")
        print(driver.page_source[:500])
        
    driver.switch_to.default_content()
    # It might open a new frame or just reload bodyframe
    time.sleep(2)
    driver.switch_to.frame("bodyframe")
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    tabs = soup.find_all(lambda tag: tag.name in ['a', 'span', 'div'] and 'Book/Page' in tag.text)
    print("Tabs containing Book/Page:")
    for t in tabs:
        print(t.name, t.get('id'), t.get('class'), t.text.strip()[:30])
        
    inputs = soup.find_all('input')
    print("\nInput fields:")
    for i in inputs:
        print(i.get('name'), i.get('id'), i.get('type'))
        
finally:
    driver.quit()
