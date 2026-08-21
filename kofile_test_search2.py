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
    driver.find_element(By.ID, "usernameInput").send_keys("davidtitle")
    driver.find_element(By.ID, "passwordInput").send_keys("test")
    driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    time.sleep(4)
    
    driver.switch_to.frame("bodyframe")
    print("In bodyframe after login.")
    
    # Click Accept on Disclaimer if present
    try:
        accept_btn = driver.find_element(By.XPATH, "//input[@name='accept' or @value='Accept' or @value='I Accept']")
        accept_btn.click()
        print("Clicked Accept on disclaimer.")
        time.sleep(4)
    except:
        pass
        
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    print("Looking for Search Public Records or Book/Page...")
    for a in soup.find_all(['a', 'span', 'button', 'div']):
        text = a.text.strip()
        if 'Search Public' in text or 'Book/Page' in text:
            print("Found:", a.name, a.get('id'), a.get('class'), text)
            
    # Try clicking "Search Public Records"
    try:
        driver.find_element(By.XPATH, "//*[contains(text(), 'Search Public Records')]").click()
        print("Clicked Search Public Records!")
        time.sleep(4)
    except Exception as e:
        print("Could not click Search Public Records:", e)
        
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    for a in soup.find_all(['a', 'span', 'td', 'div']):
        text = a.text.strip()
        if 'Book/Page' in text:
            print("Found Book/Page tab:", a.name, a.get('id'), a.get('class'), text)

    print("\nInputs currently visible:")
    for i in soup.find_all('input'):
        print(i.get('name'), i.get('id'), i.get('type'))
        
finally:
    driver.quit()
