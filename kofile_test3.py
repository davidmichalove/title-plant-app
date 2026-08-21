import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)

try:
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    
    # Login
    user_input = driver.find_element(By.ID, "usernameInput")
    pass_input = driver.find_element(By.ID, "passwordInput")
    user_input.send_keys("davidtitle")
    pass_input.send_keys("test")
    
    login_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
    login_btn.click()
    time.sleep(5)
    
    print("Post-login frames:")
    for frame in driver.find_elements(By.TAG_NAME, "frame"):
        print("Frame:", frame.get_attribute("name"))
        
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        print("IFrame:", iframe.get_attribute("name"))
        
    driver.switch_to.frame("bodyframe")
    print("\nBodyframe HTML snippet:")
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    for a in soup.find_all('a'):
        print(a.text.strip(), a.get('href'), a.get('onclick'))
        
finally:
    driver.quit()
