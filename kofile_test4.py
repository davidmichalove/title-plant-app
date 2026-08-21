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
    time.sleep(2)
    
    driver.find_element(By.ID, "usernameInput").send_keys("davidtitle")
    driver.find_element(By.ID, "passwordInput").send_keys("test")
    driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    time.sleep(10)
    
    # Dump main source
    with open("main_source.html", "w") as f:
        f.write(driver.page_source)
        
    driver.switch_to.frame("bodyframe")
    with open("body_source.html", "w") as f:
        f.write(driver.page_source)
        
    print("Saved HTML to main_source.html and body_source.html")
    
finally:
    driver.quit()
