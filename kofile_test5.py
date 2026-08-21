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
    time.sleep(5)
    
    driver.switch_to.frame("bodyframe")
    print("In bodyframe, looking for accept...")
    try:
        accept_btn = driver.find_element(By.XPATH, "//input[@name='accept' or @value='Accept' or @value='I Accept' or contains(@onclick, 'accept')]")
        accept_btn.click()
        print("Clicked Accept on disclaimer!")
        time.sleep(10)
    except Exception as e:
        print("Could not click accept:", e)
        
    with open("body_source2.html", "w") as f:
        f.write(driver.page_source)
        
    print("Saved HTML to body_source2.html")
    
finally:
    driver.quit()
