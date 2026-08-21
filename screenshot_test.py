import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('window-size=1920x1080')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
driver = webdriver.Chrome(options=options)

try:
    print("Navigating to login...")
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    driver.save_screenshot("step1_login.png")
    
    driver.find_element(By.ID, "usernameInput").send_keys("davidtitle")
    driver.find_element(By.ID, "passwordInput").send_keys("test")
    driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    print("Logged in...")
    time.sleep(5)
    
    driver.save_screenshot("step2_post_login.png")
    
    try:
        driver.switch_to.frame("bodyframe")
        accept_btn = driver.find_element(By.XPATH, "//input[@name='accept' or @value='Accept' or @value='I Accept' or contains(@onclick, 'accept')]")
        driver.execute_script("arguments[0].click();", accept_btn)
        print("Clicked accept")
        time.sleep(5)
    except:
        pass
        
    driver.switch_to.default_content()
    driver.save_screenshot("step3_post_accept.png")
    print("Saved all screenshots")
finally:
    driver.quit()
