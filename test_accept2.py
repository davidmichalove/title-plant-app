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
    print("Navigating...")
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    driver.find_element(By.ID, "usernameInput").send_keys("davidtitle")
    driver.find_element(By.ID, "passwordInput").send_keys("test")
    driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    print("Logged in...")
    time.sleep(5)
    
    driver.switch_to.frame("bodyframe")
    print("Executing executeCommand('Accept')")
    driver.execute_script("executeCommand('Accept');")
    time.sleep(5)
    
    driver.switch_to.default_content()
    
    menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    if menus:
        print("Clicking Search Public Records...")
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        
        try:
            driver.switch_to.frame("bodyframe")
            print("Switched to bodyframe after click.")
            
            tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book/Page')]")
            if tabs:
                print("Clicking Book/Page tab...")
                driver.execute_script("arguments[0].click();", tabs[0])
                time.sleep(5)
                
                inputs = driver.find_elements(By.XPATH, "//input")
                print("Inputs found:", len(inputs))
                for inp in inputs:
                    print("Input:", inp.get_attribute("name"), inp.get_attribute("id"))
                
                driver.save_screenshot("step9_search_form.png")
            else:
                print("No Book/Page tab found.")
        except Exception as e:
            print("Error in bodyframe:", e)
    else:
        print("No search menu.")

finally:
    driver.quit()
