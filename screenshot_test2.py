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
    
    # Guest Login might be easier, let's check for guest login first
    guest_buttons = driver.find_elements(By.XPATH, "//*[@value='Login as Guest' or contains(text(), 'Guest')]")
    if guest_buttons:
        print("Found Guest button!")
        driver.execute_script("arguments[0].click();", guest_buttons[0])
    else:
        print("No Guest button. Using credentials...")
        driver.find_element(By.ID, "usernameInput").send_keys("davidtitle")
        driver.find_element(By.ID, "passwordInput").send_keys("test")
        driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
    
    time.sleep(5)
    
    try:
        driver.switch_to.frame("bodyframe")
        accept_btns = driver.find_elements(By.XPATH, "//*[@name='accept' or @value='Accept']")
        for b in accept_btns:
            print(f"Trying to click accept button: {b.get_attribute('outerHTML')}")
            try:
                b.click()
                print("Clicked with .click()")
                break
            except:
                try:
                    driver.execute_script("arguments[0].click();", b)
                    print("Clicked with execute_script")
                    break
                except Exception as e:
                    print("Failed:", e)
    except Exception as e:
        print("Error looking for accept:", e)
        
    time.sleep(5)
    driver.switch_to.default_content()
    driver.save_screenshot("step4_post_accept.png")
    print("Saved screenshot")
finally:
    driver.quit()
