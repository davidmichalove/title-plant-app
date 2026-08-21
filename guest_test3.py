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
    
    print("Executing guest login...")
    driver.execute_script("doGuestLogin(true);")
    time.sleep(5)
    
    try:
        driver.switch_to.frame("bodyframe")
        print("Clicking accept...")
        try:
            driver.execute_script("executeCommand('Accept');")
        except:
            accept_btn = driver.find_element(By.XPATH, "//input[@name='accept' or @value='Accept']")
            driver.execute_script("arguments[0].click();", accept_btn)
        time.sleep(5)
    except Exception as e:
        print("No bodyframe or disclaimer:", e)
        
    driver.switch_to.default_content()
    
    menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    if menus:
        print("Clicking Search Public Records...")
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        
        driver.switch_to.frame("bodyframe")
        tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book/Page')]")
        if tabs:
            print("Clicking Book/Page...")
            driver.execute_script("arguments[0].click();", tabs[0])
            time.sleep(5)
            
            inputs = driver.find_elements(By.XPATH, "//input")
            for i in inputs:
                if 'book' in i.get_attribute("name").lower() or 'vol' in i.get_attribute("name").lower():
                    print("Found Book Input:", i.get_attribute("name"))
                if 'page' in i.get_attribute("name").lower():
                    print("Found Page Input:", i.get_attribute("name"))
                    
            driver.save_screenshot("step10_book_page.png")
            print("Saved step10_book_page.png")
        else:
            print("No Book/Page tab found.")
    else:
        print("No Search Public Records menu found.")
finally:
    driver.quit()
