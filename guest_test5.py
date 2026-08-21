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
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    driver.execute_script("doGuestLogin(true);")
    time.sleep(5)
    
    driver.switch_to.frame("bodyframe")
    try: driver.execute_script("executeCommand('Accept');")
    except: pass
    time.sleep(5)
        
    driver.switch_to.default_content()
    
    menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    if menus:
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        
        driver.switch_to.frame("bodyframe")
        driver.switch_to.frame("dynSearchFrame")
        
        # Click the tab
        tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
        if tabs:
            print("Clicking Book/Page tab...")
            driver.execute_script("arguments[0].click();", tabs[0])
            time.sleep(5)
            
            # Print inputs
            for inp in driver.find_elements(By.XPATH, "//input"):
                print("Input:", inp.get_attribute('name'), inp.get_attribute('id'))
                
            driver.save_screenshot("step13_after_book_click.png")
            print("Saved step13_after_book_click.png")
        else:
            print("No Book/Page tab found in dynSearchFrame.")
            # Maybe searchFrame?
            driver.switch_to.default_content()
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("searchFrame")
            tabs2 = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
            if tabs2:
                print("Clicking Book/Page tab in searchFrame...")
                driver.execute_script("arguments[0].click();", tabs2[0])
                time.sleep(5)
                for inp in driver.find_elements(By.XPATH, "//input"):
                    print("Input:", inp.get_attribute('name'), inp.get_attribute('id'))
                driver.save_screenshot("step13_after_book_click.png")
            else:
                print("Not found in searchFrame either.")
finally:
    driver.quit()
