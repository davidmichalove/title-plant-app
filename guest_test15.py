import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
prefs = {'download.default_directory' : '/Users/davidmichalove/Desktop/automate/app/docs'}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=options)

def find_and_click_save(driver):
    # Search current frame
    try:
        btns = driver.find_elements(By.XPATH, "//*[contains(text(), 'Save Image')]")
        if btns:
            print("Found Save Image in frame:", driver.current_url)
            # Find the closest parent 'a' or 'div' with onclick
            elem = btns[0]
            for _ in range(3):
                onclick = elem.get_attribute('onclick')
                href = elem.get_attribute('href')
                print("  Tag:", elem.tag_name, "onclick:", onclick, "href:", href)
                elem = elem.find_element(By.XPATH, "..")
            return True
    except:
        pass
    
    # Search subframes
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in frames:
        try:
            driver.switch_to.frame(frame)
            if find_and_click_save(driver):
                driver.switch_to.parent_frame()
                return True
            driver.switch_to.parent_frame()
        except:
            driver.switch_to.parent_frame()
            pass
            
    return False

try:
    os.makedirs('/Users/davidmichalove/Desktop/automate/app/docs', exist_ok=True)
    
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
        
        tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
        if tabs:
            driver.execute_script("arguments[0].click();", tabs[0])
            time.sleep(5)
            
            driver.switch_to.frame("criteriaframe")
            book_input = driver.find_element(By.XPATH, "//input[@aria-label='Book']")
            page_input = driver.find_element(By.XPATH, "//input[@aria-label='Page']")
            book_input.send_keys("798")
            page_input.send_keys("642")
            
            page_input.send_keys(Keys.ENTER)
            time.sleep(10)
            
            driver.switch_to.default_content()
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("resultFrame")
            driver.switch_to.frame("resultListFrame")
            
            driver.execute_script("loadRecord(documentRowInfo[0]);")
            time.sleep(10)
            
            driver.switch_to.default_content()
            find_and_click_save(driver)
finally:
    driver.quit()
