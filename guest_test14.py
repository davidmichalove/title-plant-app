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
driver = webdriver.Chrome(options=options)

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
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("documentFrame")
            driver.switch_to.frame("docInfoFrame")
            
            with open("docInfoFrame.html", "w") as f:
                f.write(driver.page_source)
            print("Saved docInfoFrame.html")
            
            save_img = driver.find_elements(By.XPATH, "//*[contains(text(), 'Save Image')]")
            if save_img:
                print("Found Save Image button, click string:", save_img[0].get_attribute('onclick') or save_img[0].find_element(By.XPATH, '..').get_attribute('onclick') or save_img[0].find_element(By.XPATH, '../..').get_attribute('onclick'))
finally:
    driver.quit()
