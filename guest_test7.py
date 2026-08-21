import time
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
            
            # Now we need to figure out where the search button is and click it.
            # In Kofile, you often call `parent.search()` or `executeCommand('Search')`.
            print("Executing executeCommand('Search') on dynSearchFrame...")
            driver.switch_to.parent_frame()
            driver.execute_script("executeCommand('Search');")
            time.sleep(10)
            
            driver.save_screenshot("step14_after_search.png")
            print("Saved step14_after_search.png")
finally:
    driver.quit()
