import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchFrameException

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
prefs = {
    'download.default_directory': '/Users/davidmichalove/Desktop/automate/app/TEST_DOCS',
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'plugins.always_open_pdf_externally': True
}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=options)
try:
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    driver.find_element(By.XPATH, "//input[@name='username']").send_keys("davidtitle")
    pass_input = driver.find_element(By.XPATH, "//input[@name='password']")
    pass_input.send_keys("test")
    pass_input.send_keys(Keys.ENTER)
    time.sleep(5)
    
    driver.switch_to.frame("bodyframe")
    try: driver.execute_script("executeCommand('Accept');")
    except: pass
    time.sleep(5)
    
    driver.switch_to.default_content()
    menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    driver.execute_script("arguments[0].click();", menus[0])
    time.sleep(5)
    
    driver.switch_to.frame("bodyframe")
    driver.switch_to.frame("dynSearchFrame")
    tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
    driver.execute_script("arguments[0].click();", tabs[0])
    time.sleep(3)
    driver.switch_to.frame("criteriaframe")
    book_input = driver.find_element(By.XPATH, "//input[@aria-label='Book']")
    page_input = driver.find_element(By.XPATH, "//input[@aria-label='Page']")
    book_input.send_keys("534")
    page_input.send_keys("804")
    page_input.send_keys(Keys.ENTER)
    time.sleep(10)
    
    driver.switch_to.default_content()
    driver.switch_to.frame("bodyframe")
    driver.switch_to.frame("resultFrame")
    driver.switch_to.frame("resultListFrame")
    
    i = 1 # Result 2
    driver.execute_script(f"loadRecord(documentRowInfo[{i}]);")
    time.sleep(10)
    
    driver.switch_to.default_content()
    driver.switch_to.frame("bodyframe")
    driver.switch_to.frame("documentFrame")
    driver.save_screenshot(f"/Users/davidmichalove/Desktop/automate/app/TEST_DOCS/result2_before.png")
    
    driver.execute_script("parent.downloadDocImage(instId, '', '', '', false);")
    time.sleep(5)
    driver.save_screenshot(f"/Users/davidmichalove/Desktop/automate/app/TEST_DOCS/result2_after.png")
    time.sleep(15)
finally:
    driver.quit()
