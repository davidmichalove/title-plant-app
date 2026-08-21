import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

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
    print("Executed doGuestLogin(true)")
    time.sleep(10)
    
    driver.save_screenshot("step5_post_guest.png")
    print("Saved step5_post_guest.png")
    
    try:
        driver.switch_to.frame("bodyframe")
        driver.save_screenshot("step6_bodyframe.png")
    except:
        print("no bodyframe")

finally:
    driver.quit()
