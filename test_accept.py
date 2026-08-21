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
    print("In bodyframe. Executing executeCommand('Accept')")
    driver.execute_script("executeCommand('Accept');")
    time.sleep(5)
    
    driver.switch_to.default_content()
    driver.save_screenshot("step7_post_execute_command.png")
    
    # Check if "Search Public Records" is visible
    menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    print(f"Found {len(menus)} 'Search Public Records' elements.")
    if menus:
        print("Clicking Search Public Records...")
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        driver.save_screenshot("step8_post_menu_click.png")

finally:
    driver.quit()
