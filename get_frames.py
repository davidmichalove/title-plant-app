import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)

def print_frames(driver, indent=""):
    frames = driver.find_elements(webdriver.common.by.By.TAG_NAME, "iframe")
    for i, frame in enumerate(frames):
        name = frame.get_attribute("name")
        id_ = frame.get_attribute("id")
        print(f"{indent}Frame {i}: id={id_} name={name}")
        driver.switch_to.frame(frame)
        print_frames(driver, indent + "  ")
        driver.switch_to.parent_frame()

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
    menus = driver.find_elements(webdriver.common.by.By.XPATH, "//*[contains(text(), 'Search Public Records')]")
    if menus:
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        print("After clicking Search Public Records:")
        print_frames(driver)
finally:
    driver.quit()
