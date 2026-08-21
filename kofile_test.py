import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('window-size=1920x1080')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
driver = webdriver.Chrome(options=options)

try:
    print("Navigating to Kofile login...")
    driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
    time.sleep(3)
    
    # Check if there is a disclaimer button (sometimes named "I Accept" or similar)
    try:
        accept_btn = driver.find_element(By.XPATH, "//input[@value='I Accept' or @value='Accept']")
        accept_btn.click()
        print("Clicked Disclaimer Accept")
        time.sleep(2)
    except Exception as e:
        print("No disclaimer accept button found or needed.")
    
    # Login
    try:
        user_input = driver.find_element(By.NAME, "username")
        pass_input = driver.find_element(By.NAME, "password")
        user_input.send_keys("davidtitle")
        pass_input.send_keys("test")
        
        login_btn = driver.find_element(By.XPATH, "//input[@value='Login' or @type='submit']")
        login_btn.click()
        print("Logged in!")
        time.sleep(4)
    except Exception as e:
        print("Could not find login fields:", e)
        print("Page Source:", driver.page_source[:500])

    print("Current URL:", driver.current_url)
    # Check if "Search Public Records" is a button/link
    try:
        search_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Search Public Records')]|//span[contains(text(), 'Search Public Records')]")
        search_link.click()
        print("Clicked Search Public Records!")
        time.sleep(4)
    except Exception as e:
        print("Could not find 'Search Public Records'. Looking for iframes or frames.")
        frames = driver.find_elements(By.TAG_NAME, "frame")
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"Frames: {len(frames)}, iFrames: {len(iframes)}")
        if frames:
            driver.switch_to.frame("bodyframe") # Common kofile frame name
            print("Switched to bodyframe")
            try:
                search_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Search Public Records')]")
                search_link.click()
                print("Clicked Search Public Records inside frame!")
                time.sleep(4)
            except Exception as e2:
                print("Still didn't find it.", e2)
                print(driver.page_source[:1000])
        else:
            print(driver.page_source[:1000])
            
    # Try logging out safely
    driver.switch_to.default_content()
    try:
        if frames:
            driver.switch_to.frame("title")
        logout = driver.find_element(By.XPATH, "//a[contains(text(), 'Logout')]|//img[contains(@alt, 'Logout')]/..")
        logout.click()
        print("Logged out successfully.")
    except:
        print("Could not logout.")

finally:
    driver.quit()
