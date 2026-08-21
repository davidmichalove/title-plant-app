from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
driver = webdriver.Chrome(options=options)

try:
    print("Searching Auditor Site...")
    driver.get("https://belmontcountyauditor.org/Disclaimer")
    time.sleep(2)
    try:
        agree_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'I AGREE')]")
        driver.execute_script("arguments[0].click();", agree_btn)
        time.sleep(2)
    except Exception as e:
        print("Disclaimer error:", e)
    
    driver.get("https://belmontcountyauditor.org/RealEstate/Search")
    time.sleep(2)
    
    search_box = driver.find_element(By.ID, "quickSearch")
    search_box.send_keys("42-01003.000")
    search_box.send_keys(Keys.RETURN)
    time.sleep(4)
    
    print("URL after search:", driver.current_url)
    
    if "Results" in driver.current_url:
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        print(f"Found {len(rows)} result rows.")
        if rows:
            print("Clicking first result row...")
            try:
                # Let's try to click the first cell of the row or the row itself
                driver.execute_script("arguments[0].click();", rows[0])
                time.sleep(3)
            except Exception as e:
                print("Click row error:", e)
                
    print("URL after clicking:", driver.current_url)
finally:
    driver.quit()
