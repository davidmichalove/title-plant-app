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
    driver.get("https://belmontcountyauditor.org/Disclaimer")
    time.sleep(2)
    agree_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'I AGREE')]"))
    )
    agree_btn.click()
    time.sleep(2)
    print("Agreed to disclaimer. Current URL:", driver.current_url)

    driver.get("https://belmontcountyauditor.org/RealEstate/Search")
    time.sleep(2)

    # Now we should be on Search page
    search_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "quickSearch"))
    )
    search_box.send_keys("53-01031.000")
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(4)
    print("Current URL after search:", driver.current_url)

    if "Results" in driver.current_url:
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        if rows:
            print("Clicking first result row...")
            rows[0].click()
            time.sleep(3)
            print("After clicking result row:", driver.current_url)

    import urllib.parse
    parsed = urllib.parse.urlparse(driver.current_url)
    params = urllib.parse.parse_qs(parsed.query)
    property_id = params.get('property_Id', [None])[0]
    print("Found property_Id:", property_id)
except Exception as e:
    print("Error:", e)

driver.quit()
