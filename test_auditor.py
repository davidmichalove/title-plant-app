import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def get_property_id(parcel_no):
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    
    try:
        print(f"Searching auditor site for {parcel_no}...")
        driver.get("https://belmontcountyauditor.org/RealEstate/Search")
        time.sleep(3)
        
        # In typical auditor sites, there's a search by parcel number
        # Let's find an input that accepts a parcel number
        # Often it has 'parcel' or 'pin' in the id/name
        # We'll just look for inputs and try to guess, or we can dump the page source first
        source = driver.page_source
        with open('auditor_source.html', 'w') as f:
            f.write(source)
            
        print("Saved auditor_source.html")
    except Exception as e:
        print(e)
    finally:
        driver.quit()

get_property_id("41-00396.001")
