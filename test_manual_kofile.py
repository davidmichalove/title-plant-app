import os
import time
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchFrameException

def log(msg):
    print(f"{datetime.datetime.now().time()} - {msg}")

def test_kofile(vol, pg, docs_dir):
    log(f"Downloading deeds for Volume {vol}, Page {pg}...")
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('window-size=1920x1080')
    options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    prefs = {
        'download.default_directory': docs_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'plugins.always_open_pdf_externally': True
    }
    options.add_experimental_option('prefs', prefs)
    
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)
    try:
        log("Navigating to login...")
        driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
        time.sleep(3)
        log("Logging in as davidtitle...")
        try:
            driver.find_element(By.XPATH, "//input[@name='username']").send_keys("davidtitle")
            pass_input = driver.find_element(By.XPATH, "//input[@name='password']")
            pass_input.send_keys("test")
            pass_input.send_keys(Keys.ENTER)
        except Exception as e:
            log(f"Login error: {e}")
        time.sleep(5)
        
        log("Accepting disclaimer...")
        driver.switch_to.frame("bodyframe")
        try: driver.execute_script("executeCommand('Accept');")
        except: pass
        time.sleep(5)
        
        log("Clicking Search Public Records...")
        driver.switch_to.default_content()
        menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
        if not menus: 
            log("Menu not found")
            return
        driver.execute_script("arguments[0].click();", menus[0])
        time.sleep(5)
        
        log("Clicking Book/Page tab...")
        driver.switch_to.frame("bodyframe")
        driver.switch_to.frame("dynSearchFrame")
        tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
        if tabs:
            driver.execute_script("arguments[0].click();", tabs[0])
            time.sleep(3)
            log("Submitting search...")
            driver.switch_to.frame("criteriaframe")
            book_input = driver.find_element(By.XPATH, "//input[@aria-label='Book']")
            page_input = driver.find_element(By.XPATH, "//input[@aria-label='Page']")
            book_input.clear()
            page_input.clear()
            book_input.send_keys(str(vol))
            page_input.send_keys(str(pg))
            page_input.send_keys(Keys.ENTER)
            log("Wait 10 seconds for results...")
            time.sleep(10)
            
            try:
                alert = driver.switch_to.alert
                alert.accept()
                log("Alert accepted (no results?)")
                return
            except: pass
            
            log("Checking results frame...")
            driver.switch_to.default_content()
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("resultFrame")
            try: driver.switch_to.frame("resultListFrame")
            except NoSuchFrameException:
                log("resultListFrame not found")
                return
            
            try: num_results = driver.execute_script("return documentRowInfo ? documentRowInfo.length : 0;")
            except: num_results = 0
            
            log(f"Found {num_results} results.")
            
            for i in range(num_results):
                log(f"Processing result {i+1} of {num_results}...")
                existing_files = set(os.listdir(docs_dir))
                driver.switch_to.default_content()
                driver.switch_to.frame("bodyframe")
                driver.switch_to.frame("resultFrame")
                driver.switch_to.frame("resultListFrame")
                try: inst_type = driver.execute_script(f"return documentRowInfo[{i}].instType;")
                except: inst_type = "UNKNOWN"
                    
                driver.execute_script(f"loadRecord(documentRowInfo[{i}]);")
                time.sleep(10)
                
                driver.switch_to.default_content()
                driver.switch_to.frame("bodyframe")
                driver.switch_to.frame("documentFrame")
                log("Triggering download...")
                driver.execute_script("parent.downloadDocImage(instId, '', '', '', false);")
                
                downloaded_file = None
                log("Waiting for file...")
                for t in range(120):
                    time.sleep(1)
                    new_files = set(os.listdir(docs_dir)) - existing_files
                    new_files = [f for f in new_files if f.lower().endswith('.tif') or f.lower().endswith('.pdf')]
                    if new_files:
                        downloaded_file = list(new_files)[0]
                        log(f"Found file: {downloaded_file}")
                        break
                        
                if downloaded_file:
                    old_path = os.path.join(docs_dir, downloaded_file)
                    ext = os.path.splitext(old_path)[1]
                    new_name = f"{inst_type}_{vol}-{pg}{ext}".replace("/", "_").replace("\\", "_")
                    new_path = os.path.join(docs_dir, new_name)
                    if os.path.exists(new_path):
                        new_path = os.path.join(docs_dir, f"{inst_type}_{vol}-{pg}_{i}{ext}")
                    os.rename(old_path, new_path)
                    log(f"Renamed to: {new_path}")
                else:
                    log("Timeout waiting for file.")
    except Exception as e:
        log(f"Deed download error: {e}")
    finally:
        driver.quit()

test_kofile("534", "804", "/Users/davidmichalove/Desktop/automate/app/TEST_DOCS")
