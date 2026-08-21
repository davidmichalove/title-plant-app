import time
import os
import sys
import glob
import geopandas as gpd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchFrameException, TimeoutException

def download_documents():
    output_dir = '/Users/davidmichalove/Desktop/automate/app/docs'
    os.makedirs(output_dir, exist_ok=True)
    
    print("Reading shapefile...")
    gdf = gpd.read_file('/Users/davidmichalove/Desktop/automate/app/shape_files/Belmont_County_Parcels.shp')
    
    valid_rows = gdf[gdf['vol'].notna() & gdf['pg'].notna()]
    unique_deeds = valid_rows[['vol', 'pg']].drop_duplicates()
    
    deeds_to_fetch = list(unique_deeds.itertuples(index=False, name=None))
    print(f"Found {len(deeds_to_fetch)} unique deeds to fetch.")
    
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 5
    if limit > 0:
        deeds_to_fetch = deeds_to_fetch[:limit]
        print(f"Limiting to first {limit} deeds for this run.")
        
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    prefs = {'download.default_directory' : output_dir}
    options.add_experimental_option('prefs', prefs)
    
    driver = webdriver.Chrome(options=options)
    
    try:
        print("Logging in as davidtitle...")
        driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
        time.sleep(3)
        try:
            user_input = driver.find_element(By.XPATH, "//input[@name='username']")
            pass_input = driver.find_element(By.XPATH, "//input[@name='password']")
            user_input.send_keys("davidtitle")
            pass_input.send_keys("test")
            pass_input.send_keys(Keys.ENTER)
        except Exception as e:
            print("Could not find login fields:", e)
        time.sleep(5)
        
        driver.switch_to.frame("bodyframe")
        try: driver.execute_script("executeCommand('Accept');")
        except: pass
        time.sleep(5)
        
        for vol, pg in deeds_to_fetch:
            print(f"Searching for Volume: {vol}, Page: {pg}...")
            driver.switch_to.default_content()
            
            menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
            if not menus:
                print("Could not find 'Search Public Records' menu.")
                continue
            driver.execute_script("arguments[0].click();", menus[0])
            time.sleep(5)
            
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("dynSearchFrame")
            
            tabs = driver.find_elements(By.XPATH, "//*[contains(text(), 'Book') and contains(text(), 'Page')]")
            if tabs:
                driver.execute_script("arguments[0].click();", tabs[0])
                time.sleep(3)
                
                driver.switch_to.frame("criteriaframe")
                book_input = driver.find_element(By.XPATH, "//input[@aria-label='Book']")
                page_input = driver.find_element(By.XPATH, "//input[@aria-label='Page']")
                
                book_input.clear()
                page_input.clear()
                
                book_input.send_keys(str(vol))
                page_input.send_keys(str(pg))
                page_input.send_keys(Keys.ENTER)
                time.sleep(10)
                
                # Check if there is an alert
                try:
                    alert = driver.switch_to.alert
                    print(f"Alert found: {alert.text}")
                    alert.accept()
                    continue
                except:
                    pass
                
                driver.switch_to.default_content()
                driver.switch_to.frame("bodyframe")
                driver.switch_to.frame("resultFrame")
                try:
                    driver.switch_to.frame("resultListFrame")
                except NoSuchFrameException:
                    print(f"No results found for Volume {vol} Page {pg}")
                    continue
                
                try:
                    num_results = driver.execute_script("return documentRowInfo ? documentRowInfo.length : 0;")
                except:
                    num_results = 0
                
                print(f"Found {num_results} result(s).")
                for i in range(num_results):
                    print(f"Loading result {i+1}...")
                    
                    existing_files = set(os.listdir(output_dir))
                    
                    driver.switch_to.default_content()
                    driver.switch_to.frame("bodyframe")
                    driver.switch_to.frame("resultFrame")
                    driver.switch_to.frame("resultListFrame")
                    
                    try:
                        inst_type = driver.execute_script(f"return documentRowInfo[{i}].instType;")
                    except:
                        inst_type = "UNKNOWN"
                        
                    driver.execute_script(f"loadRecord(documentRowInfo[{i}]);")
                    time.sleep(10)
                    
                    driver.switch_to.default_content()
                    driver.switch_to.frame("bodyframe")
                    driver.switch_to.frame("documentFrame")
                    
                    driver.execute_script("parent.downloadDocImage(instId, '', '', '', false);")
                    
                    max_wait = 60
                    downloaded_file = None
                    for _ in range(max_wait):
                        time.sleep(1)
                        current_files = set(os.listdir(output_dir))
                        new_files = current_files - existing_files
                        new_files = [f for f in new_files if f.lower().endswith('.tif') or f.lower().endswith('.pdf')]
                        if new_files:
                            downloaded_file = new_files[0]
                            break
                            
                    if downloaded_file:
                        old_path = os.path.join(output_dir, downloaded_file)
                        ext = os.path.splitext(old_path)[1]
                        new_name = f"{inst_type}_{vol}-{pg}{ext}"
                        new_name = new_name.replace("/", "_").replace("\\", "_")
                        new_path = os.path.join(output_dir, new_name)
                        
                        if os.path.exists(new_path):
                            new_path = os.path.join(output_dir, f"{inst_type}_{vol}-{pg}_{i}{ext}")
                            
                        os.rename(old_path, new_path)
                        print(f"Downloaded and saved as: {os.path.basename(new_path)}")
                    else:
                        print("Download timed out or failed.")
            
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        try:
            print("Logging out...")
            driver.switch_to.default_content()
            driver.execute_script("if (typeof doLogOff == 'function') { doLogOff(); } else { window.location.href='/countyweb/loginDisplay.action'; }")
            time.sleep(3)
            driver.delete_all_cookies()
        except Exception as e:
            pass
        driver.quit()

if __name__ == "__main__":
    download_documents()
