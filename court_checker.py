import csv
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, "Court_Documents")

def init_browser():
    options = Options()
    # We want it visible so the user can pass the Captcha
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    # Optional: setup a profile so it remembers cookies/sessions if helpful
    # options.add_argument(f"user-data-dir={os.path.join(BASE_DIR, 'chrome_profile')}")

    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1200, 900)
    return driver

def wait_for_search_page(driver, timeout=300):
    """Wait for the user to solve the Captcha and reach the search page."""
    print("Waiting for user to solve Captcha and reach the Search page...")
    end_time = time.time() + timeout
    while time.time() < end_time:
        if "search.page" in driver.current_url:
            print("Successfully reached the search page!")
            return True
        time.sleep(2)
    return False

def find_input(driver, possible_names, label_text):
    """Robustly find an input field by common names or by preceding label."""
    for name in possible_names:
        try:
            el = driver.find_element(By.XPATH, f"//input[contains(@name, '{name}') or contains(@id, '{name}')]")
            if el.is_displayed():
                return el
        except:
            pass
            
    # Try finding by label
    try:
        el = driver.find_element(By.XPATH, f"//label[contains(text(), '{label_text}')]/following::input[1]")
        if el.is_displayed():
            return el
    except:
        pass
        
    return None

def process_court_records(csv_path, update_status_callback=None):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    if not os.path.exists(csv_path):
        if update_status_callback:
            update_status_callback(f"Error: CSV file not found at {csv_path}")
        return

    results = []
    
    driver = None
    try:
        if update_status_callback:
            update_status_callback("Launching browser. Please solve the Captcha if prompted.")
            
        driver = init_browser()
        driver.get("https://eservices.belmontcountycourts.com/eservices/home.page.2")
        
        # Wait up to 5 minutes for user to click through to search.page
        if not wait_for_search_page(driver, 300):
            if update_status_callback:
                update_status_callback("Timed out waiting for Search page. Did you pass the Captcha?")
            return

        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            # Normalize column headers
            fieldnames = [col.strip().lower() for col in reader.fieldnames]
            
            # Find the best match for first and last name columns
            fname_col = next((c for c in reader.fieldnames if 'first' in c.lower()), None)
            lname_col = next((c for c in reader.fieldnames if 'last' in c.lower()), None)
            
            if not fname_col or not lname_col:
                if update_status_callback:
                    update_status_callback("Error: CSV must contain columns for first name and last name.")
                return

            for row in reader:
                fname = row.get(fname_col, "").strip()
                lname = row.get(lname_col, "").strip()
                
                if not fname and not lname:
                    continue
                
                if update_status_callback:
                    update_status_callback(f"Searching for {fname} {lname}...")
                
                # Navigate to the base search page to reset
                # (We can just click the Search tab, or just use the current URL if we are already there)
                # It's safer to re-click the Name tab or just clear the form
                try:
                    # Look for clear button or just clear inputs
                    lname_input = find_input(driver, ["lastName", "last"], "Last Name")
                    fname_input = find_input(driver, ["firstName", "first"], "First Name")
                    
                    if lname_input:
                        lname_input.clear()
                        lname_input.send_keys(lname)
                    
                    if fname_input:
                        fname_input.clear()
                        fname_input.send_keys(fname)
                        
                    # Find and click Search/Submit
                    # CourtView usually has a submit button or an input with type=submit at the bottom of the form
                    submit_btn = driver.find_element(By.XPATH, "//input[@type='submit' or @value='Search'] | //button[contains(text(), 'Search')]")
                    driver.execute_script("arguments[0].click();", submit_btn)
                    
                    # Wait for results
                    time.sleep(3)
                    
                    # Parse results
                    # Check if there are no results
                    page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                    if "no results" in page_text or "0 results" in page_text or "found 0" in page_text:
                        row['Court_Status'] = "Not Found"
                        results.append(row)
                        driver.back() # Go back to search form
                        time.sleep(2)
                        continue
                        
                    # Find case status in the table
                    # CourtView usually has a table with class 'resultTable' or similar
                    has_open = False
                    try:
                        # Grab all text in the table, or look for specific keywords
                        tables = driver.find_elements(By.TAG_NAME, "table")
                        results_text = ""
                        for table in tables:
                            if "Case" in table.text or "Status" in table.text:
                                results_text += table.text.upper()
                                
                        if "OPEN" in results_text or "ACTIVE" in results_text:
                            has_open = True
                        elif "CLOSED" in results_text or "ARCHIVED" in results_text or "INACTIVE" in results_text:
                            has_open = False
                        else:
                            # If we can't tell, assume manual review is needed
                            has_open = "Manual Review"
                    except Exception as e:
                        has_open = "Manual Review"
                        
                    if has_open == True:
                        row['Court_Status'] = "Has Open Cases"
                        
                        # Make a folder
                        person_dir = os.path.join(OUTPUT_DIR, f"{lname}_{fname}")
                        os.makedirs(person_dir, exist_ok=True)
                        
                        # Try to click the first case link (just as an example, this gets complex)
                        try:
                            case_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'case.page')]")
                            if case_links:
                                driver.execute_script("arguments[0].click();", case_links[0])
                                time.sleep(3)
                                
                                # Take screenshot of case details
                                driver.save_screenshot(os.path.join(person_dir, "case_summary.png"))
                                
                                # Look for Documents or Dockets tab
                                tabs = driver.find_elements(By.XPATH, "//a[contains(text(), 'Document') or contains(text(), 'Docket')]")
                                if tabs:
                                    driver.execute_script("arguments[0].click();", tabs[0])
                                    time.sleep(3)
                                    driver.save_screenshot(os.path.join(person_dir, "dockets_tab.png"))
                                    
                                # Go back to results, then back to search
                                driver.back()
                                time.sleep(2)
                        except Exception as e:
                            print(f"Error exploring case for {fname} {lname}: {e}")
                        
                    elif has_open == "Manual Review":
                        row['Court_Status'] = "Unknown - Needs Review"
                    else:
                        row['Court_Status'] = "Closed"
                        
                    results.append(row)
                    
                    # Go back to search page
                    driver.back()
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Error processing {fname} {lname}: {e}")
                    row['Court_Status'] = "Error"
                    results.append(row)
                    
        # Write results
        out_csv = os.path.join(BASE_DIR, "court_checker_results.csv")
        if results:
            with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
                fieldnames = list(results[0].keys())
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(results)
                
        if update_status_callback:
            update_status_callback(f"Done! Results saved to {out_csv}")
            
    except Exception as e:
        if update_status_callback:
            update_status_callback(f"An error occurred: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    # For testing directly
    test_csv = os.path.join(BASE_DIR, "test_names.csv")
    if not os.path.exists(test_csv):
        with open(test_csv, "w") as f:
            f.write("first_name,last_name\nJohn,Dolton\nJane,Doe")
    process_court_records(test_csv, print)
