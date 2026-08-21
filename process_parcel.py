import sys
import os
import time
import base64
import requests
import re
import geopandas as gpd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchFrameException
import openpyxl
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.rich_text import InlineFont

TOWNSHIP_MAP = {
    'YOR': 'York', 'MEA': 'Mead', 'WAR': 'Warren', 'PEA': 'Pease',
    'COL': 'Colerain', 'FLS': 'Flushing', 'KIR': 'Kirkwood', 'SOM': 'Somerset',
    'WAY': 'Wayne', 'GOS': 'Goshen', 'UNI': 'Union', 'WHE': 'Wheeling',
    'RIC': 'Richland', 'SMI': 'Smith', 'WAS': 'Washington', 'PUL': 'Pultney'
}

def setup_driver(download_dir=None):
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--kiosk-printing') # Helps with some print dialogs
    if download_dir:
        prefs = {
            'download.default_directory': download_dir,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'plugins.always_open_pdf_externally': True
        }
        options.add_experimental_option('prefs', prefs)
    return webdriver.Chrome(options=options)

def print_to_pdf(driver, url, output_path):
    print(f"Printing {url} to {os.path.basename(output_path)}...")
    driver.get(url)
    time.sleep(5)
    try:
        pdf_data = driver.print_page()
        with open(output_path, "wb") as f:
            f.write(base64.b64decode(pdf_data))
    except Exception as e:
        print(f"Failed to print PDF: {e}")

def download_base64_pdf(url, output_path):
    print(f"Extracting real PDF from {url}...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8'
    }
    try:
        r = requests.get(url, headers=headers, timeout=20)
        match = re.search(r'data:application/pdf;base64,([^"\']+)', r.text)
        if match:
            pdf_data = base64.b64decode(match.group(1))
            with open(output_path, "wb") as f:
                f.write(pdf_data)
            print("Extracted successfully.")
        else:
            print("No base64 PDF found in source.")
    except Exception as e:
        print(f"Error fetching base64 PDF: {e}")

def download_raw_pdf(driver, url, output_path):
    print(f"Downloading raw PDF from {url}...")
    s = requests.Session()
    for cookie in driver.get_cookies():
        s.cookies.set(cookie['name'], cookie['value'])
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8'
    }
    try:
        r = s.get(url, headers=headers)
        if r.status_code == 200:
            with open(output_path, "wb") as f:
                f.write(r.content)
            print("Downloaded raw PDF successfully.")
        else:
            print(f"Failed to download raw PDF. Status: {r.status_code}")
    except Exception as e:
        print(f"Error downloading raw PDF: {e}")

def get_auditor_property_id(parcel_no):
    driver = setup_driver()
    try:
        driver.get("https://belmontcountyauditor.org/Search")
        
        # Wait up to 10 seconds for either Disclaimer or Search box
        start_time = time.time()
        while time.time() - start_time < 10:
            try:
                btns = driver.find_elements(By.XPATH, "//button[contains(text(), 'I AGREE')]")
                if btns and btns[0].is_displayed():
                    driver.execute_script("arguments[0].click();", btns[0])
                    time.sleep(2)
                    break
                    
                boxes = driver.find_elements(By.ID, "quickSearch")
                if boxes and boxes[0].is_displayed():
                    break
            except:
                pass
            time.sleep(0.5)
            
        wait = WebDriverWait(driver, 10)
        search_box = wait.until(EC.presence_of_element_located((By.ID, "quickSearch")))
        search_box.clear()
        search_box.send_keys(parcel_no)
        search_box.send_keys(Keys.ENTER)
        
        time.sleep(5)
        url = driver.current_url
        if "property_Id=" in url:
            return url.split("property_Id=")[1].split("&")[0]
            
        links = driver.find_elements(By.XPATH, "//a[contains(@href, 'property_Id=')]")
        for link in links:
            href = link.get_attribute("href")
            if parcel_no in href or "property_Id=" in href:
                return href.split("property_Id=")[1].split("&")[0]
    except Exception as e:
        print(f"Error finding property ID: {e}")
    finally:
        driver.quit()
    return None

def download_auditor_records(parcel_no, prop_id, tax_dir):
    driver = setup_driver()
    try:
        print("Accepting disclaimer for auditor records...")
        driver.get("https://belmontcountyauditor.org/Disclaimer?ReturnUrl=%2FSearch")
        time.sleep(2)
        btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'I AGREE')]")
        if btn: driver.execute_script("arguments[0].click();", btn[0])
        time.sleep(2)
        
        records = [
            (f"https://belmontcountyauditor.org/RealEstate/Tax?property_Id={prop_id}&rowNumber=0", "Current Tax"),
            (f"https://belmontcountyauditor.org/RealEstate/Payment?property_Id={prop_id}&rowNumber=0", "Payment History Card"),
            (f"https://belmontcountyauditor.org/RealEstate/Summary?property_Id={prop_id}&rowNumber=0", "Property Card")
        ]
        
        for url, name in records:
            out_path = os.path.join(tax_dir, f"PID {parcel_no} {name}.pdf")
            print_to_pdf(driver, url, out_path)
            
        tax_card_url = f"https://belmontcountyauditor.org/RealEstate/Default/TaxCard?Property_ID={prop_id}&Tax_Year=2025"
        tax_card_path = os.path.join(tax_dir, f"PID {parcel_no} Tax Card.pdf")
        download_raw_pdf(driver, tax_card_url, tax_card_path)
    finally:
        driver.quit()

def download_deeds(vol, pg, output_dir):
    print(f"Downloading deeds for Volume {vol}, Page {pg}...")
    driver = setup_driver(download_dir=output_dir)
    try:
        print("Logging in as davidtitle...")
        driver.get("https://countyfusion10.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
        time.sleep(3)
        try:
            driver.find_element(By.XPATH, "//input[@name='username']").send_keys("davidtitle")
            pass_input = driver.find_element(By.XPATH, "//input[@name='password']")
            pass_input.send_keys("test")
            pass_input.send_keys(Keys.ENTER)
        except: pass
        time.sleep(5)
        
        driver.switch_to.frame("bodyframe")
        try: driver.execute_script("executeCommand('Accept');")
        except: pass
        time.sleep(5)
        
        driver.switch_to.default_content()
        menus = driver.find_elements(By.XPATH, "//*[contains(text(), 'Search Public Records')]")
        if not menus: return
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
            
            try:
                alert = driver.switch_to.alert
                alert.accept()
                return
            except: pass
            
            driver.switch_to.default_content()
            driver.switch_to.frame("bodyframe")
            driver.switch_to.frame("resultFrame")
            try: driver.switch_to.frame("resultListFrame")
            except NoSuchFrameException: return
            
            try: num_results = driver.execute_script("return documentRowInfo ? documentRowInfo.length : 0;")
            except: num_results = 0
            
            for i in range(num_results):
                existing_files = set(os.listdir(output_dir))
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
                driver.execute_script("parent.downloadDocImage(instId, '', '', '', false);")
                
                downloaded_file = None
                for _ in range(60):
                    time.sleep(1)
                    new_files = set(os.listdir(output_dir)) - existing_files
                    new_files = [f for f in new_files if f.lower().endswith('.tif') or f.lower().endswith('.pdf')]
                    if new_files:
                        downloaded_file = new_files[0]
                        break
                        
                if downloaded_file:
                    old_path = os.path.join(output_dir, downloaded_file)
                    ext = os.path.splitext(old_path)[1]
                    new_name = f"{inst_type}_{vol}-{pg}{ext}".replace("/", "_").replace("\\", "_")
                    new_path = os.path.join(output_dir, new_name)
                    if os.path.exists(new_path):
                        new_path = os.path.join(output_dir, f"{inst_type}_{vol}-{pg}_{i}{ext}")
                    os.rename(old_path, new_path)
    except Exception as e:
        print(f"Deed download error: {e}")
    finally:
        try:
            driver.switch_to.default_content()
            driver.execute_script("if (typeof doLogOff == 'function') { doLogOff(); } else { window.location.href='/countyweb/loginDisplay.action'; }")
            time.sleep(3)
            driver.delete_all_cookies()
        except: pass
        driver.quit()

def populate_excel(parcel_data, template_path, output_path):
    print(f"Populating Excel file: {output_path}")
    wb = openpyxl.load_workbook(template_path, rich_text=True)
    sheet = wb.active
    
    red_font = InlineFont(color='FFFF0000')
    black_font = InlineFont(color='FF000000')
    
    parcel_no = str(parcel_data.get('parcel_no', ''))
    twp_code = str(parcel_data.get('twp', '')).upper()
    twp_name = TOWNSHIP_MAP.get(twp_code, twp_code)
    sec = str(parcel_data.get('sec', ''))
    t = str(parcel_data.get('t', ''))
    r = str(parcel_data.get('r', ''))
    vol = str(parcel_data.get('vol', ''))
    pg = str(parcel_data.get('pg', ''))
    ac = str(parcel_data.get('ac', ''))
    inst_type = str(parcel_data.get('type_deed', ''))
    eff_date = str(parcel_data.get('date_trans', ''))
    
    rich_string = CellRichText(
        TextBlock(black_font, "PARCEL ID #"),
        TextBlock(black_font, parcel_no if parcel_no else "XX-XXXXXXX.XXX"),
        TextBlock(black_font, f": {ac} acres, more or less, being located in the "),
        TextBlock(red_font, "QUARTER CALL"),
        TextBlock(black_font, " of Section "),
        TextBlock(black_font, sec if sec else "XX"),
        TextBlock(black_font, ", Township "),
        TextBlock(black_font, t if t else "X"),
        TextBlock(black_font, "N, Range "),
        TextBlock(black_font, r if r else "X"),
        TextBlock(black_font, "W, Township of "),
        TextBlock(black_font, twp_name if twp_name else "Mead"),
        TextBlock(black_font, ", County of Belmont, OH and being the property conveyed by "),
        TextBlock(black_font, inst_type if inst_type else "Instrument Type"),
        TextBlock(black_font, " from "),
        TextBlock(red_font, "Grantor"),
        TextBlock(black_font, " to "),
        TextBlock(red_font, "Grantee"),
        TextBlock(black_font, ", and said deed being dated "),
        TextBlock(black_font, eff_date if eff_date else "effective date XX/XX/XXXX"),
        TextBlock(black_font, " and filed of record under Volume "),
        TextBlock(black_font, vol if vol else "XX"),
        TextBlock(black_font, ", Page "),
        TextBlock(black_font, pg if pg else "XX"),
        TextBlock(black_font, ", "),
        TextBlock(red_font, "Record Type"),
        TextBlock(black_font, " Records, Belmont County, OH.")
    )
    
    sheet['B3'] = rich_string
    if t and r: sheet['A4'] = f"TOWNSHIP {t}N - RANGE {r}W"
    if sec: sheet['A5'] = f"SECTION {sec}: TRACT #1"
    if ac: sheet['B7'] = ac

    wb.save(output_path)

def main():
    if len(sys.argv) < 2:
        print("Usage: python process_parcel.py <parcel_id>")
        sys.exit(1)
        
    target_parcel = sys.argv[1]
    
    print(f"Loading shapefile to find parcel {target_parcel}...")
    shapefile_path = '/Users/davidmichalove/Desktop/automate/app/shape_files/Belmont_County_Parcels.shp'
    gdf = gpd.read_file(shapefile_path)
    
    parcel_row = gdf[gdf['parcel_no'] == target_parcel]
    if parcel_row.empty:
        print(f"Error: Parcel ID {target_parcel} not found in shapefile.")
        sys.exit(1)
        
    parcel_data = parcel_row.iloc[0].to_dict()
    
    base_dir = f'/Users/davidmichalove/Desktop/automate/PID {target_parcel}'
    docs_dir = os.path.join(base_dir, 'DOCS')
    maps_dir = os.path.join(base_dir, 'MAPS')
    tax_dir = os.path.join(base_dir, 'TAX')
    well_dir = os.path.join(base_dir, 'WELL INFO')
    
    for d in [docs_dir, maps_dir, tax_dir, well_dir]:
        os.makedirs(d, exist_ok=True)
        
    # Generate OR Excel (User said skip RS for now)
    template_1 = '/Users/davidmichalove/Desktop/automate/PID OR (DATE)_TEMPLATE (2).xlsx'
    out_1 = os.path.join(base_dir, f"PID {target_parcel} OR.xlsx")
    populate_excel(parcel_data, template_1, out_1)
    
    # Download deeds
    vol, pg = parcel_data.get('vol'), parcel_data.get('pg')
    if vol and pg and not str(vol) == 'nan' and not str(pg) == 'nan':
        download_deeds(vol, pg, docs_dir)
    else:
        print(f"No volume/page data. Skipping deed download.")
        
    driver = setup_driver()
    try:
        # Transfer Card
        transfer_url = parcel_data.get('hyperlink')
        if transfer_url and isinstance(transfer_url, str):
            out_transfer = os.path.join(tax_dir, f"PID {target_parcel} Transfer Card.pdf")
            download_base64_pdf(transfer_url, out_transfer)
        else:
            print("No hyperlink in shapefile for Transfer Card.")
            
        # Tax Map
        twp_code = str(parcel_data.get('twp', '')).upper()
        sec_num = str(parcel_data.get('sec', '')).split('.')[0] # e.g. "13.0" -> "13"
        try: sec_num = f"{int(sec_num):02d}"
        except: pass
        
        map_url = None
        if twp_code == 'WAR' and sec_num:
            map_url = f"https://belcogis.com/php/taxmapview.php?TableAndName=taxmaps2026:{sec_num}0806:War"
        elif twp_code == 'BAR' or twp_code == 'SOM': # Note Barnesville twp code is unknown, maybe BAR or WAR?
            pass
        elif twp_code == 'SOM' and sec_num:
            map_url = f"https://belcogis.com/php/taxmapview.php?TableAndName=taxmaps2026:{sec_num}0706:Som"
            
        if map_url:
            out_map = os.path.join(maps_dir, f"PID {target_parcel} Tax Map.pdf")
            download_base64_pdf(map_url, out_map)
        else:
            print(f"Tax map URL logic not defined for twp {twp_code} sec {sec_num}. Skipping Tax Map.")
            
    finally:
        driver.quit()
        
    # Auditor Records
    prop_id = get_auditor_property_id(target_parcel)
    if prop_id:
        print(f"Found Property ID: {prop_id}")
        download_auditor_records(target_parcel, prop_id, tax_dir)
    else:
        print("Failed to find Property ID on auditor site. Skipping auditor records.")

    print(f"All processing complete for {target_parcel}! Check the folder {base_dir}")

if __name__ == "__main__":
    main()
