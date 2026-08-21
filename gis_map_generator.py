import time
import os
from playwright.sync_api import sync_playwright

def generate_gis_map(parcel_num, output_dir):
    with sync_playwright() as p:
        # Use a persistent context so the browser CACHES all the heavy map files!
        cache_dir = os.path.join(output_dir, "playwright_cache")
        context = p.chromium.launch_persistent_context(
            user_data_dir=cache_dir,
            headless=False,
            viewport={'width': 1440, 'height': 900}, # Scaled down to match a typical MacBook Air screen!
            ignore_https_errors=True
        )
        page = context.pages[0] if context.pages else context.new_page()
        # Rename context for later reference in script
        browser = context
        
        print(f"Loading GIS Map for {parcel_num}...")
        page.goto(f"https://gis.belcogis.com/ParcelMap/#widget_48=text:{parcel_num}&zoom_to_selection=true")
        
        # Wait smartly! Instead of a hard 125 second sleep, we wait exactly until the Zoom Out button appears!
        print("Waiting for the GIS map to load (using smart cache checking)...")
        page.wait_for_load_state('domcontentloaded')
        try:
            page.wait_for_selector("[aria-label='Zoom out']", timeout=150000)
            print("Map loaded successfully!")
            time.sleep(3) # tiny buffer for map tiles
        except Exception as e:
            print("Map took longer than 150 seconds to load!")
            time.sleep(2)
        
        try:
            print("Waiting for network to go idle (downloading map tiles for new parcels)...")
            try:
                page.wait_for_load_state('networkidle', timeout=60000)
            except Exception:
                pass # If it times out, it's just background analytics keeping it alive.
                
            print("Giving the map 15 seconds to physically finish 'flying in' to the parcel...")
            time.sleep(15) 
            
            print("Waiting for Zoom Out button to become active...")
            page.wait_for_selector("[aria-label='Zoom out']:not([disabled])", timeout=30000)
            print("Clicking native Zoom Out button...")
            page.locator("[aria-label='Zoom out']").first.click()
            time.sleep(2)
        except Exception as e:
            print("Native Zoom out failed:", e)
            
        try:
            print("Clicking Print Widget natively...")
            page.locator("[aria-label='Print']").first.click()
            time.sleep(3)
            
            print("Entering title...")
            # Clear it first, then type slowly to ensure it registers
            title_input = page.locator("input[type='text']").nth(1) # The first is usually map search, second is title!
            if title_input.is_visible():
                title_input.fill(f"{parcel_num}")
            else:
                page.locator("input[type='text']").last.fill(f"{parcel_num}")
            time.sleep(1)
            
            print("Clicking actual Print button inside widget...")
            page.locator("button:has-text('Print')").last.click()
            
            # Smart waiting instead of 25 second hard sleep!
            print("Clicking Results tab natively...")
            try:
                page.locator("text=Results").last.click()
            except: pass
            
            print("Waiting for the loading spinner to vanish and the result link to appear...")
            # Just wait for the parcel number text to appear in the DOM!
            page.wait_for_selector(f"a:has-text('{parcel_num}')", timeout=60000)
            
            print("Catching the new tab (popup) triggered by the link...")
            try:
                with page.expect_popup(timeout=15000) as popup_info:
                    # Click the link to trigger the popup
                    page.locator(f"a:has-text('{parcel_num}')").last.click()
                
                popup = popup_info.value
                popup.wait_for_load_state()
                pdf_url = popup.url
                print(f"Successfully caught the popup URL: {pdf_url}")
                
                if ".pdf" in pdf_url.lower() or "arcgisoutput" in pdf_url.lower():
                    map_path = os.path.join(output_dir, f"PID {parcel_num} GIS 2026 Map.pdf")
                    # Use Playwright's internal request to download it using the browser's SSL context
                    response = context.request.get(pdf_url, ignore_https_errors=True)
                    with open(map_path, 'wb') as out_file:
                        out_file.write(response.body())
                    print(f"Successfully downloaded native GIS Map to: {map_path}")
                    # Open the PDF automatically on macOS for easy verification!
                    import subprocess
                    subprocess.Popen(['open', map_path])
                else:
                    print(f"Popup URL doesn't look like a PDF! It was: {pdf_url}")
            except Exception as e:
                print(f"Failed to catch popup: {e}")
            
        except Exception as e:
            print("Native Print failed, falling back to screenshot PDF...", e)
            map_path = os.path.join(output_dir, f"PID {parcel_num} GIS 2026 Map.pdf")
            page.pdf(path=map_path, print_background=True, width="1920px", height="1080px")
            print(f"Saved fallback PDF to {map_path}")
            
        print('Leaving browser open for 15 seconds so you can see what happened...')
        time.sleep(15)
        browser.close()

if __name__ == "__main__":
    import sys
    parcel = sys.argv[1] if len(sys.argv) > 1 else "42-00998.000"
    out = sys.argv[2] if len(sys.argv) > 2 else "/tmp"
    generate_gis_map(parcel, out)
