import sys
import os
import time
from playwright.sync_api import sync_playwright

def generate_odnr_map(parcel_num, output_dir):
    with sync_playwright() as p:
        # User Data Directory for caching (skips disclaimer on subsequent runs)
        user_data_dir = os.path.join(os.path.expanduser("~"), ".odnr_playwright_cache")
        browser = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            viewport={'width': 1440, 'height': 900},
            args=['--start-maximized']
        )
        
        page = browser.pages[0] if browser.pages else browser.new_page()
        
        print(f"Loading ODNR Map for {parcel_num}...")
        page.goto("https://experience.arcgis.com/experience/5fd9c17bc933417d9bd8d7c1dff404aa/page/Map-View", timeout=60000)
        
        # 1. Handle Disclaimer if it exists
        try:
            print("Pressing Escape to clear disclaimer (if present)...")
            page.wait_for_timeout(4000)
            # Click the center of the screen just to ensure focus
            page.mouse.click(10, 10)
            page.keyboard.press("Escape")
            page.wait_for_timeout(1000)
            page.keyboard.press("Escape")
            
            # Fallback if there's an actual button we can click
            try:
                ok_btn = page.locator("button:has-text('OK'), button:has-text('Agree'), button:has-text('Accept'), button:has-text('Continue')").first
                if ok_btn.is_visible(timeout=2000):
                    print("Found an OK/Accept button, clicking it...")
                    ok_btn.click()
            except:
                pass
        except Exception:
            pass
        
        # Wait for the main UI to load
        print("Waiting for network to go idle...")
        try:
            page.wait_for_load_state('networkidle', timeout=15000)
        except Exception:
            pass

        # 2. Toggle Legend Layers
        print("Configuring layers...")
        try:
            print("Clicking the Oil and Gas Fields layer row...")
            layer_label = page.locator("div.label").filter(has_text="Oil and Gas Fields").locator("visible=true").first
            layer_label.wait_for(state="visible", timeout=10000)
            layer_label.evaluate("el => el.click()")
            
            time.sleep(1)
            print("Clicking the Orphan Well layer row...")
            orphan_label = page.locator("div.label").filter(has_text="Orphan Well").locator("visible=true").first
            orphan_label.evaluate("el => el.click()")
            time.sleep(1)
        except Exception as e:
            print(f"Could not toggle layers: {e}")

        # 3. Open Parcel Search accordion
        print("Opening Parcel Search...")
        try:
            parcel_accordion = page.locator("text=Parcel Search").first
            parcel_accordion.click()
            time.sleep(1)
        except Exception as e:
            print(f"Failed to open Parcel Search accordion: {e}")

        # 4. Type Parcel Number and select Belmont
        print(f"Searching for Parcel: {parcel_num}")
        try:
            # The search input might be inside the expanded accordion.
            # We can find the input by placeholder or just the visible textbox under Parcel Search.
            # Let's find the input inside the active search widget.
            search_input = page.locator("input[placeholder*='parcel']").first
            if not search_input.is_visible():
                search_input = page.locator("input[type='text']").nth(1) # fallback
            
            search_input.click()
            search_input.fill(parcel_num)
            page.keyboard.press("Enter")
            time.sleep(3) # wait for dropdown to populate
            
            # Now find the Belmont result
            belmont_result = page.locator(f"text={parcel_num}, Belmont").first
            belmont_result.wait_for(state="visible", timeout=10000)
            belmont_result.click()
            
            print("Giving the map 8 seconds to physically finish 'flying in' to the parcel...")
            time.sleep(8)
        except Exception as e:
            print(f"Search failed: {e}")

        # 5. Zoom Out once
        try:
            print("Waiting for Zoom Out button to become active...")
            page.wait_for_selector("[aria-label='Zoom out']:not([disabled])", timeout=15000)
            print("Clicking native Zoom Out button...")
            page.locator("[aria-label='Zoom out']").first.click()
            time.sleep(3)
        except Exception as e:
            print(f"Native Zoom out failed: {e}")

        # 6. Click Print Button in top menu
        print("Clicking Print Widget natively...")
        try:
            # Top menu print button
                        # Find the element and force a direct javascript click to bypass any pointer-events interference!
            # Use get_by_role to strictly find the button element whose accessible name (aria-label/title) is Print!
            print_btn = page.get_by_role("button", name="Print").locator("visible=true").first
            print_btn.click()
            time.sleep(2)
        except Exception as e:
            print(f"Failed to click Print menu button: {e}")

        # 7. Enter Title and Print
        print("Entering title...")
        try:
            # The input might not have the value attribute directly in the DOM, so let's use get_by_label or find it relative to the 'Title' text
            title_input = page.locator("input[type='text']").filter(has_text="") # fallback if needed
            # We know the print widget is active, let's just get the last visible textbox
            visible_textboxes = page.locator("input[type='text']")
            
            # Wait for at least one textbox to be visible in the print widget
            page.wait_for_timeout(3000)
            
            # Use javascript to find the input whose current value is "Ohio Oil & Gas Wells Map"
            title_input = page.locator("input[type='text']").nth(0) # placeholder
            count = visible_textboxes.count()
            found = False
            for i in range(count):
                el = visible_textboxes.nth(i)
                if el.is_visible():
                    val = el.evaluate("el => el.value")
                    if "Ohio Oil & Gas Wells Map" in val:
                        title_input = el
                        found = True
                        break
            
            if not found:
                print("Could not find the exact title input, grabbing the last visible one!")
                for i in range(count-1, -1, -1):
                    if visible_textboxes.nth(i).is_visible():
                        title_input = visible_textboxes.nth(i)
                        break
            # Clear the box first to ensure React notices the change!
            title_input.click()
            page.keyboard.press("Meta+a")
            page.keyboard.press("Backspace")
            
            title_input.fill(f"Ohio Oil & Gas Wells Map PID {parcel_num}")
            # Press Tab to trigger the React "onBlur" event so it legally saves the text into memory before printing!
            page.keyboard.press("Tab")
            time.sleep(1)
            
            print("Clicking actual Print button inside widget...")
            # The print button inside the widget
            print_btn = page.locator("button:has-text('Print')").last
            print_btn.click()
            
            print("Waiting for the loading spinner to vanish and the result link to appear...")
            # The result is usually added to a 'Results' tab or just appears as a link.
            # In Experience Builder, we might need to click a 'Results' tab first.
            results_tab = page.locator("text=Results").first
            if results_tab.is_visible():
                results_tab.click()
            
            # Wait for the anchor tag with the title text
            page.wait_for_selector(f"a:has-text('PID {parcel_num}')", timeout=60000)
            
            print("Catching the new tab (popup) triggered by the link...")
            with page.expect_popup(timeout=15000) as popup_info:
                page.locator(f"a:has-text('PID {parcel_num}')").last.click()
            
            popup = popup_info.value
            popup.wait_for_load_state()
            pdf_url = popup.url
            print(f"Successfully caught the popup URL: {pdf_url}")
            
            if ".pdf" in pdf_url.lower() or "arcgisoutput" in pdf_url.lower():
                map_path = os.path.join(output_dir, f"PID {parcel_num} Well Interactive Map.pdf")
                response = browser.request.get(pdf_url, ignore_https_errors=True)
                with open(map_path, 'wb') as out_file:
                    out_file.write(response.body())
                print(f"Successfully downloaded native ODNR GIS Map to: {map_path}")
                import subprocess
                subprocess.Popen(['open', map_path])
            else:
                print(f"Popup URL doesn't look like a PDF! It was: {pdf_url}")

        except Exception as e:
            print("Native Print failed, falling back to screenshot PDF...", e)
            map_path = os.path.join(output_dir, f"PID {parcel_num} Well Interactive Map.pdf")
            page.pdf(path=map_path, print_background=True, width="1440px", height="900px")
            print(f"Saved fallback PDF to {map_path}")

        print('Leaving browser open for 15 seconds so you can see what happened...')
        time.sleep(8)
        browser.close()

if __name__ == "__main__":
    parcel = sys.argv[1] if len(sys.argv) > 1 else "42-00124.000"
    out = sys.argv[2] if len(sys.argv) > 2 else "/tmp"
    generate_odnr_map(parcel, out)
