import sys
import os
import time
from playwright.sync_api import sync_playwright

def launch_recorder_browser():
    user_data_dir = os.path.join(os.path.expanduser("~"), ".kofile_playwright_profile")
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            viewport=None,
            args=['--start-maximized']
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://countyfusion13.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
        
        # 1. Login as Guest
        try:
            page.locator("input[value='Login as Guest']").wait_for(timeout=5000)
            page.locator("input[value='Login as Guest']").click(no_wait_after=True)
            page.wait_for_load_state('domcontentloaded')
        except Exception:
            pass
            
        # 2. Accept disclaimer
        try:
            page.frame_locator("iframe[name='bodyframe']").locator("input#accept").wait_for(timeout=5000)
            page.frame_locator("iframe[name='bodyframe']").locator("input#accept").click()
            page.wait_for_load_state('domcontentloaded')
        except Exception:
            pass
            
        # 3. Search Public Records
        try:
            page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first.wait_for(timeout=5000)
            page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first.click()
        except Exception:
            pass
            
        # Keep process alive while browser window is open
        while True:
            try:
                if not context.pages or context.pages[0].is_closed():
                    break
                time.sleep(1)
            except Exception:
                break
                
        try:
            context.close()
        except Exception:
            pass

if __name__ == "__main__":
    launch_recorder_browser()
