import os
from title_work_automator import AutomatorApp
from playwright.sync_api import sync_playwright

def test():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.goto("https://countyfusion13.kofiletech.us/countyweb/loginDisplay.action?countyname=BelmontOH")
        page.locator("input[value='Login as Guest']").click()
        page.wait_for_timeout(3000)
        page.frame_locator("iframe[name='bodyframe']").locator("input#accept").click()
        page.wait_for_timeout(3000)
        
        page.frame_locator("iframe[name='bodyframe']").locator("text='Search Public Records'").first.click()
        page.wait_for_timeout(3000)
        
        page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").get_by_role("tab", name="Name").click()
        page.wait_for_timeout(2000)
        cf = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").frame_locator("iframe[name='criteriaframe']")
        cf.get_by_label("Name", exact=True).fill("Phillips Ross")
        page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='dynSearchFrame']").locator("img#imgSearch").click()
        
        page.wait_for_timeout(3000)
        reslist = page.frame_locator("iframe[name='bodyframe']").frame_locator("iframe[name='resultFrame']").frame_locator("iframe[name='resultListFrame']")
        reslist.locator("tr").first.wait_for(state="visible", timeout=15000)
        
        data = reslist.locator("body").evaluate("""
            () => {
                let rows = document.querySelectorAll('table tr');
                let result = [];
                for (let row of rows) {
                    let cols = row.querySelectorAll('th, td');
                    let rowData = [];
                    for (let col of cols) {
                        let text = col.innerText.trim();
                        rowData.push(text.replace(/\\n/g, ' '));
                    }
                    if (rowData.length > 2) {
                        result.push(rowData);
                    }
                }
                return result;
            }
        """)
        for i, row in enumerate(data[:10]):
            print(f"Row {i}: {row}")
        browser.close()

test()
