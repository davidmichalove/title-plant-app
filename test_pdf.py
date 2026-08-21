from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import base64

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(options=options)
driver.get("https://belmontcountyauditor.org/RealEstate/Search")
# wait a bit for it to load
driver.implicitly_wait(2)

pdf = driver.print_page()

with open("test.pdf", "wb") as f:
    f.write(base64.b64decode(pdf))

driver.quit()
print("Saved test.pdf")
