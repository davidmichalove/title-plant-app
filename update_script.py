with open('/Users/davidmichalove/Desktop/automate/app/title_work_automator.py', 'r') as f:
    content = f.read()

content = content.replace(
    "driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)",
    "driver = webdriver.Chrome(options=options)"
)
content = content.replace("from webdriver_manager.chrome import ChromeDriverManager\n", "")
content = content.replace("from selenium.webdriver.chrome.service import Service\n", "")

with open('/Users/davidmichalove/Desktop/automate/app/title_work_automator.py', 'w') as f:
    f.write(content)
