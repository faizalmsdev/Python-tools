from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

# Setup Chrome options
options = Options()
options.add_argument("--headless")  # Remove this line if you want to see the browser
options.add_argument("--disable-gpu")

# Set up the webdriver (update path if chromedriver is not in PATH)
service = Service()
driver = webdriver.Chrome(service=service, options=options)

# Open the webpage
url = "https://visualping.io/diff/827381793?disableId=e809c41fc04b4a6&mode=visual"
driver.get(url)

# Wait for the page to fully load
time.sleep(5)

# Get the page source
html_content = driver.page_source

# Save to file
with open("scraper.html", "w", encoding="utf-8") as file:
    file.write(html_content)

# Clean up
driver.quit()

print("HTML content saved to 'scraper.html'")
