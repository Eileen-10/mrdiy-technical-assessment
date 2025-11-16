import time
import os
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DOWNLOAD_DIR = Path.cwd() / "data"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

chrome_options = Options()
prefs = {
    "download.default_directory": str(DOWNLOAD_DIR),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

# 1. Currency Dataset Creation
try:
    driver.get("https://www.bnm.gov.my/exchange-rates")

    wait = WebDriverWait(driver, 10)
    export_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary"))
    )

    export_button.click()

    # Wait until file finish downloading
    file_downloaded = False
    for _ in range(30):
        files = os.listdir(DOWNLOAD_DIR)

        if any(f.endswith(".csv") for f in files):
            print("Download complete:", [f for f in files if f.endswith(".csv")][0])
            file_downloaded = True
            break

        if any(f.endswith(".crdownload") for f in files):
            time.sleep(1)
        else:
            time.sleep(1)

    if not file_downloaded:
        print("CSV download failed.")

finally:
    driver.quit()


# 2. Dummy Dataset Creation
excel_file = DOWNLOAD_DIR / "excel_sample_data_qae.xlsx"
sheets_to_export = {
    "python_test-sales": "sales.csv",
    "python_test-product": "product.csv",
    "python_test-store": "store.csv"
}

for sheet_name, output_name in sheets_to_export.items():
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    output_path = DOWNLOAD_DIR / output_name
    df.to_csv(output_path, index=False)
    print(f"Saved: {output_path}")
