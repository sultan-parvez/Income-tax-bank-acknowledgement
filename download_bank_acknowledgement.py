import time
import pytest
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import base64
import os

# Load test data from an Excel file
def load_data_from_excel():
  # Specify the dtype for the "CHALLAN" column as string
    df = pd.read_excel("data.xlsx", dtype={"CHALLAN": str})  # Replace with your Excel file path
    return df.to_dict(orient="records")  # Convert the DataFrame to a list of dictionaries

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--ignore-certificate-errors")  # Ignore SSL certificate errors
chrome_options.add_argument("--headless=new")  # Headless mode for running without UI

@pytest.mark.parametrize("data", load_data_from_excel())
def test_download_a_challan(data):
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    driver.get("http://103.48.16.132/echalan/echalan_iframe.php")

    challan_year = driver.find_element(By.CSS_SELECTOR, '#txtChalanNoA')
    challan_year.send_keys(str(data["YEAR"]))
    time.sleep(3)

    challan_id = driver.find_element(By.CSS_SELECTOR, '#txtChalanNoA1')
    challan_id.send_keys(str(data["CHALLAN"]))

    time.sleep(2)

    verify_button = driver.find_element(By.CSS_SELECTOR, '#cmdVerify1')
    verify_button.click()


    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    time.sleep(7)

    print_settings = {
        "paperWidth": 8.5,  # Width of the paper in inches
        "paperHeight": 11.0,  # Height of the paper in inches
        "printBackground": True,  # Include background graphics
        "scale": 0.80  # Scale the content to 80%
    }

    result = driver.execute_cdp_cmd("Page.printToPDF", print_settings)
    output_dir = "bank_acknowledgements"
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{data['NAME']}{data['ID']}.pdf")
    with open(output_file, "wb") as f:
        f.write(base64.b64decode(result["data"]))

    print(f"PDF saved as {output_file}")

    driver.switch_to.window(window_before)
    driver.quit()
