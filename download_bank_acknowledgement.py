import time
import pytest
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options

# Load test data from an Excel file
def load_data_from_excel():
    # Specify the dtype for the "CHALLAN" column as string
    df = pd.read_excel("data.xlsx", dtype={"CHALLAN": str})  # Replace with your Excel file path
    return df.to_dict(orient="records")  # Convert the DataFrame to a list of dictionaries


@pytest.mark.parametrize("data", load_data_from_excel())
def test_download_a_challan(data):
    # Configure Firefox options
    firefox_options = Options()
    firefox_options.headless = True  # Run in headless mode

    firefox_options.set_preference("print.always_print_silent", True)  # Silent printing
    firefox_options.set_preference("print.show_print_progress", False)
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_file", True)
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_filename", "output.pdf")
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_paper_name", "iso_a4")
    firefox_options.set_preference("print.printer_Mozilla_Save_to_PDF.print_to_file", True)

    # Disable headers and footers
    firefox_options.set_preference("print.print_headerleft", "")  # No left header
    firefox_options.set_preference("print.print_headercenter", "")  # No center header
    firefox_options.set_preference("print.print_headerright", "")  # No right header
    firefox_options.set_preference("print.print_footerleft", "")  # No left footer
    firefox_options.set_preference("print.print_footercenter", "")  # No center footer
    firefox_options.set_preference("print.print_footerright", "")  # No right footer

    driver = webdriver.Firefox(options=firefox_options)

    driver.implicitly_wait(5)
    driver.get("http://103.48.16.132/echalan/echalan_iframe.php")

    challan_year = driver.find_element(By.CSS_SELECTOR, '#txtChalanNoA')
    challan_year.send_keys(str(data["YEAR"]))

    challan_id = driver.find_element(By.CSS_SELECTOR, '#txtChalanNoA1')
    challan_id.send_keys(str(data["CHALLAN"]))

    verify_button = driver.find_element(By.CSS_SELECTOR, '#cmdVerify1')
    verify_button.click()


    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    time.sleep(10)

    # Use the DevTools Protocol to print the page as a PDF
    print_settings = {
        "paperWidth": 8.5,
        "paperHeight": 11.0,
        "printBackground": True
    }
    driver.execute_script("window.print();")

    # Allow the print dialog to open and interact with it
    time.sleep(2)  # Adjust the time as needed for the dialog to appear

    # Type the file name in the "Save As" dialog
    pyautogui.write(str(data["NAME"]) + "_"+ str(data["ID"]) + "_bank.pdf")

    # Press Enter to confirm the save
    pyautogui.press("enter")

    driver.switch_to.window(window_before)
    driver.quit()
