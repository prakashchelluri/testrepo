import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Path to your ChromeDriver executable
chromedriver_path = r"C:\Users\satv1214\Downloads\chromedriver-win64\chromedriver.exe"  # Update this path

# Create a Service object with the path to ChromeDriver
service = Service(chromedriver_path)

try:
    # Initialize the Chrome WebDriver with the Service object
    driver = webdriver.Chrome(service=service)

    # Load the Excel file
    excel_file = r"C:\Users\satv1214\Downloads\shift_data.xlsx"
    df = pd.read_excel(excel_file)

    # Function to fill out the form
    def fill_form(row):
        name = row['Name']
        erp_id = str(row['ERP ID'])  # Convert to string
        line_manager = row['Line Manager']
        project_details = row['Project Details']
        work_date = row['Work Date']
        shift_type = row['Type of On-Call / Shift']

        # URL of the form page
        url = "https://sps.netcracker.com/sites/India/HR/Lists/On%20Call%20Allowance/NewForm.aspx?Source=https%3A%2F%2Fsps%2Enetcracker%2Ecom%2Fsites%2FIndia%2FHR%2FLists%2FOn%2520Call%2520Allowance%2FAllItems%2Easpx&ContentTypeId=0x01004A2D27B1B3D9A946B497C4750155A57B&RootFolder="
        
        # Print the URL to ensure it's correct
        print(f"Navigating to: {url}")

        # Open the webpage
        driver.get(url)

        # Refresh the page
        driver.refresh()

        # Wait for the page to load and the elements to be interactable
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Name Required Field']")))
        
        # Fill in the form fields using the provided locators
        driver.execute_script("document.evaluate(\"//input[@title='Name Required Field']\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[0];", name)
        driver.execute_script("document.evaluate(\"//input[@title='ERP ID Required Field']\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[0];", erp_id)
        driver.execute_script("document.evaluate(\"//input[@title='Line Manager Required Field']\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[0];", line_manager)
        driver.execute_script("document.evaluate(\"//textarea[contains(@title,'Rich text editor Project Details Required Field')]\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[0];", project_details)
        driver.execute_script("document.evaluate(\"//input[@title='Work Date Required Field']\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[0];", work_date.strftime('%m-%d-%Y'))
        
        # Select the shift type from the dropdown
        select = Select(driver.find_element(By.ID, 'Type_x0020_of_x0020_On_x0020_Cal_b12b720b-e12e-4a46-8600-d4074608dbdb_$DropDownChoice'))
        select.select_by_visible_text(shift_type)
        
        # Don't submit the form (commented out the submit button click)
        # driver.find_element(By.ID, 'submit_form').click()

        # Wait to observe the filled form (adjust as needed)
        time.sleep(5)

    # Loop through each row in the DataFrame to fill the form
    for index, row in df.iterrows():
        fill_form(row)

except Exception as e:
    print(f"Exception occurred: {str(e)}")

finally:
    # Close the WebDriver
    driver.quit()
