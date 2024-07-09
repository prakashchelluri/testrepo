#Script for submitting the oncall shift allowance.(Mad)
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.by import By
from contextlib import contextmanager


def apply_shift_allowance(driver):
    df = pd.read_excel(r'C:\Users\satv1214\Downloads\ShiftAllowanceData.xlsx', sheet_name='Data')
    data_list = df.to_dict(orient='records')
    for data in data_list:
        sheet_name = data['Sheet Name']
        df = pd.read_excel(r'C:\Users\satv1214\Downloads\ShiftAllowanceData.xlsx', sheet_name=sheet_name)
        oncall_details = df.to_dict(orient='records')
        for oncall in oncall_details:
            fill_details(driver, data, oncall)


def fill_details(driver, data, oncall_details):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//span[text()='New']")))
    driver.find_element(By.XPATH, ("//span[text()='New']")).click()
    driver.find_element(By.XPATH, ("//input[@title='Name Required Field']")).send_keys(data['Name'])
    driver.find_element(By.XPATH, ("//input[@title='ERP ID Required Field']")).send_keys(data['ERP ID'])
    driver.find_element(By.XPATH, ("//input[@title='Line Manager Required Field']")).send_keys(data['Line Manager'])
    driver.find_element(By.XPATH, ("//textarea[contains(@title,'Project Details Required Field')]")).send_keys(data['Project Details'])
    shift_field = 'On-Call-Option 2-Rostered to remain on call and had to respond to one or more calls' if oncall_details['On Call'] == 'Y' else 'On-Call-Option 1-Rostered to remain on call but did not get any calls'
    select = Select(driver.find_element(By.XPATH, ("//select[contains(@title,'Shift Required Field')]")))
    select.select_by_visible_text(shift_field)
    driver.find_element(By.XPATH, ("//input[contains(@title,'Work Date Required Field')]")).send_keys(oncall_details['Date'])
    driver.find_element(By.XPATH, ("//table[@class='ms-formtable']/following-sibling::table//input[@value='Save']")).click()
    
    time.sleep(5)

@contextmanager
def init_driver():
    #options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #driver = webdriver.Chrome(options=options)
    driver = webdriver.Chrome()
    driver.implicitly_wait(30)
    driver.delete_all_cookies()
    driver.get("https://sps.netcracker.com/sites/India/HR/Lists/On%20Call%20Allowance/AllItems.aspx")
    driver.refresh()
    try:
        yield driver
    finally:
        driver.quit()


if __name__ == "__main__":
    with init_driver() as driver:
        apply_shift_allowance(driver)