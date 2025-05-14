import os
import re
import time
import datetime
import config
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import helper


def downloading() -> bool:
    downloaded = False
    while not downloaded:
        try:
            today = datetime.datetime.today()

            # Calculate the most recent Monday
            monday = today - datetime.timedelta(days=today.weekday())

            # Format the date as YYYY-MM-DD
            monday_date = monday.strftime("%Y-%m-%d")
            # Set up WebDriver
            edge_driver_path = r"" # ADD PATH TO YOUR WEBDRIVER
            service = Service(edge_driver_path)
            driver = webdriver.Edge(service=service)

            # Open the webpage
            driver.get("") # ADD YOUR SITE LINK
            wait = WebDriverWait(driver, 30) #timeout after a minute

            if not helper.find_and_click_web_element(driver, "//button[@data-test-id='sign-in-react__already-signed-in--continue-button']"):
                print("No login needed")
            
            helper.find_and_click_web_element(driver, "//input[@data-test-id='data-settings-react__filter-item-control-VariableShowAllTabs--picker-control--input--input-control' and @aria-label='Show All Tabs']")
            helper.find_and_click_web_element(driver, "//div[@data-test-id='data-settings-react__filter-item-control-VariableShowAllTabs--list-item' and @aria-selected='false']")
            helper.find_and_click_web_element(driver, "//button[@data-test-id='data-settings-react__open']")
            helper.find_and_click_web_element(driver, "//button[@data-test-id='workbook-react__tab-bar--overflow-menu-icon']")
            helper.find_and_click_web_element(driver, "//div[@data-test-id='workbook-react__tab-bar--basic-tab-bar--overflow-menu--list-item' and @aria-label='6']")
            helper.wait_for_element_to_disappear(driver, "//button[@data-test-id='worksheet-react__async-cancel-button']")
            helper.find_and_click_web_element(driver, "//button[@data-test-id='workbook-react__export-data-button' and @aria-label='ExportData']")
            helper.find_and_type_into_web_element(driver, "//input[@data-test-id='export-data-workbook-dialog__input-container--input-control']", f"Supply Analysis_{monday_date}", clear_first=True)
            helper.find_and_click_web_element(driver, "//button[@data-test-id='export-data-workbook-dialog__export-button']")
            
            helper.find_and_move_file(config.downloads_folder, config.supply_analysis_folder, "Supply Analysis")
            downloaded = True
        except Exception as e:
            print(f"⚠️ Error occurred: {e}")
            if helper.proceed("\nDo you want to retry downloading supply analysis file? (y/n): "):
                driver.quit()
                continue
            else:
                print("Exiting the task...")
                return False
    driver.quit()
    return True
    

if __name__ == "__main__":
    print(f"🌟🌟Starting task🌟🌟")
    if downloading():
        print("✅✅ Downloading supply analysis file completed✅✅")
        input("Hit 'Enter' to exit...")
    else:
        print("❌❌ Downloading supply analysis file failed ❌❌")
        input("Hit 'Enter' to exit...")
