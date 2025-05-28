import datetime
import config
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

            webautomation = helper.WebAutomation(browser_type="edge")

            # Open the webpage
            webautomation.goto("")

            # clicking elements
            if not webautomation.find_and_click_web_element( "//button[@data-test-id='sign-in-react__already-signed-in--continue-button']"):
                print("No login needed")
            webautomation.find_and_click_web_element("//input[@data-test-id='data-settings-react__filter-item-control-VariableShowAllTabs--picker-control--input--input-control' and @aria-label='Show All Tabs']")

            webautomation.find_and_click_web_element("//button[@data-test-id='workbook-react__tab-bar--overflow-menu-icon']")

            webautomation.find_and_click_web_element("//div[@data-test-id='workbook-react__tab-bar--basic-tab-bar--overflow-menu--list-item' and @aria-label='6']")

            # waiting for element to disappear
            webautomation.wait_for_element_to_disappear("//button[@data-test-id='worksheet-react__async-cancel-button']")
            
            webautomation.find_and_click_web_element("//button[@data-test-id='workbook-react__export-data-button' and @aria-label='ExportData']")

            #entry into an element
            webautomation.find_and_type_into_web_element("//input[@data-test-id='export-data-workbook-dialog__input-container--input-control']", f"hello_{monday_date}", clear_first=True)
            
            webautomation.find_and_click_web_element("//button[@data-test-id='export-data-workbook-dialog__export-button']")
            
            # move the file downloaded before quitting the driver
            helper.find_and_move_file(config.source, config.destination, f"hello_{monday_date}")
            downloaded = True
        except Exception as e:
            print(f"⚠️ Error occurred: {e}")
            if helper.proceed("\nDo you want to retry downloading hello file? (y/n): "):
                webautomation.quit()
                continue
            else:
                print("Exiting the task...")
                return False
    webautomation.quit()
    return True
    

if __name__ == "__main__":
    print(f"🌟🌟Starting task🌟🌟")
    if downloading():
        print("✅✅ Downloading completed✅✅")
        input("Hit 'Enter' to exit...")
    else:
        print("❌❌ Downloading failed ❌❌")
        input("Hit 'Enter' to exit...")
