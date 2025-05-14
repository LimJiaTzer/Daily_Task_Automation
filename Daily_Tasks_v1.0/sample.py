import config
import helper

def taskname()->bool:
    helper.open_outlook
    excelfilename = helper.find_and_move_file(config.sourcefolder, config.destination_folder, startswith = "", min_filename_length="", compulsory_pattern="")

    excelfilepath = r"yourdirectory" + excelfilename
    if helper.run_excel_macro(excelfilepath, "mymacro", close_workbook=False, save_changes=True, retry = True, reinitiate_sap_before=True):
        return False
    
    helper.close_excel_file(config.excelfilepath2, save_changes=True)

    helper.find_and_copy_file(config.sourcefolder, config.destination_folder, startswith = "", min_filename_length="", compulsory_pattern="")
    helper.delete_file()
    return True


if __name__ == "__main__":
    print(f"🌟🌟Starting task🌟🌟")
    if taskname():
        print("\n✅✅ Task completed✅✅")
        input("Hit 'Enter' to exit...")
    else:
        print("\n❌❌ Task failed❌❌")
        input("Hit 'Enter' to exit...")