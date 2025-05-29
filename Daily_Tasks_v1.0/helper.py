import os
import subprocess
import shutil
import webbrowser
import pyautogui
import threading
import win32com.client as win32
import psutil
import time
import config
from pywinauto import Application
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, ElementClickInterceptedException, WebDriverException
from colorama import Fore, Style
import platform

"""
----------------------------------------------
The function below is to focus on the Python script window
----------------------------------------------
"""
def focus_python_script()-> None:
    """ Function to focus on the Python script window."""
    try:
        # Connect to the Python script window (adjust title based on your terminal/editor)
        app = Application().connect(path="python.exe")  # Matches any window with "Python" in the title

        # Find the first matching window
        window = app.top_window()

        # Bring the window to focus
        window.set_focus()

    except Exception as e:
        print(f"⚠️ Error: {e} ⚠️")


"""
----------------------------------------------
The functions below is to ask user for input
----------------------------------------------
"""
def get_input_with_timeout(prompt: str, timeout: int) -> str | None:
    user_input = [None]

    def input_thread():
        try:
            user_input[0] = input(prompt).strip().lower()
        except EOFError:
            pass

    thread = threading.Thread(target=input_thread, daemon=True)
    thread.start()

    thread.join(timeout)
    
    if thread.is_alive():
        focus_python_script()
        pyautogui.press('y')  # Simulates pressing Enter
        pyautogui.press('enter')

        thread.join(1)  # Give the thread a moment to finish

    return user_input[0]

def proceed(prompt:str) -> bool:
    """ Function to get user input (yes/no) and return True or False based on input. 
    Args:
        prompt (str): The prompt to display to the user.
    Returns:
        bool: True if input is yes, False if input is no.
    """
    input_value = input(prompt).strip().lower()  # Convert to lowercase for easier checking
    while input_value not in ['y', 'n', "yes", "no"]:
        print("Invalid input. Please enter 'y' or 'n'.")
        input_value = input(prompt).strip().lower()
    return input_value in ['y', 'yes'] # Return True for 'y' or 'yes', False for 'n' or 'no'

"""
--------------------------------------------
The function below is to print out the tasks
---------------------------------------------
"""
def print_tasks(day_of_week:str, tasks:dict) -> None:
    """ Function to print tasks for the given day of the week.
    Args:
        day_of_week (str): The day of the week.
        tasks (dict): A dictionary containing tasks for each day of the week."""
    n = 1
    try:
        if tasks[day_of_week]:
            print(Fore.BLUE + f"\n📍 TASK FOR {day_of_week.upper()} 📍")
            print(Style.RESET_ALL, end="")
            for task in tasks[day_of_week]:
                print(Fore.CYAN + f"{n}.", task[0])
                n+=1
            print(Style.RESET_ALL, end="")
        else:
            print("\n⚠️ No task for today? ⚠️")
    except KeyError:
        print(f"⚠️ Error: '{day_of_week}' is not a valid key in tasks. ⚠️")
    except Exception as e:
        print(f"⚠️ Unexpected error in print_tasks: {e} ⚠️")
        
"""
--------------------------------------------------------------
The functions below are to open applications and files
---------------------------------------------------------------
"""
def open_excel_file(excelfilepath:str) -> None:
    """ Function to open the daily IBD changes file in Excel and a web browser. 
    Args:
        dailyibdfile (str): The path to the daily IBD changes file.
    """
    try:
        # Initialize the Excel application
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = True  # Set to True to make Excel visible
        excel.Workbooks.Open(excelfilepath)
        filename = os.path.basename(excelfilepath)
        focus_python_script()  # Focus on the Python script window
        print(f"\n📄 Excel file opened: {filename}")
    except FileNotFoundError as fnf:
        print(f"⚠️ {fnf} ⚠️")
    except AttributeError as ae:
        print(f"⚠️ Attribute error: {ae} ⚠️")
    except Exception as e:
        print(f"⚠️ Error: {e} ⚠️")

def open_outlook() -> None:
    """ Function to open Outlook if it's not already running.
    """
    outlook_path = config.outlook_path
    def is_outlook_running():
        try:
            for process in psutil.process_iter(['name']):
                if process.info['name'] and "OUTLOOK.EXE" in process.info['name']:
                    return True
            return False
        except Exception as e:
            print(f"⚠️ Error checking Outlook process: {e} ⚠️")
            return False
    print("Opening up outlook to ensure email is sent 📨📬")
    try:
        if not os.path.exists(outlook_path):
            raise FileNotFoundError(f"Outlook executable not found: {outlook_path}")
        if not is_outlook_running():
            subprocess.Popen([outlook_path])
    except PermissionError as pe:
        print(f"⚠️ Permission error: {pe} ⚠️")
    except Exception as e:
        print(f"⚠️ Error in open_outlook: {e} ⚠️")

def reinitiate_sap() -> None:
    """Function to terminate and reopen the SAP application process if it is found running
    """
    # Name of the SAP application process, usually 'saplogon.exe'
    sap_process_name = "saplogon.exe"
    sap_path = config.sap_path

    print(Fore.MAGENTA+"\n💡 Reinitiating SAP application...")
    # Loop through all running processes and terminate the SAP application if found
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            # If process name matches SAP
            if proc.info['name'].lower() == sap_process_name:
                proc.terminate()  # Send terminate signal
                proc.wait(timeout=5)  # Wait for process to close
                print(f"  SAP application has been found and terminated.")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    try:
        subprocess.Popen([sap_path])  # Open SAP logon only
        focus_python_script()
        print("  SAP application has been reinitiated." + Style.RESET_ALL)
    except Exception as e:
        print(f"⚠️ Error: {e} ⚠️")

"""
--------------------------------------------
The function below is to close an excel file
--------------------------------------------
"""
def close_excel_file(excelfilename: str, save_changes: bool = True) -> bool:
    """ Function to close the daily IBD changes file in Excel.
    Args:
        excelfilepath (str): excel file name for file to be closed, you can provide a section of the file name for python to identify the target (do not put the suffix eg: .xls or .xlsx or xlsm)
        save_changes (bool): Whether to save changes before closing. Default is True.
    Returns:
        bool: True if the file was closed successfully
    """
    try:
        # Initialize the Excel application
        found = False
        excel = win32.Dispatch('Excel.Application')
        excel.DisplayAlerts = False  # Disable alerts to avoid prompts
        while not found:
            for wb in excel.Workbooks:
                wbname = wb.Name
                if excelfilename.lower() in wbname.lower():
                    wb.Close(SaveChanges=save_changes)
                    found = True
                    print(f"\n❎ Excel file closed: {wbname}. Save changes: {save_changes}")
                    excel.DisplayAlerts = True # Re-enable alerts
                    return found
        if not found:
            print(f"⚠️ Excel file '{excelfilename}' not found. Please check the filename. ⚠️")
    except (AttributeError, TypeError) as e:
        print(f"⚠️ Error accessing Excel application: {e} ⚠️")
        return False
    except Exception as e:
        print(f"⚠️ Error: {e} ⚠️")
        return False

"""
------------------------------------------------------------
The function below is to run the macro in specified workbook
------------------------------------------------------------
"""

def run_excel_macro(
    excel_file_path: str,
    macro_name: str,
    close_workbook: bool = True,
    save_changes: bool = True,
    retry: bool = False,
    quit_excel:bool = False,
    reinitiate_sap_before: bool = False, # Option to restart SAP before running
) -> bool:
    """ Function to open and run a macro in the specified Excel workbook.
    Args:
        excel_file_path (str): The full path of the Excel file.
        macro_name (str): The name of the macro to run.
        close_workbook (bool): Whether to close the workbook after running the macro. Default is True.
        save_changes (bool): Whether to save changes before closing. Default is True.
        retry (bool): Whether to retry is macro fails
        quit_excel (bool) : Whether to quit excel after macro runs
        reinitiate_sap_before (bool): Whether to reinitiate SAP before running the macro. Default is False.
    Returns:
        bool: True if the macro ran successfully and status is as expected, False otherwise.
    """

    excel_filename = os.path.basename(excel_file_path)     
    open = False
    while not open:
        try:
            # Initialize Excel application and open the workbook
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = True  # Set to True to make Excel visible
            wb = excel.Workbooks.Open(excel_file_path, UpdateLinks=3)
            print(f"\n📄 Excel file opened: {excel_filename}")
            open = True
        except FileNotFoundError:
            print(f"\n⚠️ File not found: {excel_file_path} ⚠️")
            if retry:
                if proceed(f"⚠️ Excel failed to open, please check the {excel_filename} file. Do you want to retry? (yes/no): "):
                    continue
                else:
                    print("Exiting task...")
                    return False
        except PermissionError:
            print(f"\n⚠️ Permission denied opening: {excel_file_path} ⚠️")
            if retry:
                if proceed(f"⚠️ Excel failed to open, please check the {excel_filename} file. Do you want to retry? (yes/no): "):
                    continue
            else:
                print("Exiting task...")
                return False
        except Exception as e:
            print(f"\n⚠️ Error: {e} ⚠️")
            if retry:
                if proceed(f"⚠️ Excel failed to open, please check the {excel_filename} file. Do you want to retry? (yes/no): "):
                    continue
                else:
                    print("Exiting task...")
                    return False

    
    # Execute the Macro and repeat process if macro fails and user wants to retry, else return False
    while True:
        try:
            if reinitiate_sap_before:
                reinitiate_sap()
            print(f"\n🚀 Running Macro: {macro_name}")
            excel.Application.Run(macro_name)

        except Exception as e:
            print(f"\n⚠️ Error: {e} ⚠️")
            if retry:
                if proceed(f"⚠️ Macro has failed, please check the {excel_filename} file. Do you want to retry running the Macro? (yes/no): "):
                    continue
                else:
                    print("Exiting task...")
                    return False
        try:
            print(f"🎯 Macro Completed")
            # Close the worbook and save changes if specified by user
            if close_workbook:
                wb.Close(SaveChanges=save_changes)
                print("❎ Workbook closed. Save changes:", save_changes)
            if quit_excel:
                excel.Quit()
        except Exception as e:
            print(f"\n⚠️ Error: {e} ⚠️")
            return False
        finally:
            return True

"""
---------------------------------------
The functions below are to manage files
1. find and/or copy file
2. find and/or move file
3. find and/or delete file
---------------------------------------
"""
def rename_file(filepath:str, new_name:str) -> bool:
    """Action to rename a file.
    Args:
        filepath (str): Full path of the file to rename.
        new_name (str): New name for the file.
    Returns:
        bool: True if rename was successful, False otherwise.
    """
    try:
        directory = os.path.dirname(filepath)
        new_path = os.path.join(directory, new_name)
        os.rename(filepath, new_path)
        print(f"  ➜ Renamed '{filepath}' to '{new_path}'")
        return True
    except FileNotFoundError:
        print(f"⚠️ Source file '{filepath}' not found.")
        return False
    except PermissionError as pe:
        print(f"⚠️ Permission denied renaming '{filepath}': {pe}")
        return False
    except Exception as e:
        print(f"⚠️ Error renaming '{filepath}': {e}")
        return False
    
def copy_file(initial_file_path: str, dest_folder: str) -> bool:
    """Action to copy a single file to the destination folder.
    Args:
        initial_file_path (str): initial filepath of the target file
        dest_folder (str): destination folder path to be copied into
    Returns:
        bool: True if file copied, else False
    """
    filename = os.path.basename(initial_file_path)
    destination_path = os.path.join(dest_folder, filename)
    try:
        os.makedirs(dest_folder, exist_ok=True)
        if os.path.exists(destination_path):
            print(f"⚠️ Warning: File '{filename}' already exists in '{dest_folder}'. Replacing file.")
            try:
                os.remove(destination_path)
            except PermissionError as pe:
                print(f"⚠️ Permission denied removing existing file: {pe}")
                return False
            except Exception as e:
                print(f"⚠️ Error removing existing file: {e}")
                return False
        shutil.copy2(initial_file_path, destination_path)
        print(f"  ➜ Copied '{filename}' to '{dest_folder}'")
        return True
    except FileNotFoundError:
        print(f"⚠️ Source file '{filename}' not found.")
        return False
    except PermissionError as pe:
        print(f"⚠️ Permission denied copying '{filename}': {pe}")
        return False
    except Exception as e:
        print(f"⚠️ Error copying '{filename}': {e}")
        return False

def delete_file(file_path: str) -> bool:
    """Action to delete a single file.
    Args:
        file_path (str): filepath of the target file
    Returns:
        bool: True if file deleted, else False
    """
    filename = os.path.basename(file_path)
    try:
        os.remove(file_path)
        print(f"  🗑️ Deleted '{filename}'")
        return True
    except FileNotFoundError:
        print(f"⚠️ File '{filename}' not found for deletion.")
        return False
    except PermissionError as pe:
        print(f"⚠️ Permission denied deleting '{filename}': {pe}")
        return False
    except Exception as e:
        print(f"⚠️ Error deleting '{filename}': {e}")
        return False
   
def delete_all_files_in_folder(folderpath: str) -> bool:
    """Action to delete al files in a folder.
    Args:
        folderpath (str): folder path of target folder
    Returns:
        bool: True if all files deleted, else False
    """
    print("\n🗑️ Deleting remaining files")
    try:
        files = os.listdir(folderpath)
    except FileNotFoundError:
        print(f"⚠️ Folder '{folderpath}' not found.")
        return False
    except PermissionError as pe:
        print(f"⚠️ Permission denied accessing folder '{folderpath}': {pe}")
        return False
    except Exception as e:
        print(f"⚠️ Error accessing folder '{folderpath}': {e}")
        return False
    error = False
    for file in files:
        file_path = os.path.join(folderpath, file)
        if os.path.isfile(file_path):
            try:
                os.remove(file_path)
                print(f"🗑️ Deleted: {file_path}")
            except FileNotFoundError:
                print(f"\n⚠️ File '{file_path}' not found for deletion. ⚠️")
                error = True
            except PermissionError as pe:
                print(f"\n⚠️ Permission denied deleting '{file_path}': {pe} ⚠️")
                error = True
            except Exception as e:
                print(f"\n⚠️ Error deleting file {file_path}: {e} ⚠️")
                error = True
    if error:
        return False
    return True
    
def move_file(initial_file_path: str, dest_folder: str) -> bool:
    """Action to move a single file to the destination folder.
    Args:
        initial_file_path (str): Full path of the file to move.
        dest_folder (str): Full path of the destination folder.
    Returns:
        bool: True if move was successful, False otherwise.
    """
    filename = os.path.basename(initial_file_path)
    destination_path = os.path.join(dest_folder, filename)
    try:
        # Ensure destination exists
        os.makedirs(dest_folder, exist_ok=True)
         # Check if file already exists at destination
        if os.path.exists(destination_path):
            print(f"⚠️ Warning: File '{filename}' already exists in '{dest_folder}'. Replacing file.")
            os.remove(destination_path)

        shutil.move(initial_file_path, destination_path)
        print(f"  ➜ Moved '{filename}' to '{dest_folder}'")
        return True
    except PermissionError as pe:
        print(f"\n⚠️ Permission denied for '{filename}': {pe} ⚠️")
        return False
    except Exception as e:
        print(f"⚠️ Error moving '{filename}': {e}")
        return False

def find_and_rename_file(source_folder: str, new_name: str, startswith: str, min_filename_length: int = None, compulsory_pattern: str = None, poll_interval: int = 5) -> str:
    """Continuously monitors the source folder until a file matching the criteria is found, then renames it and stops.
    Args:
        source_folder (str): Folder to monitor for the file.
        new_name (str): New name for the file.
        startswith (str): Compulsory prefix for the filename (case-insensitive).
        min_filename_length (int, optional): Minimum length of the filename. Defaults to None.
        compulsory_pattern (str, optional): A pattern that must be present in the filename. Defaults to None.
        poll_interval (int, optional): Seconds to wait between checking the folder. Defaults to 5.
    Returns:
        Str: Function returns filename that has been found and copied
    """
    print(f"\n🔎 Checking for file in '{source_folder}' that starts with '{startswith}' to rename to '{new_name}'...")
    found_and_processed = False

    while not found_and_processed:
        try:
            if not os.path.isdir(source_folder):
                print(f"⚠️ Source folder '{source_folder}' not found. Retrying in {poll_interval} seconds...")
                time.sleep(poll_interval)
                continue # Skip to next iteration

            for filename in os.listdir(source_folder):
                file_path = os.path.join(source_folder, filename)

                # Check if it's a file (and not a directory)
                if not os.path.isfile(file_path):
                    continue

                filenamelower = filename.lower()
                startswith_lower = startswith.lower()

                # Check criteria
                startswith_matches = filenamelower.startswith(startswith_lower)
                length_matches = min_filename_length is None or (len(filename) >= min_filename_length)
                pattern_matches = compulsory_pattern is None or (compulsory_pattern in filename)

                if startswith_matches and length_matches and pattern_matches:
                    print(f"  ✔️ Found matching file: {filename}")
                    if rename_file(file_path, new_name):
                        print(f"🎯 Renaming complete for '{filename}'. Stopping search.")
                        found_and_processed = True
                        break # Exit inner loop (listdir)
                    else:
                        print(f"   Retry will occur on next scan cycle.")

            if found_and_processed:
                return filename

        except PermissionError:
             print(f"⚠️ Permission denied accessing '{source_folder}'. Retrying in {poll_interval} seconds...")
        except Exception as e:
            print(f"⚠️ An unexpected error occurred while scanning '{source_folder}': {e}. Retrying in {poll_interval} seconds...")

        # Wait before the next scan, only if a file wasn't processed
        if not found_and_processed:
            time.sleep(poll_interval)

def find_and_copy_file(source_folder: str, destination_folder: str, startswith: str, min_filename_length: int = None, compulsory_pattern: str = None, poll_interval: int = 5) -> str:
    """Continuously monitors the source folder until a file matching the criteria is found, then copies it to the destination folder and stops.
    Args:
        source_folder (str): Folder to monitor for the file.
        destination_folder (str): Folder to copy the file to.
        startswith (str): Compulsory prefix for the filename (case-insensitive).
        min_filename_length (int, optional): Minimum length of the filename. Defaults to None.
        compulsory_pattern (str, optional): A pattern that must be present in the filename. Defaults to None.
        poll_interval (int, optional): Seconds to wait between checking the folder. Defaults to 5.
    Returns:
        Str: Function returns filename that has been found and copied
    """
    print(f"\n🔎 Checking for file in '{source_folder}' that starts with '{startswith}' to copy to '{destination_folder}'...")
    found_and_processed = False

    while not found_and_processed:
        try:
            if not os.path.isdir(source_folder):
                print(f"⚠️ Source folder '{source_folder}' not found. Retrying in {poll_interval} seconds...")
                time.sleep(poll_interval)
                continue # Skip to next iteration

            for filename in os.listdir(source_folder):
                file_path = os.path.join(source_folder, filename)

                # Check if it's a file (and not a directory)
                if not os.path.isfile(file_path):
                    continue

                filenamelower = filename.lower()
                startswith_lower = startswith.lower()

                # Check criteria
                startswith_matches = filenamelower.startswith(startswith_lower)
                length_matches = min_filename_length is None or (len(filename) >= min_filename_length)
                pattern_matches = compulsory_pattern is None or (compulsory_pattern in filename)

                if startswith_matches and length_matches and pattern_matches:
                    print(f"  ✔️ Found matching file: {filename}")
                    if copy_file(file_path, destination_folder):
                        print(f"🎯 Copy complete for '{filename}'. Stopping search.")
                        found_and_processed = True
                        break # Exit inner loop (listdir)
                    else:
                        print(f"   Retry will occur on next scan cycle.")

            if found_and_processed:
                return filename

        except PermissionError:
             print(f"⚠️ Permission denied accessing '{source_folder}'. Retrying in {poll_interval} seconds...")
        except Exception as e:
            print(f"⚠️ An unexpected error occurred while scanning '{source_folder}': {e}. Retrying in {poll_interval} seconds...")

        # Wait before the next scan, only if a file wasn't processed
        if not found_and_processed:
            time.sleep(poll_interval)

def find_and_delete_file(source_folder: str, startswith: str, min_filename_length: int = None, compulsory_pattern: str = None, poll_interval: int = 5) -> str:
    """Continuously monitors the source folder until a file matching the criteria is found, then deletes it and stops.

    Args:
        source_folder (str): Folder to monitor for the file to delete.
        startswith (str): Compulsory prefix for the filename (case-insensitive).
        min_filename_length (int, optional): Minimum length of the filename. Defaults to None.
        compulsory_pattern (str, optional): A pattern that must be present in the filename. Defaults to None.
        poll_interval (int, optional): Seconds to wait between checking the folder. Defaults to 5.
    Returns:
        Str: Function returns filename that has been found and deleted
    """
    print(f"\n🔎 Checking for file in '{source_folder}' that starts with '{startswith}' to delete...")
    found_and_processed = False

    while not found_and_processed:
        try:
            if not os.path.isdir(source_folder):
                print(f"⚠️ Source folder '{source_folder}' not found. Retrying in {poll_interval} seconds...")
                time.sleep(poll_interval)
                continue

            for filename in os.listdir(source_folder):
                file_path = os.path.join(source_folder, filename)

                if not os.path.isfile(file_path):
                    continue

                filenamelower = filename.lower()
                startswith_lower = startswith.lower()

                # Check criteria
                startswith_matches = filenamelower.startswith(startswith_lower)
                length_matches = min_filename_length is None or (len(filename) >= min_filename_length)
                pattern_matches = compulsory_pattern is None or (compulsory_pattern in filename)

                if startswith_matches and length_matches and pattern_matches:
                    print(f"  ✔️ Found matching file: {filename}")
                    if delete_file(file_path):
                        print(f"🎯 Deletion complete for '{filename}'. Stopping search.")
                        found_and_processed = True
                        break # Exit inner loop (listdir)
                    else:
                        # Delete failed, print error (handled in delete_file_action)
                        print(f"   Retry will occur on next scan cycle.")

            if found_and_processed:
                return filename

        except PermissionError:
             print(f"⚠️ Permission denied accessing '{source_folder}'. Retrying in {poll_interval} seconds...")
        except Exception as e:
            print(f"⚠️ An unexpected error occurred while scanning '{source_folder}': {e}. Retrying in {poll_interval} seconds...")

        if not found_and_processed:
            time.sleep(poll_interval)

def find_and_move_file(source_folder: str, destination_folder: str, startswith: str, min_filename_length: int = None, compulsory_pattern: str = None, poll_interval: int = 5) -> str:
    """Continuously monitors the source folder until a file matching the criteria is found, then moves it to the destination folder and stops.
    Args:
        source_folder (str): Folder to monitor for the file.
        destination_folder (str): Folder to move the file to.
        startswith (str): Compulsory prefix for the filename (case-insensitive).
        min_filename_length (int, optional): Minimum length of the filename. Defaults to None.
        compulsory_pattern (str, optional): A pattern that must be present in the filename. Defaults to None.
        poll_interval (int, optional): Seconds to wait between checking the folder. Defaults to 5.
    Returns:
        Str: Function returns filename that has been found and moved
    """
    print(f"\n🔎 Checking for file in '{source_folder}' that starts with '{startswith}' to move to '{destination_folder}'...")
    found_and_processed = False

    while not found_and_processed:
        try:
            if not os.path.isdir(source_folder):
                print(f"⚠️ Source folder '{source_folder}' not found. Retrying in {poll_interval} seconds...")
                time.sleep(poll_interval)
                continue

            for filename in os.listdir(source_folder):
                file_path = os.path.join(source_folder, filename)

                if not os.path.isfile(file_path):
                    continue

                filenamelower = filename.lower()
                startswith_lower = startswith.lower()

                # Check criteria
                startswith_matches = filenamelower.startswith(startswith_lower)
                length_matches = min_filename_length is None or (len(filename) >= min_filename_length)
                pattern_matches = compulsory_pattern is None or (compulsory_pattern in filename)

                if startswith_matches and length_matches and pattern_matches:
                    print(f"  ✔️ Found matching file: {filename}")
                    if move_file(file_path, destination_folder):
                        print(f"🎯 Move complete for '{filename}'. Stopping search.")
                        found_and_processed = True
                        break # Exit inner loop (listdir)
                    else:
                         # Move failed, print error (handled in move_file_action)
                        print(f"   Retry will occur on next scan cycle.")


            if found_and_processed:
                return filename

        except PermissionError:
             print(f"⚠️ Permission denied accessing '{source_folder}'. Retrying in {poll_interval} seconds...")
        except Exception as e:
             print(f"⚠️ An unexpected error occurred while scanning '{source_folder}': {e}. Retrying in {poll_interval} seconds...")

        if not found_and_processed:
             time.sleep(poll_interval)

"""
---------------------------------------
The class below is to interact with web elements, it contains the following methods
1. find and click buttons
2. find and type into an element
3. wait for an element to disappear (typically used to detect loading)
4. Open up a website on your default browser
---------------------------------------
"""

class WebAutomation:
    def __init__(self, browser_type:str = "chrome"):
        """Initializes the WebAutomation class with a website link and optional parameters.
        Args:   
            browser_type(str): Enter the browser type for your driver. Supported browsers are chrome(chrome), safari(safari) and microsoft edge(edge). Default is Chrome
        """
        self.browser_type = browser_type.lower()

        if not hasattr(config, 'driver_path') or not config.driver_path:
            raise NameError("Edge driver path (edge_driver_path) not defined in config.")
        service = Service(config.driver_path)

        try:
            if self.browser_type == "edge":
                self.driver = webdriver.Edge(service=service)
                print("✅ Microsoft Edge WebDriver initialized successfully.")

            elif self.browser_type == "chrome":
                self.driver = webdriver.Chrome(service=service)
                print("✅ Google Chrome WebDriver initialized successfully.")

            elif self.browser_type == "safari":
                if platform.system() != "Darwin": # Darwin is macOS
                    print("⚠️ Safari WebDriver is primarily for macOS. Initialization might fail on other OS.")
                    self.driver = webdriver.Safari()
                print("✅ Apple Safari WebDriver initialized successfully.")
                print("   Ensure 'Allow Remote Automation' is enabled in Safari's Develop menu on macOS.")
            else:
                print(f"⚠️ Unsupported browser type: '{browser_type}'. Please use 'edge', 'chrome', or 'safari'.")
                raise ValueError(f"Unsupported browser type: '{browser_type}'.")

        except NameError as ne:
            print(f"⚠️ Configuration Error: {ne}")
            self.driver = None
        except WebDriverException as wde:
            driver_path_to_show = "default system path"
            if self.browser_type == "edge":
                driver_path_to_show = getattr(config, 'driver_path', 'N/A')
            elif self.browser_type == "chrome":
                driver_path_to_show = getattr(config, 'driver_path', 'N/A')
            elif self.browser_type == "safari":
                 driver_path_to_show = getattr(config, 'driver_path', 'default system path')
            print(f"⚠️ WebDriver Error for {self.browser_type.capitalize()}: {wde}")
            print(f"   Please ensure the driver at '{driver_path_to_show}' is correct, executable, and its version matches your browser version.")
            if self.browser_type == "safari" and platform.system() == "Darwin":
                 print("   For Safari, also ensure 'Allow Remote Automation' is enabled in Develop menu and `safaridriver` is working (`safaridriver --enable`).")
            self.driver = None
        except Exception as e:
            print(f"⚠️ An unexpected error occurred initializing {self.browser_type.capitalize()} WebDriver: {e}")
            self.driver = None

    def goto(self, web_url: str) -> bool:
        """Navigate to a URL
        Args:
            web_url(str): URL to your targetted website
        Returns:
            bool: Returns True if successful, False otherwise
        """
        if not self.driver:
            print("⚠️ WebDriver not initialized. Cannot navigate to URL.")
            return
        try:
            self.driver.get(web_url)
            print(f"✅ Navigated to URL: {web_url[0:10]}...")
            return True
        except Exception as e:
            print(f"⚠️ Error navigating to URL: {e}")
            return False
        
    def quit(self) -> bool:
        """Quit the webdriver"""
        if not self.driver:
            return
        try:
            self.driver.quit()
            print("✅ WebDriver quit successfully.")
            self.driver = None
            return True
        except Exception as e:
            print(f"⚠️ Error quitting WebDriver: {e}")
            return False
        
    def wait_for_element_to_disappear(self, element:str) -> bool:
        """Detect an element and return True when the target button disappears.
        Args:
            element(str): a string containing the target html element. Navigates using XPATH.
        Returns:
            bool: Return True when button has disappeared, false is error occurs
        """
        try:
            print(f"⏳ Waiting for target element to disappear")
            while True:
                buttons = self.driver.find_elements(By.XPATH, element)
                if len(buttons) == 0:
                    break  # Exit the loop when all buttons are gone
                time.sleep(3)  # Short delay to avoid excessive checks
            
            print("✅ Element has disappeared. Proceeding with the script...")
            return True
        except Exception as e:
            print(f"⚠️ An unexpected error occurred: {e}")
            return False

    def find_and_type_into_web_element(self, element: str, text_to_type: str, clear_first: bool = True, timeout: int = 30) -> bool:
        """
        Waits for a WebElement to be visible, optionally clears it, and types text into it.

        Args:
            driver: The Selenium WebDriver instance.
            element(str): The Selenium WebElement (e.g., input, textarea) to type into.
            text_to_type: The string to type into the element.
            clear_first: If True (default), clears the text in the element before typing.
            timeout: Maximum time in seconds to wait for the element to be visible. Default is 30 seconds. Increase if loading time is long.

        Returns:
            bool: True if text successfully typed in , False otherwise.
        """
        try:
            # Wait for the element to be present and visible
            wait = WebDriverWait(self.driver, timeout)
            visible_element = wait.until(EC.visibility_of_element_located((By.XPATH, element)))

            # Optional: Clear the field before typing
            if clear_first:
                visible_element.send_keys(Keys.CONTROL + "a")  # Select all text
                visible_element.send_keys(Keys.DELETE)  # Delete selected text

            # Perform the typing action
            visible_element.send_keys(text_to_type)
            print(f"✅ Successfully typed '{text_to_type}' into the element")
            return True
        except TimeoutException:
            print(f"⚠️ Error: Element was not found or not visible within {timeout} seconds.")
            return False
        except ElementNotInteractableException as e:
            print(f"⚠️ Error: Element found but could not be interacted with (might be disabled or not an input field). Details: {e}")
            return False
        except Exception as e:
            print(f"⚠️ An unexpected error occurred during typing: {e}")
            return False

    def find_and_click_web_element(self, element:str, timeout: int = 30)-> bool:
        """Waits for a WebElement to be clickable and then clicks it.
        Args:
            driver: The Selenium WebDriver instance.
            element (str): The Selenium WebElement to click.
            timeout: Maximum time in seconds to wait for the element to be clickable. Default is 30 seconds. Increase if loading time is long

        Returns:
            bool: True if the click was successful, False otherwise.
        """
        try:
            # Wait for the element to be present, visible, and clickable
            wait = WebDriverWait(self.driver, timeout)
            clickable_element = wait.until(EC.element_to_be_clickable((By.XPATH, element)))

            # Perform the click
            clickable_element.click()
            print(f"✅ Element clicked")
            return True
        except TimeoutException:
            print(f"⚠️ Error: Element was not found or not clickable within {timeout} seconds.")
            return False
        except (ElementNotInteractableException, ElementClickInterceptedException) as e:
            print(f"⚠️ Error: Element found but could not be clicked (might be obscured, disabled, or not interactive). Details: {e}")
            return False
        except Exception as e:
            print(f"⚠️ An unexpected error occurred during click: {e}")
            return False
        
    @staticmethod
    def open_website(link:str) -> None:
        """ Function to open a website in the default web browser. Use this when you want to open website without driver so that browser doesn't close when script ends or quit method is called.
        Args:
            link (str): The URL of the website to open.
        """
        try:
            webbrowser.open(link)
            focus_python_script()
            print(f"🌐 Website opened: {link}")
        except ValueError as ve:
            print(f"⚠️ {ve} ⚠️")
        except webbrowser.Error as we:
            print(f"⚠️ Webbrowser error: {we} ⚠️")
        except Exception as e:
            print(f"⚠️ Error: {e} ⚠️")
