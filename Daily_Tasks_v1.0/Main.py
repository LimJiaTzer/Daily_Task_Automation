import cowsay
import datetime
import sys
import time
from colorama import Fore, Style
import config
import helper

def main():
    """ Main function to display greeting, tasks, and ask user for input """
    """Main Parameters"""
    today = datetime.date.today()
    current_time = datetime.datetime.now()
    day_of_week = today.strftime('%A')


    # Determine greeting based on the time of day
    current_hour = current_time.hour
    if current_hour < 12:
        greeting = "good morning🌞"
    elif 12 <= current_hour < 16:
        greeting = "good afternoon🌤️"
    else:
        greeting = "good evening🌅"

    # Display greeting with Cowsay
    cowsay.cow(f"Hello👋, {greeting}, it is a {day_of_week}!")
    # Open Microsoft Outlook to ensure email can be sent
    helper.open_outlook()
    time.sleep(2)  # Allow the applications some time to open
    helper.focus_python_script()

    # run end of day tasks if the time is past the time specified in the config file
    if config.hour is not None and config.minute is not None:
        if current_time.hour >= config.hour and current_time.minute>=config.minute:
            timenow = f"{current_time.hour}:{current_time.minute}"
            answer = helper.get_input_with_timeout(f"\nThe time now is {timenow}. Do you want to run end of the day tasks (y/n)? [10s timeout]: ", 10)
            while answer not in ["yes", "y", "no", "n", None]:
                print("Invalid input, please enter yes or no")
                answer = helper.get_input_with_timeout("\nDo you want to run end of the day tasks (y/n)? [10s timeout]", 10)
            if answer in ['no', 'n']:
                print("Returning to main tasks...")
            else:
                print(Fore.GREEN + f"\n🌟🌟Starting end of day tasks🌟🌟"+ Style.RESET_ALL)
                y = 0
                for task_name, task_function in config.lasttasksofday:
                    print(Fore.GREEN + f"\n🌟🌟Starting task: {task_name}🌟🌟"+ Style.RESET_ALL)
                    if task_function() == False:
                        print(Fore.RED + f"\n❌❌Task '{task_name}' failed. Skipping...❌❌" + Style.RESET_ALL)
                        y+=1
                    else:
                        print(Fore.GREEN + f"\n✅✅Task '{task_name}' processed✅✅" + Style.RESET_ALL)
                if y > 0:
                    print(Fore.RED + f"\n⚠️⚠️ {y} task(s) failed. Please check the logs for details.⚠️⚠️" + Style.RESET_ALL)
                print("\nEnd Of Day Task completed")
                print("\nYou have reached the end of the day. Goodbye👋👋")
                input("Hit 'Enter' to exit...")
                sys.exit(0)

    # Print tasks for the current day
    helper.print_tasks(day_of_week, config.tasks)
    valid_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    
    task_number = None
    retry_count = 0
    # get user input to see which task(s) to perform with a timeout
    while True:
        # stops user if they retry excessively
        if retry_count == 6:
            print("\n\nUMM, YOU RETRIED 6 TIMES ALREADY. Exitting the script for you to think through first :(")
            input("Hit 'Enter' once or twice to exit...")
            sys.exit(0) # Exit the scipt


        user_response = helper.get_input_with_timeout(
    "\nDo you want to? (20s timeout)\n"
    f"1. Execute all tasks as listed  ➝  (Enter 'yes' or 'no')\n"
    f"2. Execute an individual task  ➝  (Enter task number as listed)\n"
    f"3. Change the day             ➝  (Enter the day name)\n"
    f"{Fore.YELLOW}By default, all tasks will be ran in the order as listed{Style.RESET_ALL}\n"
    "Your choice: ",
    20
    )
        if user_response is None:
            task_number = None
            print(f"\nNo input received. Proceeding with {day_of_week}'s tasks...")
            break  # Exit the loop and continue with the script

        elif user_response.lower() in ["yes", "y"]:
            print(f"\nProceeding with {day_of_week}'s tasks...")
            break  # Exit the loop and continue with the script

        elif user_response.lower() in ["no", "n"]:
            print("\nEnding the script...")
            input("Hit 'Enter' to exit...")
            sys.exit(0) # Exit the scipt

        elif user_response.isdigit():
            task_number = int(user_response) - 1
            if 0 <= task_number < len(config.tasks[day_of_week]):
                proceed = helper.proceed(f"\nYou have selected the task: {config.tasks[day_of_week][task_number][0]}. Do you wish to proceed (yes/no)? : ")
                if proceed:
                    break  # Exit the loop and continue with the script
                # if retry is a no, return to main while loop and ask user again for what they want to do
            else:
                print("⚠️ Invalid task number⚠️. Please enter a valid number as indicated in the list to select specific task.")

        else:
            for day in valid_days:
                if user_response.capitalize() in day and len(user_response) >= 2:
                    day_of_week = day
                    print(f"\nDay changed to {day_of_week}. Displaying tasks...")
                    helper.print_tasks(day_of_week, config.tasks)
                    task_number = None
                    break
            else:
                print("⚠️ Invalid response⚠️. Please enter a valid day of the week (at least first 2 letters).")

        retry_count+=1
        
    x = 0
    if task_number is None:
        # Run tasks if the day has tasks else exit the script
        if config.tasks[day_of_week]:
            # Run tasks with threading and skipping
            for task_name, task_function in config.tasks[day_of_week]:
                print(Fore.GREEN + f"\n🌟🌟Starting task: {task_name}🌟🌟"+ Style.RESET_ALL)
                if task_function() == False:
                    print(Fore.RED + f"\n❌❌Task '{task_name}' failed. Skipping...❌❌" + Style.RESET_ALL)
                    x+=1
                else:
                    print(Fore.GREEN + f"\n✅✅Task '{task_name}' processed✅✅" + Style.RESET_ALL)
            if x>0:
                print(Fore.RED + f"\n⚠️⚠️ {x} task(s) failed. Please check the logs for details.⚠️⚠️" + Style.RESET_ALL)
                print("Please proceed with the manual tasks if needed, goodbye👋👋")
            else:
                print("\nAll tasks for the day have been processed :)")
                print("\nPlease proceed with the manual tasks if needed, goodbye👋👋")
        else:
            print("\nNo task for today or today is not a working day? Might be working too hard ngl🤷.")
            input("Press enter to exit...")
            sys.exit(0)
    else:
        print(Fore.GREEN + f"\n🌟🌟Starting task: {config.tasks[day_of_week][task_number][0]}🌟🌟" + Style.RESET_ALL)
        config.tasks[day_of_week][task_number][1]()
        print(Fore.GREEN + f"\n✅✅{config.tasks[day_of_week][task_number][0]} task has been processed. :)✅✅" + Style.RESET_ALL)
        print("\nPlease proceed with the manual tasks if needed, goodbye👋👋")

    
    input("Press enter to exit...")
    sys.exit(0)

if __name__ == "__main__":
    main()

