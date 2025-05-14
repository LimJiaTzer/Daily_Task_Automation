import datetime
import os

"""
----------------------------------
Edit File locations here
----------------------------------
"""
outlook_path = r"" # Path to your outlook executable
sap_path = r"" # Path to your SAP executable

"""
-----------------------------------------------------------------------------------------------------------------------------------
Edit functions and daily task here
-> first drop the python file to the same folder as this file
-> import the function in the same way as the others: from filename import function_name
-> add the function name to the dictionary/list of tasks in the same way as the others using a tuple: ("task_name", function_name)
-> for end of day tasks, set the time to after which you want them to run, there is only one time stamp for now
-> do the exact opposite to remove a task
-----------------------------------------------------------------------------------------------------------------------------------
"""

from sample import taskname

# for daily tasks
tasks = {
    "Monday": [("task1", taskname), 
               ("task2", )],
    "Tuesday": [], 
    "Wednesday": [],
    "Thursday": [],
    "Friday": [],
    "Saturday": [],
    "Sunday": []
}

# for task after a certain time
# set your time here
hour = 17
minute = 15
# add your tasks here, it will be ran from left to right
lasttasksofday = [("longdayatwork", )]