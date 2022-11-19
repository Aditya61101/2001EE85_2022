from datetime import datetime
start_time = datetime.now()

from platform import python_version
ver = python_version()

import openpyxl

import os
os.system("cls")

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


inputAttendance = "input_attendance.csv"
inputRegisteredFile = "input_registered_students.csv"

roll_to_name = {}
roll_attendance = {}
dates = []

def attendance_report():
    # Method to map roll num to name
    map_roll_to_num_func()

def map_roll_to_num_func():
    # Opening input file
    f = open(inputRegisteredFile, "r")

    # Reading label line
    f.readline()

    # Reading all lines of input
    all_lines = f.readlines()
    
    # Iterating all lines
    for line in all_lines:
        line = line.strip()
        list = line.split(",")

        rollNum = list[0].strip()
        name = list[1].strip()
        roll_to_name[rollNum] = name

    f.close()

attendance_report()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))