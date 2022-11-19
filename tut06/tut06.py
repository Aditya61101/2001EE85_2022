import os
import openpyxl
from platform import python_version
from datetime import datetime
start_time = datetime.now()

ver = python_version()


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


def roll_attendance_func():
    title = ["Date", "Roll", "Name", "Total Attendance Count",
             "Real", "Duplicate", "Invalid", "Absent"]

    for rollNum in roll_to_name.keys():
        outputFileName = "output/" + rollNum + ".xlsx"
        # outputFile.save(outputFileName)

        outputFile = openpyxl.Workbook()
        outputSheet = outputFile.active

        for i, word in enumerate(title):
            outputSheet.cell(row=1, column=i+1).value = word
        outputSheet.cell(row=2, column=2).value = rollNum
        outputSheet.cell(row=2, column=3).value = roll_to_name[rollNum]

        attendance = roll_attendance[rollNum]  # map of date -> array

        for i, date in enumerate(attendance.keys()):
            outputSheet.cell(row=3+i, column=1).value = date
            list = attendance[date]
            total = list[0]+list[1]+list[2]

            outputSheet.cell(row=3+i, column=4).value = total
            outputSheet.cell(row=3+i, column=5).value = list[0]
            outputSheet.cell(row=3+i, column=6).value = list[1]
            outputSheet.cell(row=3+i, column=7).value = list[2]

            if total == 0:
                outputSheet.cell(row=3+i, column=8).value = 1
            else:
                outputSheet.cell(row=3+i, column=8).value = 0

        outputFile.save(outputFileName)

def valid_day(date):
    # check date validity - 0 for monday and 3 for thursday
    # dt = '7-11-2022' - dummy date
    ans = datetime.strptime(date, '%d-%m-%Y').weekday()

    if ans == 0 or ans == 3:
        return True

    return False

def valid_time(time):
    # e.g., time = 15:30
    hr, min = time.split(":")
    hr = int(hr)
    min = int(min)
    if hr == 15 and min == 0:
        return True
    if hr == 14:
        return True
    return False

def attendance_count_func():
    # Opening input file
    f = open(inputAttendance, "r")

    f.readline()

    all_lines = f.readlines()

    for rolls in roll_to_name.keys():
        roll_attendance[rolls] = {}

    for line in all_lines:
        line = line.strip()
        timestamp, naming = line.split(",")
        date, time = timestamp.split(" ")
        rollNumber = naming[:8]

        if valid_day(date):
            if date not in dates:
                dates.append(date)

            if date not in roll_attendance[rollNumber]:
                roll_attendance[rollNumber][date] = [0, 0, 0]

            if valid_time(time):
                if (roll_attendance[rollNumber][date][0] == 0):
                    roll_attendance[rollNumber][date][0] = 1
                else:
                    roll_attendance[rollNumber][date][1] += 1

            else:
                roll_attendance[rollNumber][date][2] += 1

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

def attendance_report():
    # Method to map roll num to name
    map_roll_to_num_func()
    # Method to count attendance
    attendance_count_func()
    # Saving rollNum attendance
    roll_attendance_func()


attendance_report()

# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
