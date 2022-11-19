from datetime import datetime
start_time = datetime.now()
import os
os.system("cls")

import openpyxl
import glob

from openpyxl.styles import Color, PatternFill, Font, Border, Side

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

##Read all the excel files in a batch format from the input/ folder. Only xlsx to be allowed
##Save all the excel files in a the output/ folder. Only xlsx to be allowed
## output filename = input_filename[_octant_analysis_mod_5000].xlsx , ie, append _octant_analysis_mod_5000 to the original filename.

octant_sign = [1,-1,2,-2,3,-3,4,-4]
octant_name_id_mapping = {1:"Internal outward interaction", -1:"External outward interaction", 2:"External Ejection", -2:"Internal Ejection", 3:"External inward interaction", -3:"Internal inward interaction", 4:"Internal sweep", -4:"External sweep"}
yellow = "00FFFF00"
yellow_bg = PatternFill(start_color=yellow, end_color= yellow, fill_type='solid')
black = "00000000"
double = Side(border_style="thin", color=black)
blackBorder = Border(top=double, left=double, right=double, bottom=double)

###Code
def entry_point(input_file, mod):
	fileName = input_file.split("\\")[-1]
	fileName = fileName.split(".xlsx")[0]
	outputFileName = "output/" + fileName + "_octant_analysis_mod_" + str(mod) + ".xlsx"

	outputFile = openpyxl.Workbook()
	outputSheet = outputFile.active

	outputSheet.cell(row=1, column=14).value = "Overall Octant Count"
	outputSheet.cell(row=1, column=24).value = "Rank #1 Should be highligted Yellow"
	outputSheet.cell(row=1, column=35).value = "Overall Transition Count"
	outputSheet.cell(row=1, column=45).value = "Longest Subsquence Length"
	outputSheet.cell(row=1, column=49).value = "Longest Subsquence Length with Range"
	outputSheet.cell(row=2, column=36).value = "To"

	headers = ["T", "U", "V", "W", "U Avg", "V Avg", "W Avg", "U'=U - U avg", "V'=V - V avg","W'=W - W avg", "Octant"]
	for i, header in enumerate(headers):
		outputSheet.cell(row=2, column=i+1).value = header
		
	outputFile.save(outputFileName)

def octant_analysis(mod=5000):
	path = os.getcwd()
	csv_files = glob.glob(os.path.join(path + "\input", "*.xlsx"))
	
	for file in csv_files:
		entry_point(file, mod)
mod=5000
octant_analysis(mod)

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
