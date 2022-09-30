import openpyxl
wb = openpyxl.load_workbook(r'input_octant_longest_subsequence.xlsx')
sheet = wb.active
row_count=sheet.max_row
total_count=row_count-1

# List to store octant signs
Octant_Sign_List = [1, -1, 2, -2, 3, -3, 4, -4]

#Help https://youtu.be/H37f_x4wAC0
from datetime import datetime
start_time = datetime.now()

def avg_calc():
    data_U=0
    data_V=0
    data_W=0
    for i in range(2, row_count + 1):
        try:
            data_U += sheet.cell(row=i,column=2).value
            data_V += sheet.cell(row=i,column=3).value
            data_W += sheet.cell(row=i,column=4).value
        except FileNotFoundError:
            print('File not found!')
    sheet['E1']='u_avg'
    sheet['F1']='v_avg'
    sheet['G1']='w_avg'
    # average of u, v, w
    try:
        u_avg=data_U/total_count
        v_avg=data_V/total_count
        w_avg=data_W/total_count
    except(ZeroDivisionError):
        print("No input data found!!\nDivision by zero is not allowed!")
        exit()

    try:
        #saving average of U in the sheet
        sheet['E2']=u_avg
        #saving average of V in the sheet
        sheet['F2']=v_avg
        #saving average of W in the sheet
        sheet['G2']=w_avg
    except FileNotFoundError:
        print("File not found!!")
        exit()
    except ValueError:
        print("Row or column values must be at least 1 ")
        exit()

    sheet['H1']='U-u_avg'
    sheet['I1']='V-v_avg'
    sheet['J1']='W-w_avg'
    for i in range(2, row_count + 1):
        #calculating and saving U-u_avg in the sheet
        sub_u_avg = sheet.cell(row=i,column=2).value-u_avg
        sheet.cell(row=i,column=8).value=sub_u_avg

        #calculating and saving V-v_avg in the sheet
        sub_v_avg = sheet.cell(row=i,column=3).value-v_avg
        sheet.cell(row=i,column=9).value=sub_v_avg

        #calculating and saving W-w_avg in the sheet
        sub_w_avg = sheet.cell(row=i,column=4).value-w_avg
        sheet.cell(row=i,column=10).value=sub_w_avg

def octant_identification():
    #function to calculate and save average value of U, V, W
    avg_calc()
    sheet['K1']='Octant'

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

octant_identification()
#saving the file in the given xlsx file
wb.save('output_octant_longest_subsequence.xlsx')

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
