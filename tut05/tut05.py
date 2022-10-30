
import openpyxl
wb = openpyxl.load_workbook(r'octant_input.xlsx')
sheet = wb.active
row_count=sheet.max_row
total_count=row_count-1
from datetime import datetime
start_time = datetime.now()
# List to store octant signs
Octant_Sign_List = [1, -1, 2, -2, 3, -3, 4, -4]

#Help https://youtu.be/N6PBd4XdnEw

def octant_range_names(mod=5000):
    
    octant_name_id_mapping = {"1":"Internal outward interaction", "-1":"External outward interaction", "2":"External Ejection", "-2":"Internal Ejection", "3":"External inward interaction", "-3":"Internal inward interaction", "4":"Internal sweep", "-4":"External sweep"}

###Code

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

def octant_identification(mod):
    if mod>30000:
        raise Exception('mod value should be less than or equal to 30000')
    #function to calculate and save average value of U, V, W
    avg_calc()
    
from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod=5000
try:
    octant_identification(mod)
except NameError:
    print('Either the function name is wrong or the function does not exists')
octant_range_names(mod)

wb.save('octant_output_ranking_excel.xlsx')

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
