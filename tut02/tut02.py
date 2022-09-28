import openpyxl

def octant_identification(mod):
    wb = openpyxl.load_workbook(r'input_octant_transition_identify.xlsx')
    sheet = wb.active
    row_count=sheet.max_row

    #for calculating average of U
    data_U=0
    for i in range(2, row_count + 1):
        data_U += sheet.cell(row=i,column=2).value
    print('\n')
    u_avg=data_U/row_count
    print('Average value of U: ',data_U/(row_count))
    sheet['E1']='u_avg'
    sheet['E2']=u_avg

    #for calculating average of V
    data_V=0
    for i in range(2, row_count + 1):
        data_V += sheet.cell(row=i,column=3).value
    print('\n')
    v_avg=data_V/row_count
    print('Average value of U: ',data_V/(row_count))
    sheet['F1']='v_avg'
    sheet['F2']=v_avg

    #for calculating average of W
    data_W=0
    for i in range(2, row_count + 1):
        data_W += sheet.cell(row=i,column=4).value
    print('\n')
    w_avg=data_W/row_count
    print('Average value of U: ',data_W/(row_count))
    sheet['G1']='w_avg'
    sheet['G2']=w_avg
    #saving the file in given xlsx file
    wb.save('output_octant_transition_identify.xlsx')

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)