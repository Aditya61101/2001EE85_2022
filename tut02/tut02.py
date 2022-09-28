import openpyxl

def checkOct(u, v, w):
    if u > 0:
        if v > 0:
            if w > 0:
                # this means (u,v) is +ve hence 1st quad and w is +ve so +1
                return 1
            else:
                # this means (u,v) is +ve hence 1st quad but w is -ve so -1
                return -1
        else:
            if w > 0:
                # this means u>0 and v<0 hence 4th quad and w is +ve so +4
                return 4
            else:
                # this means u>0 and v<0 hence 4th quad but w is -ve so -4
                return -4
    else:
        if v > 0:
            if w > 0:
                # this means u<0 and v>0 hence 4th quad and w is +ve so +2
                return 2
            else:
                # this means u<0 and v>0 hence 4th quad but w is -ve so -2
                return -2
        else:
            if w > 0:
                # this means u<0 and v<0 hence 4th quad and w is +ve so +3
                return 3
            else:
                # this means u<0 and v<0 hence 4th quad and w is -ve so -3
                return -3

def avg_calc(sheet,row_count):

    #for calculating average of U
    data_U=0
    for i in range(2, row_count + 1):
        data_U += sheet.cell(row=i,column=2).value
    u_avg=data_U/row_count
    print('Average value of U: ',data_U/(row_count))
    sheet['E1']='u_avg'
    sheet['E2']=u_avg

    #for calculating average of V
    data_V=0
    for i in range(2, row_count + 1):
        data_V += sheet.cell(row=i,column=3).value
    v_avg=data_V/row_count
    print('Average value of U: ',data_V/(row_count))
    sheet['F1']='v_avg'
    sheet['F2']=v_avg

    #for calculating average of W
    data_W=0
    for i in range(2, row_count + 1):
        data_W += sheet.cell(row=i,column=4).value
    w_avg=data_W/row_count
    print('Average value of U: ',data_W/(row_count))
    sheet['G1']='w_avg'
    sheet['G2']=w_avg
    sub_u_avg=0
    sub_v_avg=0
    sub_w_avg=0
    #saving U-u_avg...
    sheet['H1']='U-u_avg'
    for i in range(2, row_count + 1):
        sub_u_avg = sheet.cell(row=i,column=2).value-u_avg
        sheet.cell(row=i,column=8).value=sub_u_avg
    #saving V-v_avg...
    sheet['I1']='V-v_avg'
    for i in range(2, row_count + 1):
        sub_v_avg = sheet.cell(row=i,column=3).value-v_avg
        sheet.cell(row=i,column=9).value=sub_v_avg
    #saving W-w_avg...
    sheet['J1']='W-w_avg'
    for i in range(2, row_count + 1):
        sub_w_avg = sheet.cell(row=i,column=4).value-w_avg
        sheet.cell(row=i,column=10).value=sub_w_avg

def octant_identification(mod):
    wb = openpyxl.load_workbook(r'input_octant_transition_identify.xlsx')
    sheet = wb.active
    row_count=sheet.max_row

    #function to calculate and save average value of U, V, W
    avg_calc(sheet,sheet.max_row)

    #saving the sign of the octant
    sheet['K1']='Octant'
    for i in range(2,row_count+1):
        sub_u_avg=sheet.cell(row=i,column=8).value
        sub_v_avg=sheet.cell(row=i,column=9).value
        sub_w_avg=sheet.cell(row=i,column=10).value
        region=checkOct(sub_u_avg,sub_v_avg,sub_w_avg)
        sheet.cell(row=i,column=11).value=region



    
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