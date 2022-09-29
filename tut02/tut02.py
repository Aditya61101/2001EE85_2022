import openpyxl
wb = openpyxl.load_workbook(r'input_octant_transition_identify.xlsx')
sheet = wb.active
row_count=sheet.max_row

def count_in_range(mod):
    ini=2
    fill=4
    row_max=row_count
    while ini<row_max:

        # loop to initialize the cell value as 0
        for j in range(14,22):
            sheet.cell(row=fill,column=j).value=0

        # loop to fill the count of octant in appropriate cell
        for i in range(ini,mod+ini):
            if sheet.cell(row=i,column=11).value==1:
                sheet.cell(row=fill,column=14).value+=1
            elif sheet.cell(row=i,column=11).value==-1:
                sheet.cell(row=fill,column=15).value+=1
            elif sheet.cell(row=i,column=11).value==2:
                sheet.cell(row=fill,column=16).value+=1
            elif sheet.cell(row=i,column=11).value==-2:
                sheet.cell(row=fill,column=17).value+=1
            elif sheet.cell(row=i,column=11).value==3:
                sheet.cell(row=fill,column=18).value+=1
            elif sheet.cell(row=i,column=11).value==-3:
                sheet.cell(row=fill,column=19).value+=1
            elif sheet.cell(row=i,column=11).value==4:
                sheet.cell(row=fill,column=20).value+=1
            elif sheet.cell(row=i,column=11).value==-4:
                sheet.cell(row=fill,column=21).value+=1

        # fill range value in appropriate cell
        x1=str(ini-2)
        y1=str(min(row_max-2,mod+ini-3))
        sheet.cell(row=fill,column=13).value=x1+'-'+y1

        #increasing initial value for for loop by mod
        ini+=mod
        #increasing filling row by 1
        fill+=1

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

def avg_calc():

    #for calculating average of U
    data_U=0
    for i in range(2, row_count + 1):
        data_U += sheet.cell(row=i,column=2).value
    u_avg=data_U/row_count
    sheet['E1']='u_avg'
    sheet['E2']=u_avg

    #for calculating average of V
    data_V=0
    for i in range(2, row_count + 1):
        data_V += sheet.cell(row=i,column=3).value
    v_avg=data_V/row_count
    sheet['F1']='v_avg'
    sheet['F2']=v_avg

    #for calculating average of W
    data_W=0
    for i in range(2, row_count + 1):
        data_W += sheet.cell(row=i,column=4).value
    w_avg=data_W/row_count
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

    #function to calculate and save average value of U, V, W
    avg_calc()

    sheet['K1']='Octant'

    #initializing count values of each octant sign as 0
    for i in range(14,22):
        sheet.cell(row=2,column=i).value=0

    #saving the sign of the octant
    for i in range(2,row_count+1):
        sub_u_avg=sheet.cell(row=i,column=8).value
        sub_v_avg=sheet.cell(row=i,column=9).value
        sub_w_avg=sheet.cell(row=i,column=10).value
        region=checkOct(sub_u_avg,sub_v_avg,sub_w_avg)
        if region==1:
            sheet.cell(row=2,column=14).value+=1
        elif region==-1:
            sheet.cell(row=2,column=15).value+=1
        elif region==2:
            sheet.cell(row=2,column=16).value+=1
        elif region==-2:
            sheet.cell(row=2,column=17).value+=1
        elif region==3:
            sheet.cell(row=2,column=18).value+=1
        elif region==-3:
            sheet.cell(row=2,column=19).value+=1
        elif region==4:
            sheet.cell(row=2,column=20).value+=1
        else:
            sheet.cell(row=2,column=21).value+=1
        sheet.cell(row=i,column=11).value=region
    sheet['L3']='user_input'
    sheet['M1']='Octant ID'
    sheet['M2']='overall count'
    sheet['M3']=mod
    val=-1
    for i in range(14,22):
        if i%2==0:
            sheet.cell(row=1,column=i).value=abs(val)
        else:
            sheet.cell(row=1,column=i).value=val
        if i%2!=0:
            val-=1
    count_in_range(mod)
    
from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)
#saving the file in given xlsx file
wb.save('output_octant_transition_identify.xlsx')
