from platform import python_version
import pandas as pd


def saveOctant(octant_op, mod):
    if mod>30000:
        raise Exception('mod value should be less than or equal to 30000')
    size = len(octant_op['octant'])
    actual_size = len(octant_op['octant'])
    index = 0

    # using a while loop to split the data in the given range
    while (size > 0):
        temp = mod
        if index == 0:  # starting from value 0
            x = 0
        else:
            x = index*temp

        y = min(actual_size-1, index*temp+mod-1)
        if size < mod:
            mod = size
            size = 0

        # inserting range and their corresponding data
        lower_bound = str(x)
        upper_bound = str(y)
        octant_op.at[index+2, 'Octant ID'] = lower_bound + '-'+ upper_bound
        octant_in_range = octant_op.loc[x:y]
        octant_op.at[index+2, '-1'] = octant_in_range['octant'].value_counts()[-1]
        octant_op.at[index+2, '1'] = octant_in_range['octant'].value_counts()[1]
        octant_op.at[index+2, '-2'] = octant_in_range['octant'].value_counts()[-2]
        octant_op.at[index+2, '2'] = octant_in_range['octant'].value_counts()[2]
        octant_op.at[index+2, '-3'] = octant_in_range['octant'].value_counts()[-3]
        octant_op.at[index+2, '3'] = octant_in_range['octant'].value_counts()[3]
        octant_op.at[index+2, '-4'] = octant_in_range['octant'].value_counts()[-4]
        octant_op.at[index+2, '4'] = octant_in_range['octant'].value_counts()[4]

        index = index + 1
        size = size - mod

    # writing in octant_output.csv file
    octant_op.to_csv("octant_output.csv")


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


def octant_identification(mod):
    if mod>30000:
        raise Exception('mod value should be less than or equal to 30000')
        
    try:
        # reading octant_input.csv file using pandas module
        octant_ip_md = pd.read_csv("octant_input.csv")
    except FileNotFoundError:
        print('File not found')
        exit()

    # data preprocessing
    # calculating the mean values of U, V, W coordinates upto 9 decimal places
    u_avg = octant_ip_md['U'].mean()
    v_avg = octant_ip_md['V'].mean()
    w_avg = octant_ip_md['W'].mean()
    # inserting new columns
    octant_ip_md.at[0, 'u_avg'] = u_avg
    octant_ip_md.at[0, 'v_avg'] = v_avg
    octant_ip_md.at[0, 'w_avg'] = w_avg

    # calculating pre-processed value of u_avg, v_avg, w_avg by subtracting it from its mean values
    octant_ip_md['U-u_avg'] = octant_ip_md['U']-octant_ip_md.at[0, 'u_avg']
    octant_ip_md['V-v_avg'] = octant_ip_md['V']-octant_ip_md.at[0, 'v_avg']
    octant_ip_md['W-w_avg'] = octant_ip_md['W']-octant_ip_md.at[0, 'w_avg']

    # applying the function made to categorize the data using .apply function
    octant_ip_md['octant'] = octant_ip_md.apply(
        lambda r: checkOct(r['U-u_avg'], r['V-v_avg'], r['W-w_avg']), axis=1)

    # leaving an empty column
    octant_ip_md[' '] = ''

    # counting overall using value_counts function
    octant_ip_md.at[1, ' '] = 'user input'
    octant_ip_md.at[0, 'Octant ID'] = 'overall count'
    octant_ip_md.at[0, '1'] = octant_ip_md['octant'].value_counts()[1]
    octant_ip_md.at[0, '-1'] = octant_ip_md['octant'].value_counts()[-1]
    octant_ip_md.at[0, '2'] = octant_ip_md['octant'].value_counts()[2]
    octant_ip_md.at[0, '-2'] = octant_ip_md['octant'].value_counts()[-2]
    octant_ip_md.at[0, '3'] = octant_ip_md['octant'].value_counts()[3]
    octant_ip_md.at[0, '-3'] = octant_ip_md['octant'].value_counts()[-3]
    octant_ip_md.at[0, '4'] = octant_ip_md['octant'].value_counts()[4]
    octant_ip_md.at[0, '-4'] = octant_ip_md['octant'].value_counts()[-4]

    octant_ip_md.at[1, 'Octant ID'] = mod
    saveOctant(octant_ip_md, mod)

ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod = 5000
try:
    octant_identification(mod)
except NameError:
    print('Either the function name is wrong or the function does not exists')