import pandas as pd
def octant_identification(mod=5000):

    #reading octant_input.csv file using pandas module
    octant_ip_md=pd.read_csv("octant_input.csv")

    
    #data preprocessing
    #calculating the mean values of U, V, W coordinates upto 9 decimal places
    u_avg=octant_ip_md['U'].mean().__round__(9)
    v_avg=octant_ip_md['V'].mean().__round__(9)
    w_avg=octant_ip_md['W'].mean().__round__(9)
    #inserting new columns
    octant_ip_md.at[0,'u_avg']=u_avg
    octant_ip_md.at[0,'v_avg']=v_avg
    octant_ip_md.at[0,'w_avg']=w_avg
    #calculating pre-processed value of u_avg, v_avg, w_avg by subtracting it from its mean values
    octant_ip_md['U-u_avg']=octant_ip_md['U']-octant_ip_md.at[0,'u_avg']
    octant_ip_md['V-v_avg']=octant_ip_md['V']-octant_ip_md.at[0,'v_avg']
    octant_ip_md['W-w_avg']=octant_ip_md['W']-octant_ip_md.at[0,'w_avg']
    


from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)