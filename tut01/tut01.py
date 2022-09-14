import pandas as pd

def checkOct(u,v,w):
    if u>=0:
      if v>=0:
        if w>=0:
          #this means (u,v) is +ve hence 1st quad and w is +ve so +1
          return 1
        else:
          #this means (u,v) is +ve hence 1st quad but w is -ve so -1
          return -1
      else:
        if w>=0:
          #this means u>0 and v<0 hence 4th quad and w is +ve so +4
          return 4
        else:
          #this means u>0 and v<0 hence 4th quad but w is -ve so -4
          return -4
    if u<0:
      if v>=0:
        if w>=0:
          #this means u<0 and v>0 hence 4th quad and w is +ve so +2
          return 2
        else:
          #this means u<0 and v>0 hence 4th quad but w is -ve so -2
          return -2
      else:
        if w>=0:
          #this means u<0 and v<0 hence 4th quad and w is +ve so +3
          return 3
        else:
          #this means u<0 and v<0 hence 4th quad and w is -ve so -3
          return -3

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

    #applying the function made to categorize the data using .apply function
    octant_ip_md['octant'] = octant_ip_md.apply(lambda r: checkOct(r['U-u_avg'], r['V-v_avg'], r['W-w_avg']),axis=1)


from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)