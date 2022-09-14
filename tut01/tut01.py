import pandas as pd
def octant_identification(mod=5000):

    #reading octant_input.csv file using pandas module
    octant_ip=pd.read_csv("octant_input.csv")
    #calculating the mean values of U, V, W coordinates upto 9 decimal places
    u_avg=octant_ip['U'].mean().__round__(9)
    v_avg=octant_ip['V'].mean().__round__(9)
    w_avg=octant_ip['W'].mean().__round__(9)
    print(u_avg, v_avg, w_avg)

    


from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)