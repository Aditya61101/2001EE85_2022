import pandas as pd
def octant_identification(mod=5000):
    #reading octant_input.csv file using pandas module
    octant_ip=pd.read_csv("octant_input.csv")
    print(octant_ip.head())

    


from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

mod=5000
octant_identification(mod)