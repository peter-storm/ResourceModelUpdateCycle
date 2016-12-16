
# Created on Tue Oct 13 23:46:55 2015
# Author: Tong Wu
 
# This Module serves to launch Intergration Managers of OpenAir to export operational data
# All Integration Manager shortcuts has been stored in:
# C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\IM ShortCut\\New RM IM Shortcuts\\
# with an extension of 'lnk'

import subprocess,glob,time

path="C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\New RM IM Shortcuts\\"
folder=glob.glob(path+"*.lnk")


for shortcut in folder:
    
    subprocess.call(shortcut,shell=True)
    print(time.strftime("%H:%M:%S")+ " : "+shortcut[68:]+" has been done its work.")
    
