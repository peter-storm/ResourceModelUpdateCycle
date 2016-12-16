# -*- coding: utf-8 -*-
"""
Created on Wed Nov 11 00:21:48 2015

@author: tong
"""
import time
from win32com.client import Dispatch  
print (time.strftime("%H:%M:%S")+ " : Hey I am refreshing ETL for clean data. Wait a little bit..")
time.sleep(1)
xl = Dispatch('Excel.Application')  
wbr = xl.Workbooks.Open(r'C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\New_Resource_Model\\Calculations\\NRM_ETL.xlsx')
wbr.RefreshAll() 
time.sleep(30)
wbr.Close(True) 
