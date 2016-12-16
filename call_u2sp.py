# -*- coding: utf-8 -*-
"""
Created on Fri Nov 13 04:55:03 2015

@author: tong
"""

from __future__ import print_function
import unittest
import os.path
import win32com.client
import time

class ExcelMacro(unittest.TestCase):
    def test_excel_macro(self):
        try:
            xlApp = win32com.client.DispatchEx('Excel.Application')
            xlsPath = os.path.expanduser('C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\Cmd HQ\\TongWare Parts\\uploadtosp.xlsm')
            xlApp.Workbooks.Open(Filename=xlsPath)
            xlApp.Run('upload2spsub')
            xlApp.Quit()
            print(time.strftime("%H:%M:%S")+"The Most updated Resource Model has been UPLOADED to Sharepoint!")
            time.sleep(2)
        except:
            print("Error found while running the excel macro!")
            xlApp.Quit()
if __name__ == "__main__":
    unittest.main()
    
    
    
    
    