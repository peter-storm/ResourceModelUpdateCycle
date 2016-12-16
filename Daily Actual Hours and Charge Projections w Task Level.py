
 # Before start, make sure your machine has installed firefox as the running browser;
 # and the following python packages

import os,os.path,time
from selenium.webdriver.common.by import By
from selenium import webdriver

 # Initial settings for browser
 # Download path
 # Save to disk
 # Minimize browser window
 



fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList", 2)
fp.set_preference("browser.download.manager.showWhenStarting", False)
fp.set_preference("browser.download.dir","C:\SkunkWorks\SPSync_Skunkwork\SkunkWork_Report\DATA")
fp.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

report1="C:\SkunkWorks\SPSync_Skunkwork\SkunkWork_Report\DATA\Daily_Actual_Hours_and_Charge_Projections_w_Task_Level_report_pivot.csv"


browser=webdriver.Firefox(fp)
browser.set_window_size(0, 0)


 # Navigate to main webpages

browser.get('http://openair.com/index.pl?_nocookies=1')



 # Login in the page
browser.find_element_by_id('input_company').send_keys('tcube')
browser.find_element_by_id('input_user').send_keys('interface')
browser.find_element_by_id('input_password').send_keys('Tqsolutions123')
browser.find_element_by_id('oa_comp_login_submit').click()

r0=browser.find_element_by_link_text('Daily Actual Hours and Charge Projections w/ Task Level')
r0.click()


 # Find click download buttons
button1=browser.find_element_by_link_text('Click here')
if os.path.exists(report1) == True :
    os.remove(report1)

button1.click()
time.sleep(2)

while os.path.exists(report1) == False :
    button1.click()

print (time.strftime("%H:%M:%S")+ " : Report <Daily Actual Hours and Charge Projections w/ Task Level> is downloaded.")
time.sleep(3)
browser.close()




