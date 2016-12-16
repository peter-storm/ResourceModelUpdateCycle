import time
from Tongware import lastmodified as lm




print ("LoBmm Tongware: Working for lazy people")
time.sleep(1)
print ("----PAT Update Machine----")


#Initial Paths Settings
Tongware_accessory_Path="C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\Cmd HQ\\TongWare Parts\\"
DAHCP=Tongware_accessory_Path+"Daily Actual Hours and Charge Projections w Task Level.py"
IM_Shorcuts=Tongware_accessory_Path+"call_im_shortcut.py"
NRM_Data_update=Tongware_accessory_Path+"Refresh_NRM.py"
ETL_update=Tongware_accessory_Path+"Refresh_ETL.py"
ETL_Tool="C:\\SkunkWorks\\SPSync_Skunkwork\\SkunkWork_Report\\New_Resource_Model\\Calculations\\NRM_ETL.xlsx"
callu2sp=Tongware_accessory_Path+"call_u2sp.py"
LRW=Tongware_accessory_Path+"LastRefresh_Writer.py"


#Tongware OFFICIALLY start working
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 1 - Launch Integration Manager to export TCube Operational Data.")
time.sleep(1)
print (time.strftime("%H:%M:%S")+ " : It might take 20 minutes.")
time.sleep(1)
print (time.strftime("%H:%M:%S")+ " : Please wait...")
time.sleep(1)

with open(IM_Shorcuts) as f:
    code = compile(f.read(), IM_Shorcuts, 'exec')
    exec(code)



#Scraping OpenAir calculated tables
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 2 - Scraping OpenAir calculated tables.")
time.sleep(1)
print (time.strftime("%H:%M:%S")+ " : It might take 1 minute.")
time.sleep(1)
print (time.strftime("%H:%M:%S")+ " : Please wait...")
time.sleep(1)

with open(DAHCP) as f:
    code = compile(f.read(), DAHCP, 'exec')
    exec(code)


#Tongware refresh ETL Tools.
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 3 - Refresh ETL Tool & Resource Model Data.")
time.sleep(1)
print (time.strftime("%H:%M:%S")+ " : It might take less than 1 minute.")
time.sleep(1)
with open(ETL_update) as f:
    code = compile(f.read(), ETL_update, 'exec')
    exec(code)

time.sleep(3)

with open(NRM_Data_update) as f:
    code = compile(f.read(), NRM_Data_update, 'exec')
    exec(code)


#Tongware last check before uploads.
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 4 - Check if data is refreshed.")
print("Daily Actual Report last modified at: "+lm(DAHCP))
print("NRM ETL Tool last modified at: "+lm(ETL_Tool))
time.sleep(1)



#Tongware writes last refresh time.
#----------------------------------------------------------------------------------------------
with open(LRW) as f:
    code = compile(f.read(), LRW, 'exec')
    exec(code)





#Tongware uploads NRM to SharePoint
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 5 - TongWare is ready for uploading the new Resource Model at SharePoint Site...")
time.sleep(1)

print (time.strftime("%H:%M:%S")+ " : It might take 15 seconds.")
with open(callu2sp) as f:
    code = compile(f.read(), callu2sp, 'exec')
    exec(code)





#Tongware has done its work.
#----------------------------------------------------------------------------------------------
print (time.strftime("%H:%M:%S")+ " : PART 6 - Tongware has succuessfully done all work.")
time.sleep(40)
