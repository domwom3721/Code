#Author: Mike Leahy
#Summary: Opens the download excel file from Environics and places the sheets inside the hood teplate, closes the Environics file (does not save hood template, keeps it open)
from win32com.client import Dispatch
import os
import shutil


path1 = os.path.join(os.environ['USERPROFILE'], 'Downloads'                                                         ,'Executive Dashboard.xlsx') #the downloaded file
path2 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','HOOD template','HOOD template.xlsm')       #the template we're moving to
path3 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','HOOD template','Executive Dashboard.xlsx') #where we move the downloaded file to



xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

wb1 = xl.Workbooks.Open(Filename=path1)
wb2 = xl.Workbooks.Open(Filename=path2)

for i in range(1,8):    
    ws1 = wb1.Worksheets(i)
    ws1.Copy(Before=wb2.Worksheets("-=-=-=-=-=-=-"))

wb1.Close(SaveChanges=False) #close environics worksheet
# xl.Quit()

#Moves environics worksheet from downloads folder to the hood template folder
shutil.move(path1, path3)
