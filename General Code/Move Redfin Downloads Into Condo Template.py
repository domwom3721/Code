#Author: Mike Leahy
#Summary:
from win32com.client import Dispatch
import os
import shutil
import pandas as pd

#Define the location of the files we downloaded from Redfin.com
path1 = os.path.join(os.environ['USERPROFILE'], 'Downloads'                                                         ,'data.xlsx') #the main downloaded file
path2 = os.path.join(os.environ['USERPROFILE'], 'Downloads'                                                         ,'ppsf.xlsx') #the price per square foot downloaded file

path3 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','Redfin Condo Template','data.xlsx') #where we move the downloaded file to
path4 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','Redfin Condo Template','ppsf.xlsx') #where we move the downloaded file to

path5 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','Redfin Condo Template','REDFIN Condo Template.xlsx') #the template
path6 = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Templates','Redfin Condo Template','Combined Redfin Data.xlsx') #new file with the redfin data combined

#Move Redfin data files into template folder from our downloads folder
if os.path.exists(path1):
    shutil.move(path1, path3)

if os.path.exists(path2):
    shutil.move(path2, path4)

if os.path.exists(path3) and os.path.exists(path4) :
    #Import the downloaded redfin data as pandas dataframes
    main_redfin_data_df = pd.read_excel(path3)
    ppsf_redfin_data_df = pd.read_excel(path4,header = 1)

    #Fill in the region name because it's only in the first row of the main data file
    main_redfin_data_df['Region'] =  main_redfin_data_df['Region'].fillna(method='ffill')

    #Convert the price per sqft file from wide to long
    ppsf_redfin_data_df   = ppsf_redfin_data_df.transpose().reset_index(col_level=1)
    master_temp_df =  pd.DataFrame(columns=['Month of Period End','Region'])
    for i in range(0,len(ppsf_redfin_data_df.columns) - 1):
        temp_df = ppsf_redfin_data_df[['index',i]]
        temp_df['Region'] = temp_df[i].iloc[0]
        temp_df = temp_df.iloc[1:]
        temp_df   = temp_df.rename(columns={'index': "Month of Period End",i:'Price/SF'})
        master_temp_df = master_temp_df.append(temp_df)
        
    #Merge the main data df and the price per sqft df
    combined_redfin_df    = pd.merge(main_redfin_data_df,master_temp_df,on=('Month of Period End','Region'),how='left')
    combined_redfin_df.to_excel(path6,index=False) #export the combined redfin to it's own excel file

    #Delete the  redfin files we downloaded
    os.remove(path3)
    os.remove(path4)


xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

wb1 = xl.Workbooks.Open(Filename=path6)
wb2 = xl.Workbooks.Open(Filename=path5)

  
ws1 = wb1.Worksheets(1)
ws1.Copy(Before=wb2.Worksheets("datadump"))     #Move redfin data into template file
wb1.Close(SaveChanges=False)                    #close redfin data file

wb2.Worksheets('Sheet1').Name          = "Redfin Data"   #change the name of the redfin data from Sheet1 to Redfin Data 
redfin_sheet                           = wb2.Sheets("Redfin Data")
redfin_sheet.Cells(3,'Y').Formula      = '=UNIQUE(A:A)'





