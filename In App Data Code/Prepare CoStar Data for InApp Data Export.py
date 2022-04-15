#Summary: This program reads 4 csv files from CoStar.com, 1 for each major proprety type. 
#         Each file contains quarterly summary statistics on different markets and submarkets.
#         It appends them together and does some light manipulation as well as creating some new variables and exports a clean file.
#         This clean file will then be pasted directly into a Google Sheet that is connected to the Bowery Appraisal WebApp. 
#         This data can then be inserted into appraisal reports on the WebApp.
#Author: Mike Leahy 04/15/2022

#Import packages we will be using
import os
import pandas as pd

#Define file location pre paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)')  
costar_data_location           =  os.path.join(dropbox_root, 'Research','Projects','App','Data')  

#Import raw CoStar data as pandas dataframes
df_multifamily  = pd.read_excel(os.path.join(costar_data_location,'multifamily_raw.xlsx') ,
                dtype={
                      }      ) 

df_office       = pd.read_excel(os.path.join(costar_data_location,'office_raw.xlsx') ,
                  dtype={
                        }     
                            )

df_retail       = pd.read_excel(os.path.join(costar_data_location,'retail_raw.xlsx') ,
                  dtype={
                       }
                            )

df_industrial   = pd.read_excel(os.path.join(costar_data_location,'industrial_raw.xlsx') ,
                  dtype={
                        }
                             )


print(df_multifamily)
print(df_office)
print(df_retail)
print(df_industrial)