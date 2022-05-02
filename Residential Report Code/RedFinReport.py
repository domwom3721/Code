#Date: 5/2/2022
#Author: Mike Leahy
#Summary: Injests RedFin residential real estate data and produces report documents on the selcted areas

import os
import pandas as pd

#Define file pre-paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project') 
data_location                  =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Clean') 

#Import our clean RedFin data
df = pd.read_csv(os.path.join(data_location, 'Clean RedFin Data.csv'), 
                dtype={         'Type': str,
                                'Region':str,
                                'Month of Period End':str,
                                'Median Sale Price':float	,
                                'Median Sale Price MoM':float ,	
                                'Median Sale Price YoY':float ,	
                                'Homes Sold':float,
                                'Homes Sold MoM':float,
                                'Homes Sold YoY':float,	
                                'New Listings':float,
                                'New Listings MoM':float, 	
                                'New Listings YoY':float,	
                                'Inventory':float,
                                'Inventory MoM':float,	 
                                'Inventory YoY':float,	
                                'Days on Market':float,	
                                'Days on Market MoM':float ,	
                                'Days on Market YoY':float	,
                                'Average Sale To List':float ,
                                'Average Sale To List MoM':float ,	
                                'Average Sale To List YoY':float,
                                },                         
                            )

print(df)