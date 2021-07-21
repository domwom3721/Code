import pandas as pd
import os
from pathlib import Path
import matplotlib
import matplotlib.pyplot as plt
#matplotlib.use('Agg')
matplotlib.pyplot.ion()
import glob
import time
#from tkinter import *
import shutil
import numpy as np
from docx import Document
from docx.shared import Pt
from docx.shared import Inches



project_location = r'C:\Users\Michael Leahy\Dropbox (Bowery)\Work - Personal Connection\Bowery Project'

#Data Pre Paths
data_location           = os.path.join(project_location,'Data')
map_location            = os.path.join(data_location,'Maps')
costar_data_location       = os.path.join(data_location,'Costar Data') 
costar_summary_location    = os.path.join(data_location,'Costar Summaries') 


#Import data as data frames
df_multifamily  = pd.read_csv(os.path.join(costar_data_location,'mf_raw.csv')) 
df_industrial   = pd.read_csv(os.path.join(costar_data_location,'industrial_raw.csv')) 


#Drop clusters
df_multifamily  = df_multifamily.loc[df_multifamily['Geography Type'] != 'Cluster']
df_industrial   = df_industrial.loc[df_industrial['Geography Type'] != 'Cluster']



#Average Sale Price Variable
for df in [df_multifamily,df_office,df_retail,df_industrial]:
    df['Average Sale Price'] = df['Average Sale Price'].str.replace('$', '', regex=False)
    df['Average Sale Price'] = df['Average Sale Price'].str.replace(',', '', regex=False)
    df['Average Sale Price'] = pd.to_numeric(df['Average Sale Price'])
    
    #Clean Cap rate variable
    df['Market Cap Rate'] = df['Market Cap Rate'].str.replace('%', '',regex=False)
    df['Market Cap Rate'] = df['Market Cap Rate'].str.replace(',', '',regex=False)
    df['Market Cap Rate'] = pd.to_numeric(df['Market Cap Rate'])


df = df_industrial
df = df[df['Geography Name'] == 'New York - NY']


############################Figures start here##################################################:
absorption_plot = df.plot.line(x='Period',y='Average Sale Price',color='Blue',title='Absorption & Vacancy Rates')
df.plot.bar(x='Period',y= 'Market Cap Rate', ax=absorption_plot)

absorption_plot.set_ylabel("")
absorption_plot.set_xlabel("Absorption & Vacancy Rates", fontname="Avenir Next LT Pro", fontsize=10.5)



figure = absorption_plot.get_figure()


#plt.close(figure)
