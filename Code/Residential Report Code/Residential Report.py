#Residential Report
#By Research Q3 2021:
#Summary:
    #Imports 5 clean datafiles with summary statistics from Realtor.com for residential real estate (all types covered)
    #Loops through these 5 files, loops through each of the geographic areas and creates a directory and word document
    #The word document is a report that reports tables and graphs generated from the data files

import os
import time
import numpy as np
import pandas as pd
from tkinter import *

from docx import Document
from docx.dml.color import ColorFormat
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from Graph_Functions import *  
from Language_Functions import *  
from Table_Functions import * 

#Define file pre paths
start_time = time.time() #Used to track runtime of script

dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project') #Main Folder that stores all output, code, and documentation

# output_location                = os.path.join(project_location,'Output','Market Reports')             #The folder where we store our current reports, testing folder
output_location                = os.path.join(dropbox_root,'Research','Market Analysis','Market')       #The folder where we store our current reports, production

#map_location                   = os.path.join(project_location,'Data','Maps','CoStar Maps')            #Folders with maps png files  
realtor_data_location           = os.path.join(project_location,'Data','Realtor Data')                  #Folder with clean realtor CSV files
realtor_writeup_location        = os.path.join(project_location,'Data','Realtor Writeups')              #Folder with clean realtor CSV files

#Import cleaned data from Clean Realtor Data.py
df_national            = pd.read_csv(os.path.join(realtor_data_location,'national_clean.csv'))
df_state               = pd.read_csv(os.path.join(realtor_data_location,'state_clean.csv'))
df_metro               = pd.read_csv(os.path.join(realtor_data_location,'metro_clean.csv')) 
df_county              = pd.read_csv(os.path.join(realtor_data_location,'county_clean.csv')) 
df_zip                 = pd.read_csv(os.path.join(realtor_data_location,'zip_clean.csv')) 


#Define functions used to handle the clean Realtor data and help write our reports
def CreateMarketDictionary(df): #Creates a dictionary where each key is a State and the items are lists of its counties
     df_states         = df.loc[df['state'] == 'State'] 
     df_metros         = df.loc[df['cbsa_title'] == 'Metro']
     df_counties       = df.loc[df['county_name'] == 'County']
     df_zips           = df.loc[df['postal_code'] == 'Zip']
     unique_states_list    = df_states['state'].unique()
     unique_metros_list    = df_metros['cbsa_title'].unique()
     unique_counties_list  = df_counties['county_name'].unique()
     unique_zipss_list     = df_zips['postal_code'].unique()

     #Now create dictionary to track which counties belong to each metro
     market_dictionary = {}
     for metro in unique_metros_list:
         counties = [county for county in unique_metros_list if metro in county ] #list of counties within current metro
         zips = [zip for zip in unique_counties_list if county in zip ] #list of zips within current county
         market_dictionary.update({metro:county}) 
        #  market_dictionary.update({county:zip}) 

     return(market_dictionary)

def CleanMarketName(market_name):
    clean_market_name = market_name.replace("""/""",' ')
    clean_market_name = clean_market_name.replace(""":""",'')
    clean_market_name = clean_market_name.replace("""'""",'')
    clean_market_name = clean_market_name.strip()
    # if clean_market_name == 'Manhattan - NY':
    #     clean_market_name = 'Manhattan'

    return(clean_market_name)

def CreateEmptySalesforceLists():
    global  dropbox_primary_markets,dropbox_markets,dropbox_sectors, dropbox_sectors_codes
    global dropbox_links,dropbox_research_names,dropbox_analysis_types,dropbox_states,dropbox_versions,dropbox_statuses,dropbox_document_names
    dropbox_primary_markets        = []
    dropbox_markets                = []
    dropbox_sectors                = []
    dropbox_sectors_codes          = []
    dropbox_links                  = []
    dropbox_research_names         = []
    dropbox_analysis_types         = []
    dropbox_states                 = []
    dropbox_versions               = []
    dropbox_statuses               = []
    dropbox_document_names         = []