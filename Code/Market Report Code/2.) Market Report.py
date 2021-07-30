#By Mike Leahy April 29, 2021:
#Summary:
    #Imports 4 clean datafiles with summary statistics from CoStar.com on commerical real estate for the 4 main sectors
    #Loops through these 4 files, loops through each of the markets and submarkets (geographic areas) and creates a directory and word document
    #The word document is a report that reports tables and graphs generated from the data files

import os
import time
import numpy as np
import pandas as pd
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

output_location                = os.path.join(project_location,'Output','Market Reports')         #The folder where we store our current reports, testing folder
output_location                = os.path.join(dropbox_root,'Research','Market Analysis','Market')          #The folder where we store our current reports, production

map_location                   = os.path.join(project_location,'Data','Maps','CoStar Maps')       #Folder with clean CoStar CSV files
costar_data_location           = os.path.join(project_location,'Data','Costar Data')              #Folders with maps png files

#Import cleaned data from 1.) Clean Costar Data.py
df_multifamily  = pd.read_csv(os.path.join(costar_data_location,'mf_clean.csv')) 
df_office       = pd.read_csv(os.path.join(costar_data_location,'office_clean.csv'))
df_retail       = pd.read_csv(os.path.join(costar_data_location,'retail_clean.csv'))
df_industrial   = pd.read_csv(os.path.join(costar_data_location,'industrial_clean.csv')) 

df_multifamily_slices  = pd.read_csv(os.path.join(costar_data_location,'mf_slices_clean.csv')) 
df_office_slices       = pd.read_csv(os.path.join(costar_data_location,'office_slices_clean.csv'))
df_retail_slices       = pd.read_csv(os.path.join(costar_data_location,'retail_slices_clean.csv'))
df_industrial_slices   = pd.read_csv(os.path.join(costar_data_location,'industrial_slices_clean.csv')) 




#Define functions used to handle the clean CoStar data and help write our repots
def CreateMarketDictionary(df): #Creates a dictionary where each key is a market and the items are lists of its submarkets
     df_markets             = df.loc[df['Geography Type'] == 'Metro'] 
     df_submarkets          = df.loc[df['Geography Type'] == 'Submarket']
     df_clusters            = df.loc[df['Geography Type'] == 'Cluster']
     unique_markets_list    = df_markets['Geography Name'].unique()
     unique_submarkets_list = df_submarkets['Geography Name'].unique()
     unique_cluster_list    = df_clusters['Geography Name'].unique()

     #Now create dictionary to track which submarkets belong to each market
     market_dictionary = {}
     for market in unique_markets_list:
         submarkets = [submarket for submarket in unique_submarkets_list if market in submarket ] #list of sumarkets within current market
         clusters = [cluster for cluster in unique_cluster_list if market in cluster ] #list of sumarkets within current market
         market_dictionary.update({market:submarkets}) 
        #  market_dictionary.update({market:clusters}) 

     return(market_dictionary)

def CleanMarketName(market_name):
    clean_market_name = market_name.replace("""/""",' ')
    clean_market_name = clean_market_name.replace(""":""",'')
    clean_market_name = clean_market_name.replace("""'""",'')
    clean_market_name = clean_market_name.strip()
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

def UpdateSalesforceMarketList(markets_list, submarkets_list, sector_list, sector_code_list, dropbox_links_list):
    #Add to lists that track our markets and submarkets for salesforce
    markets_list.append(state + '-' + primary_market_name_for_file)
    if market == primary_market:
        submarkets_list.append('')
        dropbox_analysis_types.append('Market')
    else:
        submarkets_list.append(state + '-' + market_title)
        dropbox_analysis_types.append('Submarket')

    sector_list.append(sector)
    if sector == 'Multifamily':
        sector_code_list.append('MF')
    else:
        sector_code_list.append(sector[0])

    #Use the output directory to back into the right dropbox link 
    dropbox_link = output_directory.replace(dropbox_root,r'https://www.dropbox.com/home')
    dropbox_link = dropbox_link.replace("\\",r'/')
    dropbox_links_list.append(dropbox_link)

    dropbox_states.append(state)
    latest_quarter = df_market_cut.iloc[-1]['Period']
    dropbox_versions.append(latest_quarter)

    if market == primary_market:
        dropbox_research_names.append(state + ' - ' + primary_market_name_for_file + ' - ' + sector )
    else:
        if market_title == primary_market_name_for_file : #if market and submarket have same name
            dropbox_research_names.append(state + ' - ' + market_title + ' SUB' + ' - ' + sector )
        else:
            dropbox_research_names.append(state + ' - ' + market_title + ' - ' + sector )


    dropbox_document_names.append(report_file_title)
    dropbox_statuses.append('Outdated')
    
def CreateOutputDirectory():
    sector_folder        = os.path.join(output_location,sector)
    state_folder         = os.path.join(output_location,sector,state)
    market_folder        = os.path.join(state_folder,primary_market_name_for_file)

    if market == primary_market:
        output_directory     = market_folder                    #Folder where we write report to
    else:
        output_directory     = os.path.join(market_folder,str(market_title)) 


    #Check if output,map, and summary folder already exists, and if it doesnt, make it
    for folder in [sector_folder,state_folder,market_folder,output_directory]:
        if os.path.exists(folder):
            pass #does nothing
        else:
            os.mkdir(folder) #Create new folder for market or submarket
    return(output_directory)

def CreateMapDirectory():
    state_folder_map         = os.path.join(map_location,sector,state)
    market_folder_map        = os.path.join(state_folder_map,primary_market_name_for_file)

    if market == primary_market:
        map_directory        = market_folder_map                #Folder where we store map for market or submarket
    else:
        map_directory        = os.path.join(market_folder_map,str(market_title))
    
    #Check if output,map, and summary folder already exists, and if it doesnt, make it
    for folder in [state_folder_map,market_folder_map,map_directory]:
        if os.path.exists(folder):
            pass 
        else:
            os.mkdir(folder) #Create new folder for market or submarket
    
    return(map_directory)

def CreateReportFilePath():
    global report_file_title
    #Create report
    if market == primary_market:
        market_file_name = primary_market_name_for_file
        macro_or_sub = 'Market'
    else:
        market_file_name = market_title
        macro_or_sub = 'Submarket'
    
    if sector == "Multifamily":
        sector_code = "MF"
    else:
        sector_code = sector[0]

    report_file_title =   latest_quarter  + ' ' +  state + ' - '   + market_file_name + ' - ' + sector + ' ' + macro_or_sub  + '_draft' + '.docx'

    #Make sure we don't hit the 255 max file path limit
    if len(os.path.join(output_directory,report_file_title)) <= 257:
        report_path = os.path.join(output_directory,report_file_title)
    else:
        report_path = os.path.join(output_directory,(latest_quarter + '-' + market_file_name + '_draft' + '.docx')  )

    
    assert os.path.exists(output_directory)
    return(report_path)

def SetPageMargins():
    #Page Margins
    sections = document.sections
    for section in sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

def SetStyle():
    #Style
    style = document.styles['Normal']
    font = style.font
    font.name = 'Avenir Next LT Pro Light'
    font.size = Pt(9)

def MakeReportTitle():
    #Write title Heading
    if market == primary_market:
        title = document.add_heading(market_title + ': ' + sector + ' Market Analysis',level=1)
    else:
        title = document.add_heading(market_title + ': ' + sector + ' Submarket Analysis',level=1)
    title.style = document.styles['Heading 2']
    title.paragraph_format.space_after  = Pt(6)
    title.paragraph_format.space_before = Pt(12)
    title_style = title.style
    title_style.font.name = "Avenir Next LT Pro Light"
    title_style.font.size = Pt(14)
    title_style.font.bold = False
    title_style.font.color.rgb = RGBColor.from_string('3F65AB')
    title_style.element.xml
    rFonts = title_style.element.rPr.rFonts
    rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro Light")

def MakeCoStarDisclaimer():
    #Write Costar disclaimer as of '  + latest_quarter + ' 
    if market == primary_market:
        disclaimer = document.add_paragraph('The information contained in this report was provided using ' 
                                            + latest_quarter + 
                                            ' CoStar data for the ' + 
                                            market + 
                                            ' ' + 
                                            sector + 
                                            """ Market ("Market").""")
    #Submarket disclaimer
    else:
        disclaimer = document.add_paragraph('The information contained in this report was provided using ' +
                                            latest_quarter + 
                                            ' CoStar data for the ' + 
                                            market_title + 
                                            ' ' + 
                                            sector + 
                                            """ Submarket ("Submarket") """ +
                                            'located in the ' +
                                            primary_market +
                                            """ Market ("Market"). """ +
                                            'The Submarket includes the neighborhoods/towns of .'
                                            )                
    
    
    disclaimer.style.font.name = "Avenir Next LT Pro (Body)"
    disclaimer.style.font.size = Pt(9)
    disclaimer.paragraph_format.space_after  = Pt(6)
    disclaimer.paragraph_format.space_before = Pt(0)
    disclaimer.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    disclaimer.paragraph_format.keep_together = True

def CleanUpPNGs():
    #Report writing done, delete figures
    try:
        files = os.listdir(output_directory)
        for image in files:
            if image.endswith(".png"):
                os.remove(os.path.join(output_directory, image))
                
    except Exception as e: print(e)
   
def AddMap():
    
    #Add image of map if there is one in the appropriate map folder
    if os.path.exists(os.path.join(map_directory,'map.png')):
        map = document.add_picture(os.path.join(map_directory,'map.png'), width=Inches(6.5) )
    else:
        map = document.add_paragraph('')

    last_paragraph = document.paragraphs[-1] 
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def SupplyDemandSection():
    #Supply and Demand Section
    AddHeading(document,'Supply & Demand',2)
    supply_demand_paragraph = document.add_paragraph(demand_language)
    supply_demand_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    supply_demand_paragraph.paragraph_format.space_after  = Pt(6)
    supply_demand_paragraph.paragraph_format.space_before = Pt(0)
    supply_demand_paragraph_style = supply_demand_paragraph.style
    supply_demand_paragraph_style.font.name = "Avenir Next LT Pro (Body)"

    #Vacancy Table
    vacancy_table_title_paragraph = document.add_paragraph('Vacancy Rates')
    vacancy_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    vacancy_table_title_paragraph.paragraph_format.space_after  = Pt(6)
    vacancy_table_title_paragraph.paragraph_format.space_before = Pt(12)
    for run in vacancy_table_title_paragraph.runs:
        font = run.font
        font.name = 'Avenir Next LT Pro Medium'
    
 

    vacancy_table_width = 0.64
    AddTable(document,data_for_vacancy_table,vacancy_table_width)
    
            
    blank_paragraph_after_vac_table = document.add_paragraph('')
    blank_paragraph_after_vac_table.paragraph_format.space_after = Pt(6)
    blank_paragraph_after_vac_table.paragraph_format.space_after = Pt(0)


    #Absorption rate Graph
    if os.path.exists(os.path.join(output_directory,'absorption_rate.png')):
        absorption_figrue = document.add_picture(os.path.join(output_directory,'absorption_rate.png'),width=Inches(6.5))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        absorption_format = document.styles['Normal'].paragraph_format
        absorption_format.space_after = Pt(0)
    

def OverviewSection():

    #Overview Heading
    AddHeading(document,'Overview',2)
    
    #Overview Paragraph
    summary_paragraph = document.add_paragraph(overview_language)
    summary_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    summary_paragraph.paragraph_format.space_after = Pt(0)
    summary_paragraph.paragraph_format.space_after = Pt(6)
    summary_paragraph_style = summary_paragraph.style
    summary_paragraph_style.font.name = "Avenir Next LT Pro (Body)"
    summary_paragraph_style.font.size = Pt(9)
    summary_format = document.styles['Normal'].paragraph_format
    summary_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


   

    #Overview table title
    overview_table_title_paragraph = document.add_paragraph('Sector Fundamentals')
    overview_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    overview_table_title_paragraph.paragraph_format.space_after  = Pt(6)
    overview_table_title_paragraph.paragraph_format.space_before = Pt(12)
    overview_table_title_paragraph.keep_with_next = True
    overview_table_title_paragraph.keep_together  = True

    for run in overview_table_title_paragraph.runs:
                    font = run.font
                    font.name = 'Avenir Next LT Pro Medium'
    
    #Overview table
    if sector == 'Multifamily':
        AddOverviewTable(document,8,7,data_for_overview_table,1.2)
    else:
        AddOverviewTable(document,9,7,data_for_overview_table,1.2)


    #Preamble to historical performance table
    if df_market_cut.equals(df_primary_market):
        market_or_submarket = 'Market'
    else:
        market_or_submarket = 'Submarket'
    
    document.add_paragraph('')
    preamble_language = ('Supply and demand indicators, including inventory levels, absorption, vacancy, and rental rates for ' +
                         sector.lower() +
                         ' space in the ' +
                         market_or_submarket +
                         ' are presented in the ensuing table.')

    
    
    table_preamble = document.add_paragraph(preamble_language)
    table_preamble.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    #Add a market performance table
    if market == primary_market:
        performance_table_title_paragraph = document.add_paragraph('Historical ' + sector  + ' Performance: ' +  market_title + ' Market' )
        performance_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        performance_table_title_paragraph.paragraph_format.space_after  = Pt(6)
        performance_table_title_paragraph.paragraph_format.space_before = Pt(12)

        for run in performance_table_title_paragraph.runs:
                    font = run.font
                    font.name = 'Avenir Next LT Pro Medium'

        AddMarketPerformanceTable(document = document,market_data_frame = df_primary_market,col_width = 1.2,sector=sector)
    else:
        performance_table_title_paragraph = document.add_paragraph('Historical ' + sector  + ' Performance: ' +  market_title + ' Submarket')
        performance_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        performance_table_title_paragraph.paragraph_format.space_after  = Pt(6)
        performance_table_title_paragraph.paragraph_format.space_before = Pt(12)

        for run in performance_table_title_paragraph.runs:
                    font = run.font
                    font.name = 'Avenir Next LT Pro Medium'

        AddMarketPerformanceTable(document = document,market_data_frame = df_market_cut ,col_width = 1.2,sector=sector)


    
def RentSecton():
    
    #Rent Section
    AddHeading(document,'Rents',3)       
    rent_paragraph = document.add_paragraph(rent_language)
    rent_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rent_paragraph.paragraph_format.space_after = Pt(6)
    rent_paragraph.paragraph_format.space_before = Pt(0)
    

    #Rent Table
    rent_table_title_paragraph = document.add_paragraph('Market Rents')
    rent_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rent_table_title_paragraph.paragraph_format.space_after  = Pt(6)
    rent_table_title_paragraph.paragraph_format.space_before = Pt(12)
    for run in rent_table_title_paragraph.runs:
                    font = run.font
                    font.name = 'Avenir Next LT Pro Medium'
 

    rent_table_width = 1.2
    AddTable(document,data_for_rent_table,rent_table_width)
  

    blank_paragraph_after_rent_table = document.add_paragraph('')
    blank_paragraph_after_rent_table.paragraph_format.space_after = Pt(6)
    blank_paragraph_after_rent_table.paragraph_format.space_after = Pt(0)

    #Insert rent growth graph
    if os.path.exists(os.path.join(output_directory,'rent_growth.png')):
        rent_growth_figure = document.add_picture(os.path.join(output_directory,'rent_growth.png'),width=Inches(6.5))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pass

def ConstructionSection():
    #Construction Section
    AddHeading(document,'Construction & Future Supply',2)
    constr_paragraph = document.add_paragraph(construction_languge)
    constr_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    constr_paragraph.paragraph_format.space_after = Pt(6)
    constr_paragraph.paragraph_format.space_before = Pt(0)

    #Insert construction graph
    if os.path.exists(os.path.join(output_directory,'construction_volume.png')):
        construction_graph = document.add_picture(os.path.join(output_directory,'construction_volume.png'),width=Inches(6.5))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  
    else:
        pass

def CapitalMarketsSection():
    #Captial Markets Section
    AddHeading(document,'Capital Markets',2)
    capital_paragraph = document.add_paragraph(sale_language)
    capital_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    capital_paragraph.paragraph_format.space_after  = Pt(6)
    capital_paragraph.paragraph_format.space_before = Pt(0)

    #Sales Volume Graphs
    if os.path.exists(os.path.join(output_directory,'sales_volume.png')):
        sales_volume_graph = document.add_picture(os.path.join(output_directory,'sales_volume.png'),width=Inches(6.5))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph('')
    else:
        pass 

    #Key Sales Table
    sales_table_title_paragraph = document.add_paragraph('Key Sales Transactions ' + latest_quarter)
    sales_table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sales_table_title_paragraph.paragraph_format.space_after  = Pt(6)
    sales_table_title_paragraph.paragraph_format.space_before = Pt(12)
    for run in sales_table_title_paragraph.runs:
        font = run.font
        font.name = 'Avenir Next LT Pro Medium'

    #Create data for sales table (Market)
    if df_market_cut.equals(df_primary_market):
        if sector == 'Multifamily':
            data_for_sales_table = [['Property',	'Submarket',	'Tenant',	'Units',	'Type'],['X' for i in range(5)],['X' for i in range(5)],['X' for i in range(5)],['X' for i in range(5)]]
        else:
            data_for_sales_table = [['Property',	'Submarket',	'Tenant',	'SF', 'Type'],['X' for i in range(5)],['X' for i in range(5)],['X' for i in range(5)],['X' for i in range(5)]]
    
    #Create data for sales table (Submarket)
    else:
        if sector == 'Multifamily':
            data_for_sales_table = [['Property',		'Tenant',	'Units',	'Type'],['X' for i in range(4)],['X' for i in range(4)],['X' for i in range(4)],['X' for i in range(4)]]
        else:
            data_for_sales_table = [['Property',		'Tenant',	'SF', 'Type'],['X' for i in range(4)],['X' for i in range(4)],['X' for i in range(4)],['X' for i in range(4)]]

    AddTable(document,data_for_sales_table,col_width=1)
    document.add_paragraph('')
    
    #Asset Value  Graph
    if os.path.exists(os.path.join(output_directory,'asset_values.png')):
        asset_value_graph = document.add_picture(os.path.join(output_directory,'asset_values.png'),width=Inches(6.5))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pass

def OutlookSection():
    #Outlook Section
    AddHeading(document,'Outlook',2)
    conclusion_paragraph = document.add_paragraph(outlook_language)
    conclusion_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    conclusion_paragraph.paragraph_format.space_after  = Pt(6)
    conclusion_paragraph.paragraph_format.space_before = Pt(0)

def GetLanguage():
    global overview_language, demand_language,sale_language,rent_language,construction_languge,outlook_language
    overview_language    = CreateOverviewLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)
    demand_language      = CreateDemandLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)
    sale_language        = CreateSaleLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)
    rent_language        = CreateRentLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)
    construction_languge = CreateConstructionLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)
    outlook_language     = CreateOutlookLanguage(df_market_cut,df_primary_market,df_nation,market_title,primary_market,sector)

def GetOverviewTable():
    #Create Data for overview table
    
    #There are 4 possible permuations for this table (market/apt, market/nonapt, submarket/apt, submakert/nonapt)
    if sector == 'Multifamily':
        data_for_overview_table = [ [],[],[],[],[],[],[],[] ]
    else:
        data_for_overview_table = [ [],[],[],[],[],[],[],[],[] ]


    #Write Top Row of Report
    if market == primary_market: #market report
        data_for_overview_table[0] = ['',market,'YoY','QoQ','National','YoY','QoQ']
    else:
        data_for_overview_table[0] = ['',market_title,'YoY','QoQ',primary_market,'YoY','QoQ']

    #Rows for non-apt
    if sector != 'Multifamily':
        #Rent Growth Row
        data_for_overview_table[1] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Market Rent/SF',
                                                                'YoY Rent Growth',
                                                                'QoQ Rent Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Market Rent/SF')
        #Vacancy Row
        data_for_overview_table[2] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Vacancy Rate',
                                                                'YoY Vacancy Growth',
                                                                'QoQ Vacancy Growth',
                                                                '%',
                                                                'bps',
                                                                'bps',
                                                                'Vacancy Rate')
        #Availability Rate Row
        data_for_overview_table[3] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Availability Rate',
                                                                'YoY Availability Rate Growth',
                                                                'QoQ Availability Rate Growth',
                                                                '%',
                                                                'bps',
                                                                'bps',
                                                                'Availability Rate')
        
        #Absorption Row
        data_for_overview_table[4] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Net Absorption SF',
                                                                'YoY Net Absorption SF Growth',
                                                                'QoQ Net Absorption SF Growth',
                                                                '',
                                                                '%',
                                                                '%',
                                                                'Net Absorption SF')
        
        #Asset Value Row
        data_for_overview_table[5] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Asset Value/Sqft',
                                                                'YoY Asset Value/Sqft Growth',
                                                                'QoQ Asset Value/Sqft Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Asset Value/SF')
        
        #Market Cap Rate Row
        data_for_overview_table[6] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Market Cap Rate',
                                                                'YoY Market Cap Rate Growth',
                                                                'QoQ Market Cap Rate Growth',
                                                                '%',
                                                                'bps',
                                                                'bps',
                                                                'Market Cap Rate')
        
        #Transaction Count Row
        data_for_overview_table[7] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Sales Volume Transactions',
                                                                'YoY Transactions Growth',
                                                                'QoQ Transactions Growth',
                                                                '',
                                                                '%',
                                                                '%',
                                                                'Transaction Count')
        #Sales Volume Row
        data_for_overview_table[8] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Total Sales Volume',
                                                                'YoY Total Sales Volume Growth',
                                                                'QoQ Total Sales Volume Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Sales Volume')

    #Rows for apt
    if sector == 'Multifamily':
        #Rent row
        data_for_overview_table[1] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Market Effective Rent/Unit',
                                                                'YoY Market Effective Rent/Unit Growth',
                                                                'QoQ Market Effective Rent/Unit Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Market Rent/Unit')

        #Vacancy row
        data_for_overview_table[2] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Vacancy Rate',
                                                                'YoY Vacancy Growth',
                                                                'QoQ Vacancy Growth',
                                                                '%',
                                                                'bps',
                                                                'bps',
                                                                'Vacancy Rate')

        #Absorption row
        data_for_overview_table[3] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Absorption Units',
                                                                'YoY Absorption Units Growth',
                                                                'QoQ Absorption Units Growth',
                                                                '',
                                                                '%',
                                                                '%',
                                                                'Net Absorption Units')

        #Asset value row
        data_for_overview_table[4] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Asset Value/Unit',
                                                                'YoY Asset Value/Unit Growth',
                                                                'QoQ Asset Value/Unit Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Asset Value/Unit')

        #Market Cap rate row
        data_for_overview_table[5] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Market Cap Rate',
                                                                'YoY Market Cap Rate Growth',
                                                                'QoQ Market Cap Rate Growth',
                                                                '%',
                                                                'bps',
                                                                'bps',
                                                                'Market Cap Rate')

        #Transaction Count row
        data_for_overview_table[6] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Sales Volume Transactions',
                                                                'YoY Transactions Growth',
                                                                'QoQ Transactions Growth',
                                                                '',
                                                                '%',
                                                                '%',
                                                                'Transaction Count')

        #Sales volume row
        data_for_overview_table[7] =    CreateRowDataForTable(df_market_cut,
                                                                df_primary_market,
                                                                df_nation,
                                                                'Total Sales Volume',
                                                                'YoY Total Sales Volume Growth',
                                                                'QoQ Total Sales Volume Growth',
                                                                '$',
                                                                '%',
                                                                '%',
                                                                'Sales Volume')

    return(data_for_overview_table)
    
def GetRentTable():
    #Create data for rent Table
    if sector == 'Multifamily':
        return(CreateRowDataForWideTable(data_frame = df_market_cut, data_frame2 = df_primary_market, data_frame3 = df_nation,data_frame4 = df_slices,var1 = 'Market Effective Rent/Unit',modifier = '$',sector=sector))

    else:
        return(CreateRowDataForWideTable(data_frame = df_market_cut, data_frame2 = df_primary_market, data_frame3 = df_nation,data_frame4 = df_slices,var1 = 'Market Rent/SF',modifier = '$',sector=sector))



def CreateMarketReport():
    global market_clean,market_title,output_directory,map_directory
    global df_market_cut,df_primary_market,df_nation,df_slices
    global latest_quarter,document,data_for_overview_table,data_for_vacancy_table,data_for_rent_table,report_path
    
    # remove slashes from market names so we can save as folder name
    market_clean = CleanMarketName(market)
    market_title = market_clean.replace(primary_market + ' -','').strip()
    

    output_directory = CreateOutputDirectory()
    map_directory    = CreateMapDirectory()

 
    #Create seperate dataframes with only rows from the current (sub)market, the primary market, and the nation 
    df_market_cut     = df[df['Geography Name'] == market].copy()                  #df for the market or submarket only
    df_primary_market = df[df['Geography Name'] == primary_market].copy()          #df for the market only
    df_nation         = df[df['Geography Type'] == 'National'].copy()              #df for the USA
    df_slices         = df2[df2['Geography Name'] == primary_market].copy()        #df for the primary market with the quality/subtype slices

    #This function calls all the graph functions defined in the Graph_Functions.py file
    CreateAllGraphs(df_market_cut,df_primary_market,df_nation,output_directory,market_title,primary_market,sector) 
    
    #Get latest quarter
    latest_quarter = df_market_cut.iloc[-1]['Period']


    #Create Data for overview table
    #There are 4 possible permuations for this table (market/apt, market/nonapt, submarket/apt, submakert/nonapt)
    data_for_overview_table = GetOverviewTable()

    #Create data for vacancy Table
    data_for_vacancy_table = CreateRowDataForWideTable(data_frame = df_market_cut, data_frame2 = df_primary_market, data_frame3 = df_nation, data_frame4 = df_slices,var1 = 'Vacancy Rate',modifier = '%',sector=sector)
    
    #Create data for rent Table
    data_for_rent_table = GetRentTable()

    #Get language for paragraphs 
    GetLanguage()

    report_path = CreateReportFilePath()

    #Skip the reports we have already done
    if os.path.exists(report_path.replace('_draft','_FINAL',1)):
        pass
    else:
        #Start writing report
        document = Document()
        SetPageMargins()
        SetStyle()
        MakeReportTitle()
        MakeCoStarDisclaimer()
        AddMap()       
        OverviewSection()
        SupplyDemandSection()
        RentSecton()
        ConstructionSection()
        CapitalMarketsSection()
        OutlookSection()

        # Save report
        document.save(report_path)

    
    #Report writing done, delete figures
    CleanUpPNGs()

    #Add to lists that track our markets and submarkets for salesforce
    UpdateSalesforceMarketList(markets_list = dropbox_primary_markets, submarkets_list = dropbox_markets, sector_list = dropbox_sectors, sector_code_list = dropbox_sectors_codes, dropbox_links_list = dropbox_links)






#Define these empty lists we will fill during the loops, this is to create a list of markets and submarkets and their dropbox links for Salesforce mapping
CreateEmptySalesforceLists()


#Loop through the 4 dataframes, get list of unique markets, loop through those markets creating folders and writing market reports
for df,df2,sector in zip(      [df_multifamily, df_office, df_retail, df_industrial],
                               [df_multifamily_slices, df_office_slices, df_retail_slices, df_industrial_slices],
                               ['Multifamily','Office','Retail','Industrial']):
    print('--',sector,'--')

    #Create dictionary with each market as key and a list of its submarkets as items
    market_dictionary = CreateMarketDictionary(df)

    #Loop through the market dictionary creating reports for each market and their submarkets
    for primary_market,submarkets in market_dictionary.items():

        state = primary_market[-2:] #Get State to make folder that stores markets

        if  primary_market != 'New Haven - CT' :    
            pass
            # continue
        print(primary_market)

        primary_market_clean         = CleanMarketName(primary_market)
        primary_market_name_for_file = primary_market_clean.replace(' - ' + state,'' ).strip() #Make a string with just name of market (without the '- STATECODE' portion)

        #"market" is the general variable name used in all functions for the market OR submarket we are doing report for   
        market = primary_market 
        CreateMarketReport()

        for submarket in submarkets:
            market = submarket
            print(submarket)
            CreateMarketReport()
            


    


#Now create dataframe with list of markets and export to a CSV for Salesforce
dropbox_df = pd.DataFrame({"Market":dropbox_primary_markets,
                           "Submarket":dropbox_markets,
                           'Market Research Name':dropbox_research_names,
                           'Analysis Type': dropbox_analysis_types,
                           'State':         dropbox_states,
                           "Property Type":dropbox_sectors,
                           'Property Type Code':dropbox_sectors_codes,
                           "Dropbox Links":dropbox_links,
                           'Version':dropbox_versions,
                           'Status':dropbox_statuses,
                           'Document Name': dropbox_document_names})


#We are now going to merge our dataframe with the list of markets and submarkets with the zip codes associated with each market and submarket
#We first import and clean that zip code level dataset (convert to one row per submarket with a list of zip codes in it)
df_zipcodes  = pd.read_excel(os.path.join(costar_data_location,'Zip to Submarket.xlsx'), dtype={'PostalCode': object} ) 
df_zipcodes  = df_zipcodes.loc[(df_zipcodes['PropertyType'] == 'Office') | (df_zipcodes['PropertyType'] == 'Retail') | (df_zipcodes['PropertyType'] == 'Industrial') | (df_zipcodes['PropertyType'] == 'Multi-Family')]
df_zipcodes.loc[df_zipcodes['PropertyType'] == 'Multi-Family', 'PropertyType'] = 'Multifamily'
df_zipcodes['state'] = df_zipcodes['MarketName'].str[-2:]
df_zipcodes['SubmarketName']  = df_zipcodes['SubmarketName'].apply(CleanMarketName)

df_zipcodes['Market Research Name'] =  df_zipcodes['state'] + ' - ' + df_zipcodes['SubmarketName'] +  ' - ' +  df_zipcodes['PropertyType'] #form a variable to match on
df_zipcodes  = df_zipcodes.groupby(['Market Research Name'])['PostalCode'].apply(list)
df_zipcodes  = df_zipcodes.reset_index()

#Now merge the zip code data with our costar markets csv
dropbox_df = pd.merge(dropbox_df, df_zipcodes, on='Market Research Name',how = 'left') 

#Now get all the zip codes associated with all the submarkets in each market and place them in the zipcode field for our markets
#Split into 2 dataframes, one with submarkets one with markets then append them together
dropbox_submarkets_df              =  dropbox_df.loc[dropbox_df['Analysis Type'] == 'Submarket'] 
dropbox_markets_df                 =  dropbox_df.loc[dropbox_df['Analysis Type'] == 'Market'] 
dropbox_markets_collapsed_df       =  dropbox_df.groupby(['Market','Property Type']).agg({'PostalCode':list}).reset_index()

#Drop the postalcode column in the markets dataframe and replace it with the collapsed dataframe via merge
dropbox_markets_df = dropbox_markets_df.drop(columns =['PostalCode'])
dropbox_markets_df = pd.merge(dropbox_markets_df, dropbox_markets_collapsed_df, on=['Market','Property Type'],how = 'left') 

#Now that we have added the zip codes to the markets dataframe, we will flaten the list of lists currently in the PostalCode field and remove duplicates
# function used for removing nested 
# lists in python. 
def reemovNestings(l):
    for i in l:
        if type(i) == list:
            reemovNestings(i)
        else:
            output.append(i)
    return(output)

def unique(list1):
    x = np.array(list1)
    x = np.unique(x)
    x = x[x!='nan']
    return(x)

for i in range(len(dropbox_markets_df)):
    output = []
    dropbox_markets_df['PostalCode'].iloc[i] = reemovNestings(dropbox_markets_df['PostalCode'].iloc[i])
    dropbox_markets_df['PostalCode'].iloc[i] = unique(dropbox_markets_df['PostalCode'].iloc[i])
    
#edit the zip code list format in the markets dataframe
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].astype(str)
print(dropbox_markets_df['PostalCode'])
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.replace("""' '""",',')
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.replace("""'
 '""",',')
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.replace("""'""",'')
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.replace("""nan""",'')
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.replace("""[[]]""",'')
dropbox_markets_df['PostalCode'] = dropbox_markets_df['PostalCode'].str.strip()

#edit the zip code list format in the submarkets dataframe
dropbox_submarkets_df['PostalCode'] = dropbox_submarkets_df['PostalCode'].astype(str)
dropbox_submarkets_df['PostalCode'] = dropbox_submarkets_df['PostalCode'].str.replace("""'""",'')
dropbox_submarkets_df['PostalCode'] = dropbox_submarkets_df['PostalCode'].str.replace(""", """,',')
dropbox_submarkets_df['PostalCode'] = dropbox_submarkets_df['PostalCode'].str.replace("""nan""",'[]')
dropbox_submarkets_df['PostalCode'] = dropbox_submarkets_df['PostalCode'].str.strip()

#Now put the split dataframes back together
dropbox_df = dropbox_submarkets_df.append(dropbox_markets_df,ignore_index=True)
dropbox_df = dropbox_df.sort_values(by=['Property Type','Market','Submarket']).reset_index()
dropbox_df = dropbox_df.drop(columns =['index'])
#Export the CoStar Markets export    
dropbox_df.to_csv(os.path.join(output_location,'CoStar Markets.csv'))

print('Finished')
print("--- %s seconds ---" % (time.time() - start_time))        



