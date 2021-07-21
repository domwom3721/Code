#By Mike Leahy
#Started 06/30/2021
#Summary: This script creates reports on neighborhoods/cities for Bowery

import math
import os
import time
from datetime import date
from pprint import pprint
from random import randrange

import docx
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import pyautogui
import requests
import wikipedia
from bls_datasets import oes, qcew
from blsconnect import RequestBLS, bls_search
from census import Census
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Row, CT_Tc
from docx.shared import Inches, Pt, RGBColor
from fredapi import Fred
from numpy import true_divide
from plotly.subplots import make_subplots
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from us import states
from wikipedia.wikipedia import random

#Define file paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects', 'Research Report Automation Project') 
main_output_location           =  os.path.join(project_location,'Output','Neighborhood Reports') #testing
# main_output_location           =  os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood') #production
data_location                  =  os.path.join(project_location,'Data','Neighborhood Reports Data')
graphics_location              =  os.path.join(project_location,'Data','Graphics')
map_location                   =  os.path.join(project_location,'Data','Maps','Neighborhood Maps')

c    = Census('18335344cf4a0242ae9f7354489ef2f8860a9f61')

def CreateDirectory():

    global report_path,hood_folder_map,hood_folder
    
    state_folder_map         = os.path.join(map_location,state)
    hood_folder_map          = os.path.join(map_location,state,neighborhood)
    
    state_folder             = os.path.join(main_output_location,state)
    hood_folder              = os.path.join(main_output_location,state,neighborhood)

    for folder in [state_folder,hood_folder,state_folder_map,hood_folder_map]:
         if os.path.exists(folder):
            pass 
         else:
            os.mkdir(folder) 

    report_path = os.path.join(hood_folder,current_year + ' ' + state + ' - ' + neighborhood  + ' - hood' + '_draft.docx')

def GetDataAndLanguageForOverviewTable():
    

    return([ ['','Area','2000 Census','2010 Census','Change','2021 Est.','Change','2026 Projected','Change'],
             ['Population',neighborhood,'','','','','','',''],
             ['',comparison_area,'','','','','','',''],
             ['Households',neighborhood,'','','','','','',''],
             ['',comparison_area,'','','','','','',''],
             ['Family Households',neighborhood,'','','','','','',''],
             ['',comparison_area,'','','','','','',''],
              ])

def GetCountySubdivisionData():
    print('Getting City/County Subdivision Data')
    total_county_subdivision_population = c.sf1.state_county_subdivision(fields = ['H010001'], state_fips = '34', county_fips = '017',subdiv_fips='32250')[0]['H010001']
    print(total_county_subdivision_population)

def GetCensusPlaceData(state_fips, place_fips):
    print('Getting Census Place (City) Data')
    population_field     = 'H010001'
    name_field           = 'NAME'
    total_census_place_population = int(c.sf1.state_place(fields=[population_field],state_fips=state_fips,place=place_fips)[0][population_field])
    census_place_name = c.sf1.state_place(fields=[name_field],state_fips=state_fips,place=place_fips)[0][name_field]

    census_place_number_1_people_households = c.sf1.state_place(fields=['H013002'],state_fips=state_fips,place=place_fips)[0]['H013002']
    census_place_number_2_people_households = c.sf1.state_place(fields=['H013003'],state_fips=state_fips,place=place_fips)[0]['H013003']
    census_place_number_3_people_households = c.sf1.state_place(fields=[''],state_fips=state_fips,place=place_fips)[0]['']
    census_place_number_4_people_households = c.sf1.state_place(fields=[''],state_fips=state_fips,place=place_fips)[0]['']
    census_place_number_5_people_households = c.sf1.state_place(fields=[''],state_fips=state_fips,place=place_fips)[0]['']
    census_place_number_6_people_households = c.sf1.state_place(fields=[''],state_fips=state_fips,place=place_fips)[0]['']
    census_place_number_7_people_households = c.sf1.state_place(fields=[''],state_fips=state_fips,place=place_fips)[0][''] #7 or more


# H013004	Total!!3-person household	HOUSEHOLD SIZE	not required		0	int	H13
# H013005	Total!!4-person household	HOUSEHOLD SIZE	not required		0	int	H13
# H013006	Total!!5-person household	HOUSEHOLD SIZE	not required		0	int	H13
# H013007	Total!!6-person household	HOUSEHOLD SIZE	not required		0	int	H13
# H013008	Total!!7-or-more-person household

    print(total_census_place_population)
    print(census_place_name)






def GetCountyData():
    print('Getting County Data')
    #Get data on county
    # total_county_population = c.sf1.state_county(fields = ['H010001'], state_fips = '34', county_fips = '017')[0]['H010001']
    # print(total_county_population)

def GetData():
    print('Getting Data')



    
    

def CreateGraphs():
    print('Creating Graphs')
    pass
    
def CarLanguage():
    print('Writing Car Langauge')
    
    major_highways                = page.section('Major highways')
    major_Highways                = page.section('Major Highways')
    roadways                      = page.section('Roadways')
    highways                      = page.section('Highways')
    public_roadways               = page.section('Public roadways')
    major_roads                   = page.section('Major roads and highways')
    roads_and_highways            = page.section('Roads and highways')
    major_roads_and_Highways      = page.section('Major roads and Highways')
    car_language = ''
    for count,section in enumerate([major_highways,major_Highways,roadways,highways,public_roadways,major_roads,roads_and_highways,major_roads_and_Highways]):
        if (section != None) and (count == 0):
            car_language =  section 
        elif (section != None) and (count > 0):
            car_language = car_language + ' ' + "\n" + section 

    

    #If the wikipedia page is missiing all highway sections 
    if car_language == '':
        return(neighborhood + ' is not connected by any major highways or roads.')
    else:
        return(car_language)

def PlaneLanguage():
    print('Writing Plane Langauge')
    #Go though some common section names for airports
    airports              = page.section('Airports')
    air                   = page.section('Air')
    aviation              = page.section('Aviation')

    plane_language = ''
    for count,section in enumerate([airports,air,aviation]):
        if (section != None) and (count == 0):
            plane_language =  section 
        elif (section != None) and (count > 0):
            plane_language = plane_language + ' ' + "\n" + section 

    

    #If the wikipedia page is missiing all airport sections 
    if plane_language == '':
        return(neighborhood + 'is not served by any airport.')
    else:
        return(plane_language)
    
    
       
    
def BusLanguage():
    print('Writing Bus Langauge')
    bus                          =  page.section('Bus')
    intercity_bus                =  page.section('Intercity buses')
    public_Transportation        =  page.section('Public Transportation')
    
    #Add the text from the sections above to a single string variable
    bus_language = ''
    for count,section in enumerate([bus,intercity_bus,public_Transportation]):
        if (section != None) and (count == 0):
            bus_language =  section 
        elif (section != None) and (count > 0):
            bus_language = bus_language + ' ' + "\n" + section 

    
    #If the wikipedia page is missiing all airport sections return default phrase
    if bus_language == '':
        return(neighborhood + ' does not have public bus service.')
    else:
        return(bus_language)

   

def TrainLanguage():
    print('Writing Train Langauge')
    rail                         =  page.section('Rail')
    public_transportation        =  page.section('Public transportation')
    public_Transportation        =  page.section('Public Transportation')
    public_transport             =  page.section('Public transport')
    mass_transit                 =  page.section('Mass transit')
    rail_network                 =  page.section('Rail Network')

    #Add the text from the sections above to a single string variable
    train_language = ''
    for count,section in enumerate([rail,public_transportation,public_Transportation,public_transport,mass_transit,rail_network]):
        if (section != None) and (count == 0):
            train_language =  section 
        elif (section != None) and (count > 0):
            train_language = train_language + ' ' + "\n" + section 

    
    #If the wikipedia page is missiing all airport sections return default phrase
    if train_language == '':
        return(neighborhood + ' is not served by any commuter or light rail lines.')
    else:
        return(train_language)

     

def SummaryLangauge():
    return(wikipedia.summary((neighborhood + ',' + state)))

def OutlookLanguage():
    return('Neighborhood analysis can best be summarized by referring to neighborhood life cycles. ' +
          'Neighborhoods are perceived to go through four cycles, the first being growth, the second being stability, the third decline, and the fourth revitalization. ' +
          'It is our observation that the subject’s neighborhood is exhibiting several stages of the economic life, with an overall predominance of stability and both limited decline and limited revitalization in some sectors. ' +
          'The immediate area surrounding the subject, has had a historically low vacancy level and is located just to the south of the ------ submarket,' +
          """ which has multiple office and retail projects completed within the past two years and more development in the subject’s immediate vicinity either under construction or preparing to break ground."""+
          ' The proximity of the ________ and ________ will ensure the neighborhood will continue ' +
          'to attract growth in the long-term.')
    pass

def CreateLanguage():
    print('Creating Langauge')

    global bus_language,car_language,plane_language,train_language,transportation_language,summary_langauge,conclusion_langauge
    transportation_language         =  page.section('Transportation')
    bus_language   = BusLanguage()
    car_language   = CarLanguage()
    plane_language = PlaneLanguage()
    train_language = TrainLanguage()
    summary_langauge =  SummaryLangauge()
    conclusion_langauge = OutlookLanguage()
    pass

#Graph Related Functions
def SetGraphFormatVariables():
    global graph_width, graph_height, scale,tickfont_size,left_margin,right_margin,top_margin,bottom_margin,legend_position,paper_backgroundcolor,title_position

    #Set graph size and format variables
    marginInches = 1/18
    ppi = 96.85 
    width_inches = 6.5
    height_inches = 3.3

    graph_width  = (width_inches - marginInches)   * ppi
    graph_height = (height_inches  - marginInches) * ppi

    #Set scale for resolution 1 = no change, > 1 increases resolution. Very important for run time of main script. 
    scale = 3

    #Set tick font size (also controls legend font size)
    tickfont_size = 8 

    #Set Margin parameters/legend location
    left_margin   = 0
    right_margin  = 0
    top_margin    = 75
    bottom_margin = 10
    legend_position = 1.10

    #Paper color
    paper_backgroundcolor = 'white'

    #Title Position
    title_position = .95    

#Report document related functions
def SetPageMargins(document,margin_size):
    sections = document.sections
    for section in sections:
        section.top_margin    = Inches(margin_size)
        section.bottom_margin = Inches(margin_size)
        section.left_margin   = Inches(margin_size)
        section.right_margin  = Inches(margin_size)

def SetDocumentStyle(document):
    style = document.styles['Normal']
    font = style.font
    font.name = 'Avenir Next LT Pro (Body)'
    font.size = Pt(9)

def AddTitle(document):
    main_title = document.add_heading('Neighborhood & Demographic Overview',level=0) 
    main_title.style = document.styles['Heading 1']
    main_title.paragraph_format.space_after  = Pt(6)
    main_title.paragraph_format.space_before = Pt(12)
    main_title_style = main_title.style
    main_title_style.font.name = "Avenir Next LT Pro Light"
    main_title_style.font.size = Pt(18)
    main_title_style.font.bold = False
    main_title_style.font.color.rgb = RGBColor.from_string('3F65AB')
    main_title_style.element.xml
    rFonts = main_title_style.element.rPr.rFonts
    rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro Light")

    title = document.add_heading(neighborhood + ' at a Glance',level=1)
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

def AddHeading(document,title,heading_level,heading_number,font_size): #Function we use to insert the headers other than the title header
            heading = document.add_heading(title,level=heading_level)
            heading.style = document.styles[heading_number]
            heading_style =  heading.style
            heading_style.font.name = "Avenir Next LT Pro Light"
            heading_style.font.size = Pt(font_size)
            heading_style.font.bold = False
            heading.paragraph_format.space_after  = Pt(6)
            heading.paragraph_format.space_before = Pt(12)

            #Color
            heading_style.font.color.rgb = RGBColor.from_string('3F65AB')            
            heading_style.element.xml
            rFonts = heading_style.element.rPr.rFonts
            rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro")

def Citation(document,text):
    citation_paragraph = document.add_paragraph()
    citation_paragraph.paragraph_format.space_after  = Pt(6)
    citation_paragraph.paragraph_format.space_before = Pt(6)
    run = citation_paragraph.add_run('Source: ' + text)
    font = run.font
    font.name = 'Avenir Next LT Pro (Body)'
    font.size = Pt(8)
    font.italic = True
    font.color.rgb  = RGBColor.from_string('929292')
    citation_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if text != 'Google Maps':
        blank_paragraph = document.add_paragraph('')
        blank_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def AddMap(document):
    print('Adding Map')
    # Add image of map
    try:
        map = document.add_picture(os.path.join(hood_folder_map,'map.png'),width=Inches(6.5))
    except:
        #Search Google Maps for County
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        browser = webdriver.Chrome(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop','chromedriver.exe')),options=options)
        browser.get('https:google.com/maps')
        Place = browser.find_element_by_class_name("tactile-searchbox-input")
        Place.send_keys((neighborhood + ', ' + state))
        Submit = browser.find_element_by_xpath(
        "/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[1]/div[1]/div[2]/div[1]/button")
        Submit.click()
        time.sleep(5)
        zoomout = browser.find_element_by_xpath(
        """/html/body/jsl/div[3]/div[9]/div[22]/div[1]/div[2]/div[7]/div/button""")
        zoomout.click()
        time.sleep(10)
        im2 = pyautogui.screenshot(region=(1089,276, 2405, 1754) ) #left, top, width, and height
        time.sleep(.25)
        im2.save(os.path.join(hood_folder_map,'map.png'))
        im2.close()
        time.sleep(1)
        map = document.add_picture(os.path.join(hood_folder_map,'map.png'),width=Inches(6.5))
        browser.quit()
    finally:
        pass
       

def AddTable(document,data_for_table): #Function we use to insert our overview table into the report document
    #list of list where each list is a row for our table
     
    #make sure each list inside the list of lists has the same number of elements
    for row in data_for_table:
        for row2 in data_for_table:
            assert len(row) == len(row2)


    #create table object
    tab = document.add_table(rows=len(data_for_table), cols=len(data_for_table[0]))
    tab.alignment     = WD_TABLE_ALIGNMENT.CENTER
    tab.allow_autofit = True
    #loop through the rows in the table
    for current_row ,(row,row_data_list) in enumerate(zip(tab.rows,data_for_table)): 

    
        row.height = Inches(0)

        #loop through all cells in the current row
        for current_column,(cell,cell_data) in enumerate(zip(row.cells,row_data_list)):
            cell.text = str(cell_data)

            if current_row == 0:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM


            #set column widths
            if current_column == 0:
                cell.width = Inches(1.25)

            # elif current_column == 1:
            #     cell.width = Inches(1.19)

            # elif current_column == 2:
            #     cell.width = Inches(0.8)



            #add border to top row
            if current_row == 1:
                    tcPr = cell._element.tcPr
                    tcBorders = OxmlElement("w:tcBorders")
                    top = OxmlElement('w:top')
                    top.set(qn('w:val'), 'single')
                    tcBorders.append(top)
                    tcPr.append(tcBorders)
            
            #loop through the paragraphs in the cell and set font and style
            for paragraph in cell.paragraphs:
                if current_column > 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                     paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(8)
                    run.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    #make first row bold
                    if current_row == 0: 
                        font.bold = True
                        font.name = 'Avenir Next LT Pro Demi'
                        

def IntroSection(document):
    AddTitle(document = document)
    AddMap(document = document)
    Citation(document,'Google Maps')
    AddHeading(document = document, title = 'Summary',            heading_level = 2,heading_number='Heading 3',font_size=11)
    
    #Get summary section from wikipedia and add it 
    summary_paragraph           = document.add_paragraph(summary_langauge)
    summary_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    AddTable(document = document,data_for_table = GetDataAndLanguageForOverviewTable() )


def NeigborhoodSection(document):
    print('Writing Neighborhood Section')
    AddHeading(document = document, title = 'Neighborhood',            heading_level = 1,heading_number='Heading 2',font_size=14)
    AddHeading(document = document, title = 'Housing',                  heading_level = 2,heading_number='Heading 3',font_size=11)

def DemographicsSection(document):
    print('Writing Neighborhood Section')
    AddHeading(document = document, title = 'Demographics',                                   heading_level = 1,heading_number='Heading 2',font_size=14)
    AddHeading(document = document, title = 'Population',                                     heading_level = 2,heading_number='Heading 3',font_size=11)
    
    #Employment and Transportation Subsection
    AddHeading(document = document, title = 'Employment and Transportation',                  heading_level = 2,heading_number='Heading 3',font_size=11)

    
    table_paragraph = document.add_paragraph('Transportation Methods')
    table_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    transportation_paragraph = document.add_paragraph(transportation_language)

    #Insert the transit graphics(car, bus,plane, train)
    tab = document.add_table(rows=1, cols=2)
    for pic in ['car.png','train.png','bus.png','plane.png']:
        row_cells = tab.add_row().cells
        paragraph = row_cells[0].paragraphs[0]
        run = paragraph.add_run()
        if pic == 'car.png':
            run.add_text(' ')
        run.add_picture(os.path.join(graphics_location,pic))
    


    transit_language = [car_language,train_language,bus_language,plane_language]
    # transit_language = ['car_language','train_language','bus_language','plane_language']

    #Loop through the rows in the table
    for current_row ,row in enumerate(tab.rows): 
        #loop through all cells in the current row
        for current_column,cell in enumerate(row.cells):
            if current_column == 1 and current_row > 0:
                cell.text = transit_language[current_row-1]

            if current_column == 0:
                cell.width = Inches(.2)
            else:
                cell.width = Inches(6)

 
def OutlookSection(document):
    print('Writing Outlook Section')
    AddHeading(document = document, title = 'Conclusion',            heading_level = 1,heading_number='Heading 2',font_size=14)
    conclusion_paragraph           = document.add_paragraph(conclusion_langauge)
    conclusion_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    


def WriteReport():
    print('Writing Report')
    #Create Document
    document = Document()
    SetPageMargins(document   = document, margin_size=1)
    SetDocumentStyle(document = document)
    IntroSection(document = document)
    NeigborhoodSection(document     = document)
    DemographicsSection(document = document)
    OutlookSection(document = document)


    #Save report
    document.save(report_path)  

def CleanUpPNGs():
    print('Deleting PNG files')
    #Report writing done, delete figures
    files = os.listdir(hood_folder)
    for image in files:
        if image.endswith(".png"):
            os.remove(os.path.join(hood_folder, image))

def Main():
    SetGraphFormatVariables()
    CreateDirectory()
    GetData()
    CreateGraphs()
    CreateLanguage()
    WriteReport()
    CleanUpPNGs()

#Decide if you want to export data in excel files in the county folder
data_export = False


state = 'NJ'
neighborhood = 'Hoboken'

# neighborhood = input('What is the neighborhood?')
# state        = input('What is the 2 letter statecode?')



comparison_area = 'Hudson County, NJ'
# comparison_area = input('What is the name of the comparison area?')

todays_date = date.today()
current_year = str(todays_date.year)
page                          =  wikipedia.page((neighborhood + ',' + state))

# Main()
GetCensusPlaceData(state_fips = '12' , place_fips = '13275')

# https://maps.googleapis.com/maps/api/geocode/json?key=AIzaSyBMcoRFOW2rxAGxURCpA4gk10MROVVflLs&address=90%Jarvis%Place
