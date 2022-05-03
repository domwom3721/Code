#Date: 5/2/2022
#Author: Mike Leahy
#Summary: Injests RedFin residential real estate data and produces report documents on the selcted areas

import os
from pydoc import doc
from tkinter.tix import MAIN
import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from pyrsistent import v

#Define file pre-paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project') 
data_location                  =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Clean') 
output_location                =  os.path.join(project_location,'Output\Residential Reports') #Testing Output
# output_location                =  os.path.join(project_location,'Output\Residential Reports') #Production Output


#Language Related functions
def OverviewLanguage():
    try:
        return(['Overview language'])
    except Exception as e:
        print(e, 'Unable to create overview language')

def SupplyDemandLanguage():
    try:
        return(['Supply and Demand language'])
    except Exception as e:
        print(e, 'Unable to create supply and demand language')

def ValuesLanguage():
    try:
        return(['Values language'])
    except Exception as e:
        print(e, 'Unable to create values language')

def ConclusionLanguage():
    try:
        return(['Outlook language'])
    except Exception as e:
        print(e, 'Unable to create conclusion language')

def CreateLanguage():
    print('Creating Language')
    global overview_language, supply_and_demand_language, values_language, conclusion_language
    overview_language          = OverviewLanguage()
    supply_and_demand_language = SupplyDemandLanguage()
    values_language            = ValuesLanguage()
    conclusion_language        = ConclusionLanguage()

#Graph related functions
def CreateGraphs():
    print('Creating Graphs')
    pass

#Document related functions
def SetPageMargins(document, margin_size):
    sections = document.sections
    for section in sections:
        section.top_margin    = Inches(margin_size)
        section.bottom_margin = Inches(margin_size)
        section.left_margin   = Inches(margin_size)
        section.right_margin  = Inches(margin_size)

def SetDocumentStyle(document):
    style     = document.styles['Normal']
    font      = style.font
    font.name = 'Avenir Next LT Pro (Body)'
    font.size = Pt(9)

def AddTitle(document):
    title_text = 'Geography Name Property Type Market Analysis'
    title                               = document.add_heading(title_text,level=1)
    title.style                         = document.styles['Heading 2']
    title.paragraph_format.space_after  = Pt(6)
    title.paragraph_format.space_before = Pt(12)
    title_style                         = title.style
    title_style.font.name               = "Avenir Next LT Pro Light"
    title_style.font.size               = Pt(14)
    title_style.font.bold               = False
    title_style.font.color.rgb          = RGBColor.from_string('3F65AB')
    title_style.element.xml
    rFonts                              = title_style.element.rPr.rFonts
    rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro Light")

    above_map_paragraph = document.add_paragraph("""This report was created using data from Redfin, a national real estate brokerage. Data represents Property_Type's in "Geography_Name" with monthly data through current_period.""")
    above_map_style                                   = above_map_paragraph.style
    above_map_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
    above_map_style.font.size                         = Pt(9)
    above_map_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

def AddMap(document):
    #Add image of map if we already have one
    map_path = os.path.join(output_location, 'map.png')
    if os.path.exists(map_path):
        print('Adding map png to document')
        map = document.add_picture(map_path, width=Inches(6.5))
    Citation(document=document,text= 'Google Maps')

def AddHeading(document, title, heading_level): 
    #Function we use to insert the headers other than the title header
    heading                               = document.add_heading(title, level = heading_level)
    heading.style                         = document.styles['Heading 3']
    heading_style                         = heading.style
    heading_style.font.name               = "Avenir Next LT Pro"
    heading_style.font.size               = Pt(11)
    heading_style.font.bold               = False
    heading.paragraph_format.space_after  = Pt(6)
    heading.paragraph_format.space_before = Pt(12)

    #Color
    heading_style.font.color.rgb          = RGBColor.from_string('3F65AB')            
    heading_style.element.xml
    rFonts                                = heading_style.element.rPr.rFonts
    rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro")

def AddDocumentParagraph(document, language_variable):
    assert type(language_variable) == list

    for paragraph in language_variable:
        
        #Skip blank paragraphs
        if paragraph == '':
            continue
        
        par                                               = document.add_paragraph(str(paragraph))
        par.alignment                                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.paragraph_format.space_after                  = Pt(primary_space_after_paragraph)
        summary_format                                    = document.styles['Normal'].paragraph_format
        summary_format.line_spacing_rule                  = WD_LINE_SPACING.SINGLE
        style                                             = document.styles['Normal']
        font                                              = style.font
        font.name                                         = primary_font
        par.style                                         = document.styles['Normal']

def AddDocumentPicture(document, image_path, citation):
    if os.path.exists(image_path):
        fig                                         = document.add_picture(os.path.join(image_path),width=Inches(6.5))
        last_paragraph                              = document.paragraphs[-1] 
        last_paragraph.paragraph_format.space_after = Pt(0)
        last_paragraph.alignment                    = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document, citation)

def Citation(document, text):
    citation_paragraph                               = document.add_paragraph()
    citation_paragraph.paragraph_format.space_after  = Pt(6)
    citation_paragraph.paragraph_format.space_before = Pt(0)
    run                                              = citation_paragraph.add_run('Source: ' + text)
    font                                             = run.font
    font.name                                        = primary_font
    font.size                                        = Pt(8)
    font.italic                                      = True
    font.color.rgb                                   = RGBColor.from_string('929292')
    citation_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.RIGHT

def AddTableTitle(document, title):
    table_title_paragraph                               = document.add_paragraph(title)
    table_title_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.CENTER
    table_title_paragraph.paragraph_format.space_after  = Pt(6)
    table_title_paragraph.paragraph_format.space_before = Pt(12)

    for run in table_title_paragraph.runs:
                    font      = run.font
                    font.name = 'Avenir Next LT Pro Medium'

def OverviewSection(document):
    print('Writing Overview Section')
    AddHeading(document = document, title = 'At a Glance', heading_level = 2)

    #Add Overview langauge
    AddDocumentParagraph(document = document, language_variable = overview_language)

    AddTableTitle(document=document,title = 'Market Fundamentals')

def SupplyandDemandSection(document):
    print('Writing Supply and Demand Section')
    AddHeading(document = document, title = 'Supply and Demand', heading_level = 2)

    #Add Overview langauge
    AddDocumentParagraph(document = document, language_variable = supply_and_demand_language)

def ValuesSection(document):
    print('Writing Values Section')
    AddHeading(document = document, title = 'Values', heading_level = 2)

    #Add Overview langauge
    AddDocumentParagraph(document = document, language_variable = values_language)

def ConclusionSection(document):
    print('Writing Conclusion Section')
    AddHeading(document = document, title = 'Conclusion', heading_level = 2)

    #Add Overview langauge
    AddDocumentParagraph(document = document, language_variable = conclusion_language)

def WriteReport():
    print('Writing Report')
    #Create Document
    document = Document()
    SetPageMargins(       document = document, margin_size = 1)
    SetDocumentStyle(     document = document)
    AddTitle(             document = document)
    AddMap(               document=    document)
    OverviewSection(      document = document)
    SupplyandDemandSection(document = document)
    ValuesSection(document = document)
    ConclusionSection(document = document)
    
    
    #Save report
    document.save(os.path.join(output_location,'test.docx'))  

def Main():
    CreateLanguage()
    CreateGraphs()
    WriteReport()

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

#Set formatting paramaters for reports
primary_font                  = 'Avenir Next LT Pro Light' 
primary_space_after_paragraph = 8

#Set graph size and format variables
tickangle                     = 0
bowery_grey                   = "#D7DEEA"
bowery_dark_grey              = "#A6B0BF"
bowery_dark_blue              = "#4160D3"
bowery_light_blue             = "#B3C3FF"
bowery_black                  = "#404858"


Main()