import os
import time

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

#Set Graph Size
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
legend_position = 1.05

#Paper color
paper_backgroundcolor = 'white'

#Title Position
title_position = .95

#Sales volume is same for both MF and non MF
def CreateSalesVolumeGraph(data_frame,folder):
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Add Bars with Sales Volume
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame['Total Sales Volume'],
           name="Sales Volume (L)",
           marker_color="#D7DEEA")
            ,secondary_y=False
            )

    # Add scatter points for transaction counts
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame['Sales Volume Transactions'],
            name='Transaction Count (R)',
            marker=dict(color="#4160D3", size=9),
            mode = 'markers'
            )
    ,secondary_y=True)  
 
    #Default behavior for scatter plots in this package is to give some space between origin and first dot, this corrects that
    fig.update_layout(xaxis_range=[-1,len(data_frame['Period'])])



   

    #Set formatting 
    fig.update_layout(
    title_text="Sales Volume & Transaction Count",    
    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},

    yaxis = dict(rangemode = 'tozero'),
    
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                ),
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"    
                    )

    #Add $ to left axis ticks
    fig.update_yaxes(tickfont = dict(size=tickfont_size), tickprefix = '$',  title = None,                             secondary_y=False)                 #left axis
    fig.update_yaxes(tickfont = dict(size=tickfont_size), tickformat=',d',  title = None, separatethousands= True,    secondary_y=True)                  #right axis
    
    #Set x axis ticks
    #Get list with number of quarters
    quarter_list = [i for i in range(len(data_frame['Period']))]
    quarter_list = quarter_list[0::4]

    quarter_list_text = [period for period in data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
        )

    #Set size
    fig.update_layout(
    autosize=False,
    height    = graph_height,
    width     = graph_width,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin,pad=0,autoexpand = True),)
    

    fig.update_layout(margin = dict(r=0))
    fig.write_image(os.path.join(folder,'sales_volume.png'),engine='kaleido',scale=scale)

def CreateAssetValueGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector):
    #Create graph for non-multifamily with construction levels

    #Define the MF variables and labels vs the non MF
    if sector == 'Multifamily':
        asset_value_var =  'Asset Value/Unit'
        asset_value_lab =  'Asset Value/Unit (L)'
    else:
        asset_value_var =  'Asset Value/Sqft'
        asset_value_lab =  'Asset Value/SF (L)'
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    #Add Bars with Asset Values
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame[asset_value_var],
           name= asset_value_lab,
           marker_color="#D7DEEA")
            ,secondary_y=False
            )
   
    # Add line with market cap rate for market
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame['Market Cap Rate'],
            name=market_title,
            line=dict(color="#4160D3"))
    ,secondary_y=True)  

    #If its a submarket, add primary market cap rate. If it's a primary market, add national cap rate line
    if data_frame.equals(data_frame2):
        name = data_frame3['Geography Name'].iloc[0]
        if name == 'United States of America':
            name = 'National'

        if  primary_market != 'United States of America':
            fig.add_trace(
            go.Scatter(x=data_frame3['Period'],
            y=data_frame3['Market Cap Rate'],
            name=name,
            line=dict(color="#B3C3FF"))
            ,secondary_y=True)  
    
    else:
        fig.add_trace(
        go.Scatter(x=data_frame2['Period'],
        y=data_frame2['Market Cap Rate'],
        name=primary_market,
        line=dict(color="#B3C3FF"))
        ,secondary_y=True)      
  
    #Set formatting 
    fig.update_layout(
    title_text="Asset Value & Market Cap Rates",    
    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},

    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                ),
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    height    = graph_height,
    width     = graph_width,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White" 
                    )


    #Set x axis format
    #Get list with number of quarters
    quarter_list = [i for i in range(len(data_frame['Period']))]
    quarter_list = quarter_list[0::4]

    quarter_list_text = [period for period in data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
        )
    
    #Set y axis format
    #Add % to right axis ticks and $ to left axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f', secondary_y=True)  #right axis  
    fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size),secondary_y=False) #left axis
    fig.update_yaxes(automargin = True)
    fig.update_xaxes(automargin = True)    

   
    fig.update_layout(margin = dict(r=0))
    
    #Export figure as PNG file
    fig.write_image(os.path.join(folder,'asset_values.png'),engine='kaleido',scale=scale)

def CreateAbsorptionGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector):
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Add Bars with inventory growth
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame['Inventory Growth'],
           name="Inventory Growth (L)",
           marker_color="#A6B0BF")
            ,secondary_y=False
            )

    #Add Bars with net absorption rate 
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame['Absorption Rate'],
           name="Net Absorption (L)",
           marker_color="#D7DEEA")
            ,secondary_y=False
            )
    
    #Vacancy Rate for (sub)market
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame['Vacancy Rate'],
            name=market_title,
            line=dict(color="#4160D3"))
    ,secondary_y=True)

    #Market
    if data_frame.equals(data_frame2):
        name = data_frame3['Geography Name'].iloc[0]
        if name == 'United States of America':
            name = 'National'

        if  primary_market != 'United States of America':
            fig.add_trace(
            go.Scatter(x=data_frame3['Period'],
            y=data_frame3['Vacancy Rate'],
            name=name,
            line=dict(color="#B3C3FF"))
            ,secondary_y=True)  
    
    #Submarket
    else:
        fig.add_trace(
        go.Scatter(x=data_frame2['Period'],
        y=data_frame2['Vacancy Rate'],
        name=primary_market,
        line=dict(color="#B3C3FF")
                 )
        ,secondary_y=True
                    )     
    
    
    #Set x axis ticks
    quarter_list = [i for i in range(len(data_frame['Period']))] #Get list with number of quarters
    quarter_list = quarter_list[0::4]

    quarter_list_text = [period for period in data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
        )

    #Set Title
    fig.update_layout(
    title_text="Absorption & Vacancy Rates",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
                    )



    #Add % to axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f', secondary_y=True)
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)


    
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin = dict(r=0))
    fig.write_image(os.path.join(folder,'absorption_rate.png'),engine='kaleido',scale=scale)
           
def CreateConstructionGraph(data_frame,folder,sector):
    #Create graph for non-multifamily with construction levels
    
    #Define the MF variables and labels vs the non MF
    if sector == 'Multifamily':
        construction_var =   'Under Construction Units'
        construction_lab =   "Under Construction (L)"
        title            =   "Under Construction Units - Share of Inventory"
    else:
        construction_var = 'Under Construction SF'
        construction_lab =  "Under Construction (L)"
        title            = 'Under Construction SF - Share of Inventory'


    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])



    #Add Bars with SF Under Construction 
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame[construction_var],
           name=construction_lab,
           marker_color="#D7DEEA")
            ,secondary_y=False
            )
    
    # Add line with share of inventory under construction
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame['Under Construction %'],
            name="Under Construction - Share of Inventory (R)",
            line=dict(color="#4160D3"))
            ,secondary_y=True)      
  
    #Set formatting 
    fig.update_layout(
    title_text=title,    
    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},

    yaxis = dict(rangemode = 'tozero'),

    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                ),
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    height    = graph_height,
    width     = graph_width,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"    
                    )

                   
    #Y-axis format(add % to right axis ticks)
    fig.update_yaxes(ticksuffix = '%', 
                     tickfont = dict(size=tickfont_size), 
                     tickformat='.1f',
                     secondary_y=True)
    

    fig.update_yaxes(tickfont = dict(size=tickfont_size), 
                    separatethousands= True,
                    tickformat=',d',
                     secondary_y=False)

    


    #Set x axis ticks
    #Get list with number of quarters
    quarter_list = [i for i in range(len(data_frame['Period']))]
    quarter_list = quarter_list[0::4]

    quarter_list_text = [period for period in data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
        )

        
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin = dict(r=0))

    fig.write_image(os.path.join(folder,'construction_volume.png'),engine='kaleido',scale=scale)
    return(fig)

def CreateRentGrowthGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector):
    #Create graph for rent growth

    #Define the MF variables and labels vs the non MF
    if sector == 'Multifamily':
        rent_var                  = 'Market Effective Rent/Unit'
        quarterly_rent_growth_var = 'QoQ Market Effective Rent/Unit Growth'
        annual_rent_growth_var    = 'YoY Market Effective Rent/Unit Growth'
        quarterly_rent_growth_lab = "Quarterly Growth (L)"
        annual_rent_growth_lab    = "Annual Growth (L)"
        title                     = "Market Effective Rent/Unit - Annual & Quarterly Growth"
    else:
        rent_var                  = 'Market Rent/SF'
        quarterly_rent_growth_var = 'QoQ Rent Growth'
        annual_rent_growth_var    = 'YoY Rent Growth' 
        quarterly_rent_growth_lab = "Quarterly Growth (L)"
        annual_rent_growth_lab    = "Annual Growth (L)"
        title                     = "Market Rent/SF - Annual & Quarterly Growth"

    
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])
   
    #Add Dots with Quarterly Rent Growth
    fig.add_trace(
            go.Scatter(
                        x=data_frame['Period'],
                        y=data_frame[quarterly_rent_growth_var],
                        name=quarterly_rent_growth_lab,
                        marker=dict(color="#4160D3", size=9),
                        mode = 'markers',
                      )
        ,secondary_y=False
                 )
    
    #Default behavior for scatter plots in this package is to give some space between origin and first dot, this corrects that
    fig.update_layout(xaxis_range=[-1,len(data_frame['Period'])])
    


    #Add Bars with Annual Rent Growth
    fig.add_trace(
    go.Bar(x=data_frame['Period'],
           y=data_frame[annual_rent_growth_var],
           name=annual_rent_growth_lab,
           marker_color="#D7DEEA",
           base = dict(layer = 'Below'))
            ,secondary_y=False
            )


    # Add line with vacancy rate for market
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame['Vacancy Rate'],
            name='Vacancy Rate (L)',
            line=dict(color="#404858",dash='dash'))
    ,secondary_y=False)  

       
    # Add line with rent for market
    fig.add_trace(
    go.Scatter(x=data_frame['Period'],
            y=data_frame[rent_var],
            name=market_title,
            line=dict(color="#4160D3",dash='solid'))
    ,secondary_y=True) 

    #If its a submarket, add primary market rent. If it's a primary market, add national rent line
    if data_frame.equals(data_frame2):
        extra_height = 70
        name = data_frame3['Geography Name'].iloc[0]
        if name == 'United States of America':
            name = 'National'
        
        if  primary_market != 'United States of America':
            fig.add_trace(
            go.Scatter(x=data_frame3['Period'],
            y=data_frame3[rent_var],
            name=name,
            line=dict(color="#B3C3FF"))
            ,secondary_y=True)  
    
    else: #Submarkets
        extra_height = 70

        fig.add_trace(
        go.Scatter(x=data_frame2['Period'],
        y=data_frame2[rent_var],
        name=primary_market,
        line=dict(color="#B3C3FF"))
        ,secondary_y=True)      
    


    #Add % to left axis ticks and $ to right axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size), tickformat='.1f',                        secondary_y=False)  #right axis  
    fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size), separatethousands= True,secondary_y=True)   #left axis

    #Set x axis ticks
    #Get list with number of quarters
    quarter_list = [i for i in range(len(data_frame['Period']))]
    quarter_list = quarter_list[0::4]

    quarter_list_text = [period for period in data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
        )
 
    #Set formatting 
    fig.update_layout(
    title_text=title,    
    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
    
    #legend format
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position-.05,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                ),
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    height    = graph_height + extra_height , #submarket version needs to be taller to fit larger legend
    width     = graph_width ,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"    
                    )

    
        
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin = dict(r=0))

    fig.write_image(os.path.join(folder,'rent_growth.png'),engine='kaleido',scale=scale)


def CreateAllGraphs(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector):

    if primary_market == 'Manhattan - NY':
        primary_market = 'Manhattan'

    if market_title == 'Manhattan - NY':
        market_title = 'Manhattan'

    CreateSalesVolumeGraph(data_frame,folder)
    CreateAssetValueGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector)
    CreateAbsorptionGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector)
    CreateConstructionGraph(data_frame,folder,sector)
    CreateRentGrowthGraph(data_frame,data_frame2,data_frame3,folder,market_title,primary_market,sector)

    for png in ['asset_values.png','absorption_rate.png','sales_volume.png','construction_volume.png','rent_growth.png']:
        png_path = os.path.join(folder,png)
        while os.path.exists(png_path)  == False: #Wait till figure is saved before moving on, this avoids file path errors
            pass


