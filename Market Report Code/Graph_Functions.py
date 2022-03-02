import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots

#Set Graph Size
marginInches      = 1/18
ppi               = 96.85 
width_inches      = 6.5
height_inches     = 3.3

graph_width       = (width_inches  - marginInches)   * ppi 
graph_height      = (height_inches - marginInches)   * ppi

#Set scale for resolution 1 = no change, > 1 increases resolution. Very important for run time of main script. 
scale             = 6

#Set tick font size (also controls legend font size)
tickfont_size     = 8 

#Set Margin parameters/legend location
left_margin       = 0
right_margin      = 0
top_margin        = 75
bottom_margin     = 10
legend_position   = 1.05
title_position    = .95
font_size         = 10.5
backgroundcolor   = 'white'
bowery_grey       = "#D7DEEA"
bowery_dark_grey  = "#A6B0BF"
bowery_dark_blue  = "#4160D3"
bowery_light_blue = "#B3C3FF"
bowery_black      = "#404858"
font_family       = "Avenir Next LT Pro"
font_color        = '#262626'

def CreateSalesVolumeGraph(submarket_data_frame, folder):
    
    #Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Add bars with sales volume
    fig.add_trace(
    go.Bar(x            = submarket_data_frame['Period'],
           y            = submarket_data_frame['Total Sales Volume'],
           name         = "Sales Volume (L)",
           marker_color = bowery_grey),
           secondary_y  = False
                 )

    #We do not have transcation counts for custom county reports (in cases of properties outside CoStar markets)
    #This prevents the custom reports from having an incorrect title
    if submarket_data_frame['Sales Volume Transactions'].max() > 0:
        title = "Sales Volume & Transaction Count"
        
        #Add scatter points for transaction counts
        fig.add_trace(
        go.Scatter(x           = submarket_data_frame['Period'],
                   y           = submarket_data_frame['Sales Volume Transactions'],
                   name        = 'Transaction Count (R)',
                   marker      = dict(color = bowery_dark_blue, size = 9),
                   mode        = 'markers'),
                   secondary_y = True
                    )  
    
    else:
        title = "Sales Volume"

    #Default behavior for scatter plots in this package is to give some space between origin and first dot, this corrects that
    fig.update_layout(xaxis_range=[-1, len(submarket_data_frame['Period'])])

    #Set formatting 
    fig.update_layout(
        title_text    = title,    
        font_family   = font_family,
        font_color    = font_color,
        font_size     = font_size,
        paper_bgcolor = backgroundcolor,
        plot_bgcolor  = backgroundcolor,   
        
        #Set title
        title = {
            'y':       title_position,
            'x':       0.5,
            'xanchor': 'center',
            'yanchor': 'top',
                },
        
        #Y-axis range
        yaxis = dict(rangemode = 'tozero'),
        
        #Set legend
        legend=dict(
                orientation = "h",
                yanchor     = "bottom",
                y           = legend_position,
                xanchor     = "center",
                x           = 0.5,
                font_size   = tickfont_size
                    ),

                   )

    #Add $ to left axis ticks
    fig.update_yaxes(tickfont = dict(size = tickfont_size), tickprefix = '$', title = None, secondary_y = False)                 
    
    #Format right axis ticks
    fig.update_yaxes(tickfont = dict(size = tickfont_size), tickformat = ',d', title = None, secondary_y = True, separatethousands= True)                  
    
    #Set x axis ticks
    #Get list with number of quarters
    quarter_list      = [i for i in range(len(submarket_data_frame['Period']))]
    quarter_list      = quarter_list[0::4]

    quarter_list_text = [period for period in submarket_data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(
        tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size = tickfont_size)
                     )

    #Set size
    fig.update_layout(
    autosize  = False,
    height    = graph_height,
    width     = graph_width,
    margin    = dict(l = left_margin, r = right_margin, t = top_margin, b = bottom_margin, pad = 0,autoexpand = True),
                    )
    
    fig.write_image(os.path.join(folder,'sales_volume.png'), engine = 'kaleido', scale = scale)

def CreateAssetValueGraph(submarket_data_frame, market_data_frame, natioanl_data_frame, folder, market_title, primary_market, sector):

    #Define the MF variables and labels vs the non MF
    if sector == 'Multifamily':
        asset_value_var =  'Asset Value/Unit'
        asset_value_lab =  'Asset Value/Unit (L)'
    else:
        asset_value_var =  'Asset Value/Sqft'
        asset_value_lab =  'Asset Value/SF (L)'

    #Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    #Add Bars with Asset Values
    fig.add_trace(
    go.Bar(x            = submarket_data_frame['Period'],
           y            = submarket_data_frame[asset_value_var],
           name         = asset_value_lab,
           marker_color = bowery_grey
           ),
           secondary_y=False
                )
   
    # Add line with market cap rate for market
    fig.add_trace(
        go.Scatter(x = submarket_data_frame['Period'],
                   y = submarket_data_frame['Market Cap Rate'],
                name = market_title,
                line = dict(color = bowery_dark_blue)
                ),
        secondary_y=True
                )  

    #If it's a submarket, add primary market cap rate. If it's a primary market, add national cap rate line
    if submarket_data_frame.equals(market_data_frame):
        
        name = natioanl_data_frame['Geography Name'].iloc[0]
        if name == 'United States of America':
            name = 'National'

        if  primary_market != 'United States of America':
            
            fig.add_trace(
                go.Scatter(
                        x    = natioanl_data_frame['Period'],
                        y    = natioanl_data_frame['Market Cap Rate'],
                        name = name,
                        line = dict(color = bowery_light_blue)
                          ),
                secondary_y = True
                         )  
    
    else:
        fig.add_trace(
            go.Scatter(
                x    = market_data_frame['Period'],
                y    = market_data_frame['Market Cap Rate'],
                name = primary_market,
                line = dict(color = bowery_light_blue)
                      )
            ,secondary_y = True
                    )      
  
    #Set formatting 
    fig.update_layout(
        title_text    = "Asset Value & Market Cap Rates",    
        font_family   = font_family,
        font_color    = font_color,
        font_size     = font_size,
        height        = graph_height,
        width         = graph_width,
        margin        = dict(l = left_margin, r = right_margin, t = top_margin, b = bottom_margin),
        paper_bgcolor = backgroundcolor,
        plot_bgcolor  = backgroundcolor,
        
        title = {
            'y':       title_position,
            'x':       0.5,
            'xanchor': 'center',
            'yanchor': 'top'
                },

        legend = dict(
                    orientation = "h",
                    yanchor     = "bottom",
                    y           = legend_position,
                    xanchor     = "center",
                    x           = 0.5,
                    font_size   = tickfont_size
                    ),

                  )


    #Set x axis format
    #Get list with number of quarters
    quarter_list      = [i for i in range(len(submarket_data_frame['Period']))]
    quarter_list      = quarter_list[0::4]

    quarter_list_text = [period for period in submarket_data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(
        tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size=tickfont_size)
                    )
    
    #Set y axis format
    #Add % to right axis ticks and $ to left axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size), secondary_y = True, tickformat='.1f')  #right axis  
    fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size), secondary_y = False,                )  #left axis
    fig.update_yaxes(automargin = True)
    fig.update_xaxes(automargin = True)    
    
    #Export figure as PNG file
    fig.write_image(os.path.join(folder,'asset_values.png'), engine = 'kaleido', scale = scale)

def CreateAbsorptionGraph(submarket_data_frame, market_data_frame, natioanl_data_frame, folder, market_title, primary_market, sector):
    
    #Determine relevant variable based on sector
    if sector == 'Multifamily':
        absorption_var       = 'Absorption Units'
        inventory_growth_var = 'Inventory Units Growth'
    else:
        absorption_var       = 'Net Absorption SF'
        inventory_growth_var = 'Inventory SF Growth'
    
    #Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Add Bars with inventory growth
    fig.add_trace(
        go.Bar(
            x            = submarket_data_frame['Period'],
            y            = submarket_data_frame[inventory_growth_var],
            name         = "Inventory Growth (L)",
            marker_color = bowery_dark_grey
              ),
        secondary_y     = False
                )

    #Add Bars with net absorption rate 
    fig.add_trace(
        go.Bar(
            x             = submarket_data_frame['Period'],
            y             = submarket_data_frame[absorption_var],
            name          = "Net Absorption (L)",
            marker_color  = bowery_grey
              ),
        secondary_y = False
                 )
    
    #Vacancy Rate for (sub)market
    fig.add_trace(
        go.Scatter(
                x    = submarket_data_frame['Period'],
                y    = submarket_data_frame['Vacancy Rate'],
                name = market_title,
                line = dict(color=bowery_dark_blue)
                   ),
        secondary_y=True
                 )

    #Market
    if submarket_data_frame.equals(market_data_frame):
        
        name = natioanl_data_frame['Geography Name'].iloc[0]
        
        if name == 'United States of America':
            name = 'National'

        if  primary_market != 'United States of America':
            fig.add_trace(
                go.Scatter(
                    x = natioanl_data_frame['Period'],
                    y = natioanl_data_frame['Vacancy Rate'],
                name  = name,
                line  = dict(color = bowery_light_blue)
                        ),
                secondary_y=True
            )  
    
    #Submarket
    else:
        fig.add_trace(
            go.Scatter(
                x = market_data_frame['Period'],
                y = market_data_frame['Vacancy Rate'],
             name = primary_market,
             line = dict(color = bowery_light_blue)
                    ),
            secondary_y = True
                    )     
    
    
    #Set x-axis ticks
    quarter_list      = [i for i in range(len(submarket_data_frame['Period']))] #Get list with number of quarters
    quarter_list      = quarter_list[0::4]

    quarter_list_text = [period for period in submarket_data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(
        tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size = tickfont_size)
        )

    #Set Title
    fig.update_layout(
        title_text = "Absorption & Vacancy Rates",    
        font_family   = font_family,
        font_color    = font_color,
        font_size     = font_size,
        paper_bgcolor = backgroundcolor,
        plot_bgcolor  = backgroundcolor,

        title = {
            'y':       title_position,
            'x':       0.5,
            'xanchor': 'center',
            'yanchor': 'top'
                },
                        
        legend = dict(
                    orientation = "h",
                    yanchor     = "bottom",
                    y           = legend_position ,
                    xanchor     = "center",
                    x           = 0.5,
                    font_size   = tickfont_size
                    ),

    
            margin = dict(l = left_margin, r = right_margin, t = top_margin, b = bottom_margin),
            height = graph_height,
            width  = graph_width,
                    )



    #Add % to axis ticks
    fig.update_yaxes(title = None)
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size = tickfont_size), tickformat = '.1f', secondary_y = True) #right axis
    fig.update_yaxes(                  tickfont = dict(size = tickfont_size), tickformat = ',d',  secondary_y = False) #left axis


    
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin    = dict(r = 0))
    fig.write_image(os.path.join(folder,'absorption_rate.png'), engine = 'kaleido', scale = scale)
           
def CreateConstructionGraph(submarket_data_frame, folder, sector):
    
    #Define the MF variables and labels vs the non MF
    if sector == 'Multifamily':
        construction_var = 'Under Construction Units'
        construction_lab = "Under Construction (L)"
        title            = "Under Construction Units - Share of Inventory"
    else:
        construction_var = 'Under Construction SF'
        construction_lab = "Under Construction (L)"
        title            = 'Under Construction SF - Share of Inventory'


    #Create figure with secondary y-axis
    fig                  = make_subplots(specs=[[{"secondary_y": True}]])

    #Add Bars with suare footage or units under construction 
    fig.add_trace(
        go.Bar(
            x            = submarket_data_frame['Period'],
            y            = submarket_data_frame[construction_var],
            name         = construction_lab,
            marker_color = bowery_grey
             ),
        secondary_y = False
                ) 
    
    #Add line with share of inventory under construction
    fig.add_trace(
        go.Scatter(
            x    = submarket_data_frame['Period'],
            y    = submarket_data_frame['Under Construction %'],
            name = "Under Construction - Share of Inventory (R)",
            line = dict(color=bowery_dark_blue)
                  ),
        secondary_y = True
                )      
  
    #Set formatting 
    fig.update_layout(
        
        title_text    = title,    
        yaxis         = dict(rangemode = 'tozero'),
        font_family   = font_family,
        font_color    = font_color,
        font_size     = font_size,
        height        = graph_height,
        width         = graph_width,
        margin        = dict(l = left_margin, r = right_margin, t = top_margin, b = bottom_margin),
        paper_bgcolor = backgroundcolor,
        plot_bgcolor  = backgroundcolor,    
        
        title = {
            'y':       title_position,
            'x':       0.5,
            'xanchor': 'center',
            'yanchor': 'top'
                },

        legend = dict(
                    orientation = "h",
                    yanchor     = "bottom",
                    y           = legend_position,
                    xanchor     = "center",
                    x           = 0.5,
                    font_size   = tickfont_size
                     ),

                    
                    )

                   
    #Y-axis format(add % to right axis ticks)
    fig.update_yaxes(
        ticksuffix  = '%', 
        tickfont    = dict(size=tickfont_size), 
        tickformat  = '.1f',
        secondary_y = True
                    )
    
    #Format left axis
    fig.update_yaxes(
        tickfont          = dict(size = tickfont_size), 
        separatethousands = True,
        tickformat        = ',d',
        secondary_y       = False
                    )

    #Set x-axis ticks
    #Get list with number of quarters
    quarter_list      = [i for i in range(len(submarket_data_frame['Period']))]
    quarter_list      = quarter_list[0::4]

    quarter_list_text = [period for period in submarket_data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(
        tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size = tickfont_size)
                    )

        
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin    = dict(r = 0))

    fig.write_image(os.path.join(folder, 'construction_volume.png'), engine = 'kaleido', scale = scale)

def CreateRentGrowthGraph(submarket_data_frame, market_data_frame, natioanl_data_frame, folder, market_title, primary_market, sector):

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
                        x      = submarket_data_frame['Period'],
                        y      = submarket_data_frame[quarterly_rent_growth_var],
                        name   = quarterly_rent_growth_lab,
                        marker = dict(color = bowery_dark_blue, size = 9),
                        mode   = 'markers',
                      ),
            secondary_y = False
                 )
    
    #Default behavior for scatter plots in this package is to give some space between origin and first dot, this corrects that
    fig.update_layout(xaxis_range=[-1,len(submarket_data_frame['Period'])])
    


    #Add bars with annual rent growth
    fig.add_trace(
        go.Bar(
            x            = submarket_data_frame['Period'],
            y            = submarket_data_frame[annual_rent_growth_var],
            name         = annual_rent_growth_lab,
            marker_color = bowery_grey,
            base         = dict(layer = 'Below')        
               ),
        secondary_y = False
            )


    #Add line with vacancy rate for (sub)market
    fig.add_trace(
        go.Scatter(
                x    = submarket_data_frame['Period'],
                y    = submarket_data_frame['Vacancy Rate'],
                name = 'Vacancy Rate (L)',
                line = dict(color = bowery_black, dash = 'dash') 
                   ),
        
        secondary_y = False
                 )  

       
    #Add line with rent for submarket or market
    fig.add_trace(
        go.Scatter(
                x    = submarket_data_frame['Period'],
                y    = submarket_data_frame[rent_var],
                name = market_title,
                line = dict(color = bowery_dark_blue, dash='solid')
                  ),
        secondary_y = True) 


    #If its a submarket, add primary market rent. If it's a primary market, add national rent line
    if submarket_data_frame.equals(market_data_frame):
        extra_height = 70
        
        #Change USA --> National
        name = natioanl_data_frame['Geography Name'].iloc[0]
        if name == 'United States of America':
            name = 'National'
        
        if  primary_market != 'United States of America':
            fig.add_trace(
                go.Scatter(
                    x    = natioanl_data_frame['Period'],
                    y    = natioanl_data_frame[rent_var],
                    name = name,
                    line = dict(color = bowery_light_blue)
                          ),
                secondary_y = True
                        )  
    
    else: #Submarkets
        extra_height = 0

        fig.add_trace(
        go.Scatter(
            x    = market_data_frame['Period'],
            y    = market_data_frame[rent_var],
            name = primary_market,
            line = dict(color = bowery_light_blue)
                 ),
        secondary_y = True
                   )      
    


    #Add % to left axis ticks and $ to right axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size), tickformat='.1f',                        secondary_y=False)  #right axis  
    fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size), separatethousands= True,                 secondary_y=True)   #left axis

    #Set x axis ticks
    #Get list with number of quarters
    quarter_list      = [i for i in range(len(submarket_data_frame['Period']))]
    quarter_list      = quarter_list[0::4]

    quarter_list_text = [period for period in submarket_data_frame['Period']]
    quarter_list_text = quarter_list_text[0::4]

    fig.update_xaxes(
        tickmode = 'array',
        tickvals = quarter_list,
        ticktext = quarter_list_text,
        tickfont = dict(size = tickfont_size)
                    )
 
    #Set formatting 
    fig.update_layout(
        title_text = title,
        font_family   = font_family,
        font_color    = font_color,
        font_size     = font_size,
        height        = graph_height + extra_height , #submarket version needs to be taller to fit larger legend
        width         = graph_width ,
        margin        = dict(l = left_margin, r = right_margin, t = top_margin, b = bottom_margin),
        paper_bgcolor = backgroundcolor,
        plot_bgcolor  = backgroundcolor,    

        title={
            'y':        title_position,
            'x':        0.5,
            'xanchor': 'center',
            'yanchor': 'top'
            },
        
        #legend format
        legend = dict(
                    orientation = "h",
                    yanchor     = "bottom",
                    y           = legend_position-.05,
                    xanchor     = "center",
                    x           = 0.5,
                    font_size   = tickfont_size
                    ),
        
                    )

    
        
    fig.update_yaxes(automargin = True) 
    fig.update_xaxes(automargin = True)  
    fig.update_layout(margin = dict(r = 0))

    fig.write_image(os.path.join(folder,'rent_growth.png'), engine = 'kaleido',scale = scale)

def CreateAllGraphs(submarket_data_frame, market_data_frame, natioanl_data_frame, folder, market_title, primary_market, sector):
    CreateSalesVolumeGraph( submarket_data_frame = submarket_data_frame,                                                                                   folder = folder)
    CreateAssetValueGraph(  submarket_data_frame = submarket_data_frame, market_data_frame = market_data_frame, natioanl_data_frame = natioanl_data_frame, folder = folder, market_title = market_title, primary_market = primary_market, sector = sector)
    CreateAbsorptionGraph(  submarket_data_frame = submarket_data_frame, market_data_frame = market_data_frame, natioanl_data_frame = natioanl_data_frame, folder = folder, market_title = market_title, primary_market = primary_market, sector = sector)
    CreateConstructionGraph(submarket_data_frame = submarket_data_frame,                                                                                   folder = folder,                                                               sector = sector)
    CreateRentGrowthGraph(  submarket_data_frame = submarket_data_frame, market_data_frame = market_data_frame, natioanl_data_frame = natioanl_data_frame, folder = folder, market_title = market_title, primary_market = primary_market, sector = sector)

    #Wait till figure is saved before moving on, this avoids file path errors
    for png in ['asset_values.png', 'absorption_rate.png', 'sales_volume.png', 'construction_volume.png', 'rent_growth.png']:
        while os.path.exists(os.path.join(folder,png)) == False:
            pass


