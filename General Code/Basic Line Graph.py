#Author: Mike Leahy
#Date: 9/22/2021
#Summary: This is a utility script to quickly create simple graphs with 1 to 2 variables

import pandas as pd
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots


#Define file path of data file
desktop_location = os.path.join(os.environ['USERPROFILE'],'Desktop')
file_location    = os.path.join(desktop_location,'CPIAUCSL.csv') 

#Import data folder
df  = pd.read_csv(file_location)
df["Monthly % Change"] = ((df['CPIAUCSL']/df['CPIAUCSL'].shift(1)) - 1 ) * 100




#Set graph parameters
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
#Parameters now set



#Create Graph
fig = make_subplots(specs=[[{"secondary_y": True}]])

# Add line 
fig.add_trace(
go.Scatter(x=df['DATE'],
        y=df['CPIAUCSL'],
        name='Consumer Price Index (R)',
        line=dict(color="#4160D3", width=3),
        mode = 'lines'
        )
,secondary_y=True)  

#Add Bars
fig.add_trace(
go.Bar(x=df['DATE'],
        y=df["Monthly % Change"],
        name="Monthly % Change (L)",
        marker_color="#D7DEEA")
        ,secondary_y=False
        )

#Update y-axis
fig.update_yaxes(tickfont = dict(size=tickfont_size), ticksuffix = '%',  title = None, secondary_y=False)                


#Set formatting 
fig.update_layout(
title_text="Cosumer Price Index and Monthly % Change",    
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

fig.write_image(os.path.join(desktop_location,'cpi_graph.png'),engine='kaleido',scale=3)

 

