from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui
import time
import os

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
browser = webdriver.Chrome(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop','chromedriver.exe')),options=options)
browser.get('https:google.com/maps')
Place = browser.find_element_by_class_name("tactile-searchbox-input")
Place.send_keys(('Butler County, PA'))
Submit = browser.find_element_by_xpath(
"/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[1]/div[1]/div[2]/div[1]/button")
Submit.click()
time.sleep(10)
print(pyautogui.position())
im2 = pyautogui.screenshot(region=(1372,516, 2200, 1600) ) #left, top, width, and height
time.sleep(.25)
im2.save(os.path.join(os.environ['USERPROFILE'],'Desktop','map.png'))
im2.close()
browser.quit()




def GetCountyGDPIndustryBreakdown(fips,year):



    #Returns data on each industries contribution (in percentage points) to annual county GDP growth 
    df = pd.DataFrame({'Code':[], 'GeoFips':[],'GeoName':[],'TimePeriod':[],'CL_UNIT':[],'Percentage points':[],'UNIT_MULT':[],'DataValue':[]})
    
    for year in range(int(year),int(year)-10,-1):
        for line_code in ['91','92']: #91 is private goods, 92 is private services
            bea_url = ('https://apps.bea.gov/api/data/?userid=' + 
                bea_api_key    + 
                '&method=GetData&datasetname=Regional&'   + 
                'year='                                   +
                str(year)                                     + 
                '&TableName=CAGDP11&'                     + 
                'GeoFips='                                +
                fips                                      + 
                '&LineCode='                              +
                line_code)
            data = requests.get(bea_url)
            data = data.json()
            data = data['BEAAPI']['Results']['Data']
            df = df.append(data,ignore_index=True)
            time.sleep(.25)

    #Drop the industries where data is supressed
    df = df.loc[df['DataValue'] != '(D)']
    df['DataValue'] = df['DataValue'].astype(float)

    df['Description'] = ''
    df.loc[df['Code']=='CAGDP11-92', 'Description'] = 'Private services-providing industries'
    df.loc[df['Code']=='CAGDP11-91', 'Description'] = 'Private goods-producing industries'

    df = df.sort_values(by=['Code','TimePeriod'])
    df.to_excel(os.path.join(county_folder,'County Industry GDP Growth Breakdown.xlsx'))
    return(df)

def GetStateCensusPop10(fips):
    state_fips  = fips[0:2]
    county_fips = fips[2:]
    county_2010_pop = c.sf1.state_county(fields = ['P012001'], state_fips = state_fips, county_fips = county_fips)
    county_2010_pop = county_2010_pop[0]['P012001']
    return(county_2010_pop)    

def GetMSACensusPop10(fips):
    state_fips  = fips[0:2]
    county_fips = fips[2:]
    county_2010_pop = c.sf1.state_county(fields = ['P012001'], state_fips = state_fips, county_fips = county_fips)
    county_2010_pop = county_2010_pop[0]['P012001']
    return(county_2010_pop)


def GetCountyCensusPop10(fips):
    state_fips  = fips[0:2]
    county_fips = fips[2:]
    county_2010_pop = c.sf1.state_county(fields = ['P012001'], state_fips = state_fips, county_fips = county_fips)
    county_2010_pop = county_2010_pop[0]['P012001']
    return(county_2010_pop)





def GetStatePCI(state):
    #Per Capita Personal Income
    state_pci_series_code = state + 'PCPI' 
    state_pci_df = fred.get_series(series_id = state_pci_series_code)
    state_pci_df = state_pci_df.to_frame().reset_index()
    state_pci_df.columns = ['Period','Per Capita Personal Income']
    # state_pci_df.to_csv(os.path.join(county_folder,'State Per Capita Personal Income.csv'))
    return(state_pci_df)

def CreateIndustryGDPGraph(county_data_frame,folder):

    fig = make_subplots(specs=[[{"secondary_y": False}]])
    county_data_frame_goods    =  county_data_frame.loc[county_data_frame['Code'] =='CAGDP11-91']
    county_data_frame_services =  county_data_frame.loc[county_data_frame['Code'] =='CAGDP11-92']
    fig.add_trace(
        go.Bar( y=county_data_frame_goods['DataValue'],
                x=county_data_frame_goods['TimePeriod'],
                marker_color="#4160D3",
                name = 'Private goods-producing industries'
                )
    )

    fig.add_trace(
        go.Bar( y=county_data_frame_services['DataValue'],
                x=county_data_frame_services['TimePeriod'],
                marker_color="#B3C3FF",
                name = 'Private services-providing industries'
                )
    )

    #Set formatting 
    fig.update_layout(
    title_text="Industry Contribution to GDP Growth (Percentage Points)",    
    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
    
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
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

    #Add % to left axis ticks
    # fig.update_yaxes(tickfont = dict(size=tickfont_size), ticksuffix = '%',  title = None ,   secondary_y=False)                 #left axis
    

    fig.update_xaxes(tickmode = 'array',
        # tickangle = 45,
        tickfont = dict(size=tickfont_size)
        )

    #Set size
    fig.update_layout(
    autosize=False,
    height    = graph_height,
    width     = graph_width,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin,pad=0,autoexpand = True),
    )


    fig.write_image(os.path.join(folder,'gdp_growth_by_industry.png'),engine='kaleido',scale=scale)

def GetStateUnemployment(fips,start_year,end_year): 
    #Total Unemployment
    series_name = 'LASST' + fips[0:2] + '0000000000004'
    state_ue_df = bls.series(series_name,start_year=start_year,end_year=end_year) 

    state_ue_df['year']   =    state_ue_df['year'].astype(str)
    state_ue_df['period'] =    state_ue_df['year'] + ' '  + state_ue_df['period']              
    # state_ue_df.to_csv(os.path.join(county_folder,'State Total Unemployment.csv'))
    return(state_ue_df)

def GetMSAUnemployment(cbsa,start_year,end_year): 
    #Total Unemployment
    series_name = 'LAUMT' + cbsa_main_state_fips +  cbsa + '00000004'
    msa_ue_df = bls.series(series_name,start_year=start_year, end_year=end_year) 

    msa_ue_df['year']   =    msa_ue_df['year'].astype(str)
    msa_ue_df['period'] =    msa_ue_df['year'] + ' '  + msa_ue_df['period']              
    # msa_ue_df.to_csv(os.path.join(county_folder,'MSA Total Unemployment.csv'))
    return(msa_ue_df)



def GetCountyUnemployment(fips,start_year,end_year): 
    #Total Unemployment
    series_name = 'LAUCN' + fips + '0000000004'
    county_ue_df = bls.series(series_name,start_year=start_year,end_year=end_year) 

    county_ue_df['year'] = county_ue_df['year'].astype(str)
    county_ue_df['period'] =    county_ue_df['year'] + ' '  + county_ue_df['period']              
    # county_ue_df.to_csv(os.path.join(county_folder,'County Total Unemployment.csv'))
    return(county_ue_df)
  
def GetCountyHPI(fips,observation_start):
    #Returns County All Transaction Home Price Index
    county_hpi_series_code = 'ATNHPIUS' + fips + 'A'
    county_hpi_df = fred.get_series(series_id = county_hpi_series_code,observation_start = observation_start)
    county_hpi_df = county_hpi_df.to_frame().reset_index()
    county_hpi_df.columns = ['Period','Home Price Index']
    # county_hpi_df.to_csv(os.path.join(county_folder,'County HPI.csv'))
    return(county_hpi_df)

def CreateEmployersByIndustryGraph(county_data_frame,folder):
    #Employers By Supersector Treemap
    fig = go.Figure(go.Treemap(
    values = county_data_frame['qtrly_estabs'],
    labels = county_data_frame['industry_code'],
    parents = county_data_frame['county'],
    marker=dict(
      colors=county_data_frame['avg_wkly_wage'],
      colorscale='Blues'),
                              )
                   )
    #Set font
    # fig.update_layout(uniformtext=dict(minsize=6,mode='hide'))

    #Set Title
    fig.update_layout(
    title={
        'text':'Private Establishments by Industry',
        'y':title_position - .10,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )

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

    fig.write_image(os.path.join(folder,'establishments_by_industry.png'),engine='kaleido',scale=scale)

def CreatePopulationOverTimeGraph(county_data_frame,msa_data_frame,state_data_frame,folder):
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #County Population
    fig.add_trace(
    go.Scatter(x=county_data_frame['Period'],
            y=county_data_frame['Resident Population'],
            name=county + ' (L)',
            line=dict(color="#4160D3")
                                    )      
    ,secondary_y=False)
    
    #MSA Population if applicable
    if cbsa != '':
        fig.add_trace(
        go.Scatter(x=msa_data_frame['Period'],
                y=msa_data_frame['Resident Population'],
                name=cbsa_name + ' (R)',
                line=dict(color ="#B3C3FF")
                )
        ,secondary_y=True)
    else:
        #State Population
        fig.add_trace(
        go.Scatter(x=state_data_frame['Period'],
                y=state_data_frame['Resident Population'],
                name=state_name + ' (R)',
                line = dict(color="#A6B0BF")
                )
        ,secondary_y=True)   


    #Set X-Axis Format
    fig.update_xaxes(
        type = 'date',
        dtick="M12",
        tickformat="%Y",
        tickangle = 0,
        tickfont = dict(size=tickfont_size),
        linecolor = 'black'
        )

    #Set Y-Axis format
    fig.update_yaxes( tickfont = dict(size=tickfont_size),
                      linecolor='black'  
                      )

    #Set Title
    fig.update_layout(
    title_text="Population",    

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

    

    fig.write_image(os.path.join(folder,'resident_population.png'),engine='kaleido',scale=scale)

def CreatePopulationGrowthGraph(folder):
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Calculate annualized growth rates for the county, msa (if available), and state dataframes
    county_resident_pop['Resident Population_1year_growth'] =  (((county_resident_pop['Resident Population']/county_resident_pop['Resident Population'].shift(1))  - 1) * 100)/1
    county_resident_pop['Resident Population_5year_growth'] =  (((county_resident_pop['Resident Population']/county_resident_pop['Resident Population'].shift(5))   - 1) * 100)/5
    county_resident_pop['Resident Population_10year_growth'] =  (((county_resident_pop['Resident Population']/county_resident_pop['Resident Population'].shift(10)) - 1) * 100)/10

    county_1y_growth  = county_resident_pop.iloc[-1]['Resident Population_1year_growth'] 
    county_5y_growth  = county_resident_pop.iloc[-1]['Resident Population_5year_growth'] 
    county_10y_growth = county_resident_pop.iloc[-1]['Resident Population_10year_growth']

    if cbsa != '':
        #Make sure we are comparing same years for calculating growth rates for county and msa
        msa_resident_pop = msa_resident_pop.loc[msa_resident_pop['Period'] <= (county_resident_pop['Period'].max())]
        msa_resident_pop['Resident Population_1year_growth'] =  (((msa_resident_pop['Resident Population']/msa_resident_pop['Resident Population'].shift(1))  - 1) * 100)/1
        msa_resident_pop['Resident Population_5year_growth'] =  (((msa_resident_pop['Resident Population']/msa_resident_pop['Resident Population'].shift(5))   - 1) * 100)/5
        msa_resident_pop['Resident Population_10year_growth'] =  (((msa_resident_pop['Resident Population']/msa_resident_pop['Resident Population'].shift(10)) - 1) * 100)/10

        msa_1y_growth  = msa_resident_pop.iloc[-1]['Resident Population_1year_growth'] 
        msa_5y_growth  = msa_resident_pop.iloc[-1]['Resident Population_5year_growth'] 
        msa_10y_growth = msa_resident_pop.iloc[-1]['Resident Population_10year_growth']

    #Make sure we are comparing same years for calculating growth rates for county and state
    state_resident_pop = state_resident_pop.loc[state_resident_pop['Period'] <= (county_resident_pop['Period'].max())]
    state_resident_pop['Resident Population_1year_growth'] =  (((state_resident_pop['Resident Population']/state_resident_pop['Resident Population'].shift(1))  - 1) * 100)/1
    state_resident_pop['Resident Population_5year_growth'] =  (((state_resident_pop['Resident Population']/state_resident_pop['Resident Population'].shift(5))   - 1) * 100)/5
    state_resident_pop['Resident Population_10year_growth'] =  (((state_resident_pop['Resident Population']/state_resident_pop['Resident Population'].shift(10)) - 1) * 100)/10

    state_1y_growth  = state_resident_pop.iloc[-1]['Resident Population_1year_growth'] 
    state_5y_growth  = state_resident_pop.iloc[-1]['Resident Population_5year_growth'] 
    state_10y_growth = state_resident_pop.iloc[-1]['Resident Population_10year_growth']

    #Now that we've calculated growth rates, create our plot
    years=['10 Years', '5 Years', '1 Year']
    annotation_position = 'outside'
    #If there's a MSA/CBSA include it, otherwise just use county and state
    if cbsa != '':
        fig = go.Figure(data=[

       
        
        go.Bar(
            name=cbsa_name,   
            x=years, 
            y=[msa_10y_growth, msa_5y_growth, msa_1y_growth],
            marker_color ="#B3C3FF",
            text = [msa_10y_growth, msa_5y_growth, msa_1y_growth],
            texttemplate = "%{value:.2f}%",
            textposition = annotation_position ,
            cliponaxis =  False
            ),
        
        go.Bar(
            name=county,      
            x=years, 
            y=[county_10y_growth,county_5y_growth,county_1y_growth],
            marker_color="#4160D3",
            text = [county_10y_growth,county_5y_growth,county_1y_growth],
            texttemplate = "%{value:.2f}%",
            textposition = annotation_position,
            cliponaxis =  False
                )
        ])
    
    else:
        fig = go.Figure(data=[

        go.Bar(
            name=state_name,  
            x=years, 
            y=[state_10y_growth, state_5y_growth, state_1y_growth],
            marker_color ="#A6B0BF",
            text = [state_10y_growth, state_5y_growth, state_1y_growth],
            texttemplate = "%{value:.2f}%",
            textposition = annotation_position,
            cliponaxis =  False
            ),
        
        go.Bar(
                name=county,      
                x=years, 
                y=[county_10y_growth,county_5y_growth,county_1y_growth],
                marker_color="#4160D3",
                text = [county_10y_growth,county_5y_growth,county_1y_growth],
                texttemplate = "%{value:.2f}%",
                textposition = annotation_position,
                cliponaxis =  False
                )
        ]
        )

    #Change the bar mode
    fig.update_layout(barmode='group')
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')

    #Set X-axes format
    fig.update_xaxes(
        tickfont = dict(size=tickfont_size)
        )

    #Set Y-Axes format
    fig.update_yaxes(
        ticksuffix = '%',
        tickfont = dict(size=tickfont_size),
        visible = False)                 

    #Set Title
    fig.update_layout(
    title_text="Annualized Population Growth",    

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
    margin=dict(l=left_margin, r=right_margin, t=(top_margin + .2), b = (bottom_margin + .2)),
    height    = graph_height,
    width     = graph_width,
                    )
    fig.write_image(os.path.join(folder,'population_growth.png'),engine='kaleido',scale=scale)

def CreateIncomeGrowthGraph(county_data_frame, msa_data_frame, state_data_frame,folder):
    #Growth in per capita personal income
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    #Calculate annualized growth rates for the county, msa (if available), and state dataframes
    county_data_frame['Per Capita Personal Income_1year_growth'] =  (((county_data_frame['Per Capita Personal Income']/county_data_frame['Per Capita Personal Income'].shift(1))  - 1) * 100)/1
    county_data_frame['Per Capita Personal Income_3year_growth'] =  (((county_data_frame['Per Capita Personal Income']/county_data_frame['Per Capita Personal Income'].shift(3))   - 1) * 100)/3
    county_data_frame['Per Capita Personal Income_5year_growth'] =  (((county_data_frame['Per Capita Personal Income']/county_data_frame['Per Capita Personal Income'].shift(5))   - 1) * 100)/5

    county_1y_growth  = county_data_frame.iloc[-1]['Per Capita Personal Income_1year_growth'] 
    county_3y_growth  = county_data_frame.iloc[-1]['Per Capita Personal Income_3year_growth'] 
    county_5y_growth  = county_data_frame.iloc[-1]['Per Capita Personal Income_5year_growth'] 
    
    if cbsa != '':
        #Make sure we are comparing same years for calculating growth rates for county and msa
        msa_data_frame = msa_data_frame.loc[msa_data_frame['Period'] <= (county_data_frame['Period'].max())]
        msa_data_frame['Per Capita Personal Income_1year_growth'] =  (((msa_data_frame['Per Capita Personal Income']/msa_data_frame['Per Capita Personal Income'].shift(1))  - 1) * 100)/1
        msa_data_frame['Per Capita Personal Income_3year_growth'] =  (((msa_data_frame['Per Capita Personal Income']/msa_data_frame['Per Capita Personal Income'].shift(3))   - 1) * 100)/3
        msa_data_frame['Per Capita Personal Income_5year_growth'] =  (((msa_data_frame['Per Capita Personal Income']/msa_data_frame['Per Capita Personal Income'].shift(5))   - 1) * 100)/5

        msa_1y_growth  = msa_data_frame.iloc[-1]['Per Capita Personal Income_1year_growth'] 
        msa_3y_growth  = msa_data_frame.iloc[-1]['Per Capita Personal Income_3year_growth'] 
        msa_5y_growth  = msa_data_frame.iloc[-1]['Per Capita Personal Income_5year_growth'] 

    #Make sure we are comparing same years for calculating growth rates for county and state
    state_data_frame = state_data_frame.loc[state_data_frame['Period'] <= (county_data_frame['Period'].max())]
    state_data_frame['Per Capita Personal Income_1year_growth'] =  (((state_data_frame['Per Capita Personal Income']/state_data_frame['Per Capita Personal Income'].shift(1))  - 1) * 100)/1
    state_data_frame['Per Capita Personal Income_3year_growth'] =  (((state_data_frame['Per Capita Personal Income']/state_data_frame['Per Capita Personal Income'].shift(3))   - 1) * 100)/3
    state_data_frame['Per Capita Personal Income_5year_growth'] =  (((state_data_frame['Per Capita Personal Income']/state_data_frame['Per Capita Personal Income'].shift(5))   - 1) * 100)/5

    state_1y_growth  = state_data_frame.iloc[-1]['Per Capita Personal Income_1year_growth'] 
    state_3y_growth  = state_data_frame.iloc[-1]['Per Capita Personal Income_3year_growth'] 
    state_5y_growth  = state_data_frame.iloc[-1]['Per Capita Personal Income_5year_growth'] 

    #Now that we've calculated growth rates, create our plot
    years=['5 Years', '3 Years','1 Year']
    annotation_position = 'outside'
    
    if cbsa != '':
        fig = go.Figure(data=[
    
        go.Bar(
            name = cbsa_name,  
            x=years, 
            y=[msa_5y_growth, msa_3y_growth, msa_1y_growth],
            marker_color ="#B3C3FF",
            text = [msa_5y_growth, msa_3y_growth, msa_1y_growth],
            texttemplate = "%{value:.2f}%",
            textposition = annotation_position,
            cliponaxis =  False
            ),

        go.Bar(
            name=county,      
            x=years, 
            y=[county_5y_growth,county_3y_growth,county_1y_growth],
            marker_color="#4160D3",
            text = [county_5y_growth,county_3y_growth,county_1y_growth],
            texttemplate = "%{value:.2f}%",
            textposition = annotation_position,
            cliponaxis =  False
        )
        ])

    else:
        fig = go.Figure(data=[

        go.Bar(
                name=state_name,  
                x=years, 
                y=[state_5y_growth, state_3y_growth, state_1y_growth],
                marker_color ="#A6B0BF",
                text = [state_5y_growth, state_3y_growth, state_1y_growth],
                texttemplate = "%{value:.2f}%",
                textposition = annotation_position,
                cliponaxis =  False
                ),

        go.Bar(
                name=county,      
                x=years, 
                y=[county_5y_growth,county_3y_growth,county_1y_growth],
                marker_color="#4160D3",
                text = [county_5y_growth,county_3y_growth,county_1y_growth],
                texttemplate = "%{value:.2f}%",
                textposition = annotation_position,
                cliponaxis =  False
                ),


        

        ])


    #Change the bar mode
    fig.update_layout(barmode='group')

    #Set X-axes format
    fig.update_xaxes(
        tickfont = dict(size=tickfont_size)
        )

    #Set Y-Axes format
    fig.update_yaxes(
        ticksuffix = '%',
        tickfont = dict(size=tickfont_size),
        visible = False)                 

    #Set Title
    fig.update_layout(
    title_text="Annualized Per Capita Personal Income Growth",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
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
    
    fig.write_image(os.path.join(folder,'pci_growth.png'),engine='kaleido',scale=scale)

def CreatePCIGraph(county_data_frame,msa_data_frame,state_data_frame,folder):

    fig = make_subplots(specs=[[{"secondary_y": False}]])

    #Add county PCI
    fig.add_trace(
    go.Scatter(x=county_data_frame['Period'],
            y=county_data_frame['Per Capita Personal Income'],
            name=county,
            line = dict(color="#4160D3")
            )
    ,secondary_y=False)

   #Add MSA PCI if applicable
    if cbsa != '':
        fig.add_trace(
        go.Scatter(x=msa_data_frame['Period'],
                y=msa_data_frame['Per Capita Personal Income'],
                name=cbsa_name,
                line = dict(color="#B3C3FF"),
                )
        ,secondary_y=False)
    else:
        #Add state PCI
        fig.add_trace(
        go.Scatter(x=state_data_frame['Period'],
                y=state_data_frame['Per Capita Personal Income'],
                name=state_name,
                line=dict(color='#A6B0BF'),
                )
        ,secondary_y=False)

    #Set X-Axis Format
    fig.update_xaxes(
        type = 'date',
        dtick="M12",
        tickformat="%Y",
        tickangle = 0,
        tickfont = dict(size=tickfont_size),
        linecolor = 'black'
        )

    #Set Y-Axis format
    fig.update_yaxes( tickfont = dict(size=tickfont_size),
                      linecolor='black'  
                    )

    fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size),secondary_y=False)
    # fig.update_yaxes(tickprefix = '$', tickfont = dict(size=tickfont_size),secondary_y=True)


    #Set Title
    fig.update_layout(
    title_text="Per Capita Personal Income",    

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

    # fig.update_yaxes(automargin = True)  
    fig.write_image(os.path.join(folder,'per_capita_income.png'),engine='kaleido',scale=scale)


#CreateEmployersByIndustryGraph(county_data_frame = county_industry_breakdown, folder = county_folder )



# CreateHPIGraph(county_data_frame = county_hpi ,msa_data_frame = '',state_data_frame='', folder = county_folder )


# msa_unemployment                = GetMSAUnemployment(cbsa = cbsa,start_year=start_year,end_year=end_year)

    
# CreateIndustryGDPGraph(county_data_frame = county_gdp_industry_breakdown, folder = county_folder )

# state_unemployment             = GetStateUnemployment(fips = fips,start_year=start_year,end_year=end_year)


# county_unemployment           = GetCountyUnemployment(fips = fips,start_year=start_year,end_year=end_year)

# county_hpi                      = GetCountyHPI(fips = fips,observation_start = '01/01/2000')  


# CreateMedianHHIncomeGraph(county_data_frame = county_mhhi , state_data_frame = state_mhhi, folder = county_folder )

def CreateMedianHHIncomeGraph(county_data_frame,state_data_frame,folder):

    fig = make_subplots(specs=[[{"secondary_y": False}]])

    #County household income
    fig.add_trace(
    go.Scatter(x=county_data_frame['Period'],
            y=county_data_frame['Median Household Income'],
            name=county,
            line=dict(color="#4160D3"))
    ,secondary_y=False)

    #State median household income
    fig.add_trace(
    go.Scatter(x=state_data_frame['Period'],
            y=state_data_frame['Median Household Income'],
            name=state_name,
            line=dict(color="#A6B0BF"))
    ,secondary_y=False)

    #Set formatting 
    fig.update_layout(
    title_text="Median Household Income",    
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
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"    
                    )

    #Set X-Axis Format
    fig.update_xaxes(
        type = 'date',
        dtick="M12",
        tickformat="%Y",
        tickangle = 0,
        tickfont = dict(size=tickfont_size),
        linecolor = 'black'
        )

    #Set Y-Axis format
    fig.update_yaxes( tickfont = dict(size=tickfont_size),
                      linecolor='black'  
                      )

    #Add % to left axis ticks
    fig.update_yaxes(tickfont = dict(size=tickfont_size), tickprefix = '$',  title = None ,   secondary_y=False)                 #left axis
    

    #Set size
    fig.update_layout(
    autosize=False,
    height    = graph_height,
    width     = graph_width,
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin,pad=0,autoexpand = True),)
    


    fig.write_image(os.path.join(folder,'median_hh_income.png'),engine='kaleido',scale=scale)

def GetStateMHHI(state,observation_start):
    #Median Household Income
    state_mhhi_series_code = 'MEHOINUS' + state + 'A646N'
    state_mhhi_df = fred.get_series(series_id = state_mhhi_series_code,observation_start=observation_start)
    state_mhhi_df = state_mhhi_df.to_frame().reset_index()
    state_mhhi_df.columns = ['Period','Median Household Income']
    # state_mhhi_df.to_csv(os.path.join(county_folder,'State Median Household Income.csv'))
    return(state_mhhi_df)

def GetCountyMHHI(fips,state,observation_start):
    #Median Household Income
    county_mhhi_series_code = 'MHI' + state + fips + 'A052NCEN'
    county_mhhi_df = fred.get_series(series_id = county_mhhi_series_code,observation_start=observation_start)
    county_mhhi_df = county_mhhi_df.to_frame().reset_index()
    county_mhhi_df.columns = ['Period','Median Household Income']
    # county_mhhi_df.to_csv(os.path.join(county_folder,'County Median Household Income.csv'))
    return(county_mhhi_df)
    
# county_mhhi                   = GetCountyMHHI(fips=fips,state=state,observation_start=observation_start)

# state_mhhi                       = GetStateMHHI(state=state, observation_start = observation_start)


# #Track 5 Year MSA Employment Growth
    # if isinstance(msa_employment, pd.DataFrame) == True:
    #     msa_employment_extra_month_cut_off  = msa_employment.loc[msa_employment['period'] <= (county_employment['period'].max())] #msa employemt data sometimes is released before counties 
    #                                                                                                                               # so we need to make sure we compare apples to apples for 
    #                                                                                                                               # county vs state 5 year employment growth
    #     latest_msa_employment               = msa_employment_extra_month_cut_off['Employment'].iloc[-1]
    #     five_years_ago_msa_employment       = msa_employment_extra_month_cut_off['Employment'].iloc[-61] 
    #     five_year_msa_employment_growth_pct = ((latest_msa_employment/five_years_ago_msa_employment) - 1 ) * 100
    #     five_year_msa_employment_growth_pct = "{:,.1f}%".format(five_year_msa_employment_growth_pct)
    #     state_or_metro                      =  cbsa_name
    #     five_year_state_or_metro_growth     = five_year_msa_employment_growth_pct

    # #Track 5 Year State Employment Growth
    # state_employment_extra_month_cut_off  = state_employment.loc[state_employment['period'] <= (county_employment['period'].max())] #state employemt data sometimes is released before counties 
    #                                                                                        # so we need to make sure we compare apples to apples for 
    #                                                                                        # county vs state 5 year employment growth

    # latest_state_employment               = state_employment_extra_month_cut_off['Employment'].iloc[-1]
    # five_years_ago_state_employment       = state_employment_extra_month_cut_off['Employment'].iloc[-61] 
    # five_year_state_employment_growth_pct = ((latest_state_employment/five_years_ago_state_employment) - 1 ) * 100
    # five_year_state_employment_growth_pct = "{:,.1f}%".format(five_year_state_employment_growth_pct)
    # state_or_metro                        = state_name
    # five_year_state_or_metro_growth       = five_year_state_employment_growth_pct

    

##################################
    # if cbsa != '':
 
    #     else:
    #         fig.add_trace( go.Bar(
    #             name = 'United States',  
    #             x=years, 
    #             y=[national_5y_growth, national_3y_growth, national_1y_growth],
    #             marker_color ="#000F44",
    #             text = [national_5y_growth, national_3y_growth, national_1y_growth],
    #             texttemplate = "%{value:.2f}%",
    #             textposition = annotation_position,
    #             cliponaxis =  False
    #             ),
    #             row = 1,
    #             col = 2
    #             )

    #     fig.add_trace( go.Bar(
    #         name = cbsa_name + ' (MSA)',  
    #         x=years, 
    #         y=[msa_5y_growth, msa_3y_growth, msa_1y_growth],
    #         marker_color ="#B3C3FF",
    #         text = [msa_5y_growth, msa_3y_growth, msa_1y_growth],
    #         texttemplate = "%{value:.2f}%",
    #         textposition = annotation_position,
    #         cliponaxis =  False
    #         ),
    #         row = 1,
    #         col = 2
    #    )
    #    if county_data_frame != '':
    #     fig.add_trace(go.Bar(
    #             name=county,      
    #             x=years, 
    #             y=[county_5y_growth,county_3y_growth,county_1y_growth],
    #             marker_color="#4160D3",
    #             text = [county_5y_growth,county_3y_growth,county_1y_growth],
    #             texttemplate = "%{value:.2f}%",
    #             textposition = annotation_position,
    #             cliponaxis =  False
    #         ),
    #         row = 1,
    #         col = 2
    #     )
        

    # else:

    #     fig.add_trace(go.Bar(
    #         name = 'United States',  
    #         x=years, 
    #         y=[national_5y_growth, national_3y_growth, national_1y_growth],
    #         marker_color ="#000F44",
    #         text = [national_5y_growth, national_3y_growth, national_1y_growth],
    #         texttemplate = "%{value:.2f}%",
    #         textposition = annotation_position,
    #         cliponaxis =  False
    #         ),
    #         row = 1,
    #         col = 2
    #    )
       

    #     fig.add_trace(go.Bar(
    #             name=state_name,  
    #             x=years, 
    #             y=[state_5y_growth, state_3y_growth, state_1y_growth],
    #             marker_color ="#A6B0BF",
    #             text = [state_5y_growth, state_3y_growth, state_1y_growth],
    #             texttemplate = "%{value:.2f}%",
    #             textposition = annotation_position,
    #             cliponaxis =  False
    #             ),
    #             row = 1,
    #             col = 2
    #     )
    #     if county_data_frame != '':
    #         fig.add_trace(go.Bar(
    #                 name=county,      
    #                 x=years, 
    #                 y=[county_5y_growth,county_3y_growth,county_1y_growth],
    #                 marker_color="#4160D3",
    #                 text = [county_5y_growth,county_3y_growth,county_1y_growth],
    #                 texttemplate = "%{value:.2f}%",
    #                 textposition = annotation_position,
    #                 cliponaxis =  False
    #                 ),
    #                 row = 1,
    #                 col = 2
    #         )


def CarLanguage():
    print('Writing Car Langauge')
    
    try:
        page                          =  wikipedia.page((county + ',' + state))
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
            return(county + ' is not connected by any major highways or roads.')
        else:
            return(car_language)
    except:
        return('')
 
def PlaneLanguage():
    print('Writing Plane Langauge')

    try:
        #Go though some common section names for airports
        page                  = wikipedia.page((county + ',' + state))
        airports              = page.section('Airports')
        air                   = page.section('Air')
        aviation              = page.section('Aviation')
        air_transport         = page.section('Air Transport')

        plane_language = ''
        for count,section in enumerate([airports,air,aviation,air_transport]):
            if (section != None) and (count == 0):
                plane_language =  section 
            elif (section != None) and (count > 0):
                plane_language = plane_language + ' ' + "\n" + section 

        

        #If the wikipedia page is missiing all airport sections 
        if plane_language == '':
            return(county + ' is not served by any airport.')
        else:
            return(plane_language)
    except:
        return('')        
    
def BusLanguage():
    print('Writing Bus Langauge')

    try:
        page                         =  wikipedia.page((county + ',' + state))
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
            return(county + ' does not have public bus service.')
        else:
            return(bus_language)
    except:
        return('')

def TrainLanguage():
    print('Writing Train Langauge')
    try:
        page                         =  wikipedia.page((county + ',' + state))
        rail                         =  page.section('Rail')
        public_transportation        =  page.section('Public transportation')
        public_Transportation        =  page.section('Public Transportation')
        public_transport             =  page.section('Public transport')
        public_Transit               =  page.section('Public Transit')
        mass_transit                 =  page.section('Mass transit')
        rail_network                 =  page.section('Rail Network')
        intercity_rail               =  page.section('Intercity Rail')

        #Add the text from the sections above to a single string variable
        train_language = ''
        for count,section in enumerate([rail,public_transportation,public_Transportation,public_transport,public_Transit,mass_transit,rail_network,intercity_rail]):
            if (section != None) and (count == 0):
                train_language =  section 
            elif (section != None) and (count > 0):
                train_language = train_language + ' ' + "\n" + section 

        
        #If the wikipedia page is missiing all airport sections return default phrase
        if train_language == '':
            return(county + ' is not served by any commuter or light rail lines.')
        else:
            return(train_language)
    except:
        return('')
  
