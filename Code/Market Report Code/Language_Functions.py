
import numpy as np

#Langauge for overview section
def CreateOverviewLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    #Create variables we will use in the language

    if sector == 'Multifamily':
        yoy_rent_growth = data_frame['YoY Market Effective Rent/Unit Growth'].iloc[-1]
        qoq_rent_growth = data_frame['QoQ Market Effective Rent/Unit Growth'].iloc[-1]
        unit_or_sqft                    = 'unit'
        unit_or_sqft_singular           = 'unit'
        extra_s                         = 's'
        under_construction              = data_frame['Under Construction Units'].iloc[-1]
        under_construction_share        = round(data_frame['Under Construction %'].iloc[-1],2)
        
        #Get current asset value
        asset_value          = data_frame['Asset Value/Unit'].iloc[-1]
        asset_value          = "${:,.0f}".format(asset_value)
        asset_value_change   = data_frame['YoY Asset Value/Unit Growth'].iloc[-1]
        if asset_value_change > 0:
            asset_value_change_description = 'expanded'
        elif asset_value_change < 0:
            asset_value_change_description = 'compressed'
        else:
            asset_value_change_description = 'remained constant'

        #Get Submarket and market inventory and the fraction of the inventory the submarket makes up
        submarket_inventory = data_frame['Inventory Units'].iloc[-1]
        market_inventory    = data_frame2['Inventory Units'].iloc[-1]
        submarket_inventory_fraction = (submarket_inventory/market_inventory) * 100



    else: #non multifamily
        yoy_rent_growth = data_frame['YoY Rent Growth'].iloc[-1]
        yoy_rent_growth = yoy_rent_growth
        qoq_rent_growth = data_frame['QoQ Rent Growth'].iloc[-1]
        unit_or_sqft                    = 'square feet'
        unit_or_sqft_singular           = 'SF'
        extra_s                         = ''
        under_construction              = data_frame['Under Construction SF'].iloc[-1]
        under_construction_share        = round(data_frame['Under Construction %'].iloc[-1],2)

        #Get current asset value
        asset_value          = data_frame['Asset Value/Sqft'].iloc[-1]
        asset_value          = "${:,.0f}". format(asset_value)
        asset_value_change   = data_frame['YoY Asset Value/Sqft Growth'].iloc[-1]
        if asset_value_change > 0:
            asset_value_change_description = 'expanded'
        elif asset_value_change < 0:
            asset_value_change_description = 'compressed'
        else:
            asset_value_change_description = 'remained constant'
        
        #Get Submarket and market inventory and the fraction of the inventory the submarket makes up
        submarket_inventory = data_frame['Inventory SF'].iloc[-1]
        market_inventory    = data_frame2['Inventory SF'].iloc[-1]
        submarket_inventory_fraction = (submarket_inventory/market_inventory) * 100

    #Format Variables
    under_construction              = "{:,.0f}".format(under_construction)     
    under_construction_share        = "{:,.1f}".format(under_construction_share)

    submarket_inventory             = "{:,.0f}".format(submarket_inventory) 
    market_inventory                = "{:,.0f}".format(market_inventory) 
    submarket_inventory_fraction    = "{:,.0f}%".format(submarket_inventory_fraction) 

    
    #Get langauge for rent growth
    if yoy_rent_growth > 0:
        rent_growth_description = 'expanded'
    elif yoy_rent_growth < 0:
        rent_growth_description = 'compressed'
    else:
        rent_growth_description = 'remained constant'

    yoy_rent_growth = str(abs(yoy_rent_growth))


    #Get the current quarter
    latest_quarter = data_frame['Period'].iloc[-1]

    #Get Sales info
    current_sale_volume       = data_frame['Total Sales Volume'].iloc[-1]
    try:
         current_sale_volume = round(current_sale_volume)
    except:
        pass

    current_sale_volume       = "${:,.0f}". format(current_sale_volume)
    current_transaction_count = str(round(data_frame['Sales Volume Transactions'].iloc[-1]))

    #Get current vacancy and its average over the past decade
    vacancy               = data_frame['Vacancy Rate'].iloc[-1]
    vacancy_change        = data_frame['YoY Vacancy Growth'].iloc[-1]

    if vacancy_change > 0:
        vacancy_change = 'expand'
    elif vacancy_change < 0:
        vacancy_change = 'compress'
    elif vacancy_change == 0:
        vacancy_change = 'remain constant'

    avg_vacancy    = round(data_frame['Vacancy Rate'].mean(),1)
    if vacancy > avg_vacancy:
        vacancy_avg_above_or_below = 'above'
    elif vacancy < avg_vacancy:
        vacancy_avg_above_or_below = 'below'
    elif vacancy == avg_vacancy:
        vacancy_avg_above_or_below = 'at'
    else:
        ''

    vacancy     = str(vacancy)
    avg_vacancy = str(avg_vacancy)

    #Get most recent cap rate and change in cap rate
    cap_rate               =  data_frame['Market Cap Rate'].iloc[-1] 
    avg_cap_rate           =  data_frame['Market Cap Rate'].mean() 

    if cap_rate < avg_cap_rate:
        cap_rate_above_below_average = 'below'
    elif cap_rate > avg_cap_rate:
        cap_rate_above_below_average = 'above'
    

    cap_rate_yoy_change    =  round(data_frame['YoY Market Cap Rate Growth'].iloc[-1])
    if cap_rate_yoy_change > 0:
        cap_rate_change_description = 'expanded'
    elif cap_rate_yoy_change < 0:
        cap_rate_change_description = 'compressed'
    else:
        cap_rate_change_description = 'seen minimal movement'

    cap_rate_yoy_change = str(abs(cap_rate_yoy_change))

    cap_rate             =   "{:,.1f}%".format(cap_rate)




    #Get change in demand
    demand_change   = data_frame['YoY Absorption Growth'].iloc[-1]
    if demand_change > 0:
        demand_change = 'accelerate'
    elif demand_change < 0:
        demand_change = 'slow'
    elif demand_change == 0:
        demand_change = 'remain constant'

    #Figure out change in fundamentals
    if rent_growth_description == 'expanded' and vacancy_change == 'compress': #if rent is growing and vacancy is falling we call fundamentals improving
        fundamentals_change = 'improving'

    elif rent_growth_description == 'compressed' and vacancy_change == 'expand': #if rent is falling and vacancy is rising we call fundamentals softening
        fundamentals_change = 'softening'
    else:
        fundamentals_change = 'softening/improving'




    #Market
    if data_frame.equals(data_frame2):
        market_or_submarket = 'Market'
        overview_intro_language = ('The subject property is located in the ' +
        market_title +
        ' ' +
        market_or_submarket +
        ' defined in the map above. This Market is home to ' +
        market_inventory +
        ' ' +
        unit_or_sqft + extra_s +
         ' of ' +
         sector.lower() +
         ' space. ')  

    #Submarket
    else:
        market_or_submarket = 'Submarket'
        overview_intro_language = ('The subject property is located in the ' +
        market_title +
        ' Submarket of the ' +
        primary_market +
        ' Market,' +
        ' defined in the map above. This Submarket is home to ' +
        submarket_inventory +
        ' ' +
        unit_or_sqft + extra_s +
         ' of ' +
         sector.lower() +
         ' space, ' +
         'accounting for ' +
         submarket_inventory_fraction +
         ' of the Market’s total inventory. ')  
      

    try:
        
        
        if sector == "Retail":
            retail_language =  (' The shift from brick-and-mortar stores to e-commerce has disrupted retail over the last decade and the COVID ' + 
                                'crisis appears to have accelerated that trend in the ' +
                                market_or_submarket +
                                '/, although the ' +
                                 market_or_submarket +
                                ' has emerged relatively unscathed. ')
        else:
            retail_language = '' 

        overview_language = (overview_intro_language +
            retail_language +       
            'Over the past twelve months, the ' +
                market_or_submarket +
                ' has seen demand ' +
                demand_change +
                ' causing vacancy rates to ' +
                vacancy_change             +
                ' to the current rate of ' +
                vacancy            +
                '%.'
            ' Meanwhile, rents in this ' +
                market_or_submarket +
                ' ' +
                rent_growth_description +
                ' at an annual rate of ' +
                yoy_rent_growth  +
                "% as of " +
                latest_quarter +
                '.'     +
                ' There are currently ' +
                under_construction +
                ' ' +
                unit_or_sqft +
                extra_s +
                ' underway representing an inventory expansion of ' +
                under_construction_share +
                '%.  There were ' +
                current_transaction_count +
                ' sales this quarter for a total sales volume of '+
                current_sale_volume +
                '.  With fundamentals ' +
                fundamentals_change +
                ', values have ' +
                asset_value_change_description +
                ' over the past year to the current value of ' +
                asset_value +
                '/' +
                unit_or_sqft_singular +
                ' and cap rates have ' +
                cap_rate_change_description +
                ' ' +
                cap_rate_yoy_change +
                ' bps' +
                ' to a rate of ' +
                cap_rate +
                ', falling ' +
                cap_rate_above_below_average +
                ' the long-term average.'

            
                                )


        return(overview_language)    

    #If there are are problems with the language functions, just return a simple paragraph we can edit
    except:
        return('Over the past twelve months, the ' +
                market_or_submarket +
                ' has seen demand ' +
                'accelerate/slow' +
                ' causing vacancy rates to ' +
                'expand/compress'             +
                ' to the current rate of ' +
                'X'           +
                '%.'
            ' Meanwhile, rents in this ' +
                market_or_submarket +
                ' ' +
                'expanded/compressed' +
                ' at an annual rate of ' +
                'X'  +
                "% as of " +
                latest_quarter +
                '.'     +
                ' There are currently ' +
                'X ' +
                unit_or_sqft +
                extra_s +
                ' underway representing an inventory expansion of ' +
                'X' +
                '%.  There were ' +
                'X' +
                ' sales this quarter for a total sales volume of '+
                '$X' +
                '.  With fundamentals ' +
                'improving/softening' +
                ', values have ' +
                'expanded/compressed' +
                ' over the past year to the current value of ' +
                '$X' +
                '/' +
                unit_or_sqft +
                ' and cap rates have ' +
                'compressed/expanded ' +
                'X bps' +
                ' to a rate of ' +
                'X%' +
                ', which is ' +
                'above/below ' +
                ' the long-term average.'

            
                                )


#Language for Supply and Demand Section
def CreateDemandLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    if sector == 'Multifamily':
        unit_or_sqft                    = 'units'
        net_absorption_var_name         = 'Absorption Units'
        net_absorption                  =  data_frame['Absorption Units'].iloc[-1]
        previous_quarter_net_absorption =  data_frame['Absorption Units'].iloc[-2]
    else:
        unit_or_sqft                    = 'square feet'
        net_absorption_var_name         = 'Net Absorption SF'
        net_absorption                  =  data_frame['Net Absorption SF'].iloc[-1]
        previous_quarter_net_absorption =  data_frame['Net Absorption SF'].iloc[-2]
    
    #Describe demand based on absoprtion
    if net_absorption >= 0:
        demand_description = 'attracted'
    elif net_absorption < 0:
        demand_description = 'struggled to attract'
    else:
        demand_description = 'attracted/struggled to attract'

    #Describe quarter over quarter change
    if net_absorption > previous_quarter_net_absorption:
        qoq_absorption_increase_or_decrease = 'increase'
    elif net_absorption < previous_quarter_net_absorption:
        qoq_absorption_increase_or_decrease = 'decrease'
    else:
        qoq_absorption_increase_or_decrease = 'no change'
    

    net_absorption                  = "{:,.0f}".format(net_absorption)
    previous_quarter_net_absorption = "{:,.0f}".format(previous_quarter_net_absorption)


    #Get the current quarter
    latest_quarter = str(data_frame['Period'].iloc[-1])
    latest_year    = str(data_frame['Year'].iloc[-1])
    if 'Q1' in latest_quarter:
        quarter = 'first'
    elif 'Q2' in latest_quarter:
        quarter = '2nd'
    elif 'Q3' in latest_quarter:
        quarter = 'third'
    elif 'Q4' in latest_quarter:
        quarter = 'fourth'

    #Get the current vacancy rates
    submarket_vacancy = data_frame['Vacancy Rate'].iloc[-1]
    market_vacancy    = data_frame2['Vacancy Rate'].iloc[-1]
    national_vacancy  = data_frame3['Vacancy Rate'].iloc[-1]
   
    year_ago_submarket_vacancy = data_frame['Vacancy Rate'].iloc[-5]

    #Determine if vacancy has grown or compressed
    yoy_submarket_vacancy_growth = data_frame['YoY Vacancy Growth'].iloc[-1]
    yoy_market_vacancy_growth    = data_frame2['YoY Vacancy Growth'].iloc[-1]
    qoq_submarket_vacancy_growth = data_frame['QoQ Vacancy Growth'].iloc[-1]
    qoq_market_vacancy_growth    = data_frame2['QoQ Vacancy Growth'].iloc[-1]

    if yoy_submarket_vacancy_growth > 0:
        yoy_submarket_vacancy_growth_description = 'expanded'
            
    elif yoy_submarket_vacancy_growth < 0:
        yoy_submarket_vacancy_growth_description = 'compressed'

    else:
        yoy_submarket_vacancy_growth_description = 'reamined flat'

    if qoq_submarket_vacancy_growth > 0:
        qoq_submarket_vacancy_growth_description = 'expanded'
            
    elif qoq_submarket_vacancy_growth < 0:
        qoq_submarket_vacancy_growth_description = 'compressed'

    else:
        qoq_submarket_vacancy_growth_description = 'reamined flat'

    yoy_submarket_vacancy_growth = "{:,.0f}".format(abs(yoy_submarket_vacancy_growth))
    yoy_market_vacancy_growth    = "{:,.0f}".format(abs(yoy_market_vacancy_growth))
    
    qoq_submarket_vacancy_growth = "{:,.0f}".format(abs(qoq_submarket_vacancy_growth))
    qoq_market_vacancy_growth    = "{:,.0f}".format(abs(qoq_market_vacancy_growth))


    #Determine if market or submarket
    if data_frame.equals(data_frame2):
        market_or_submarket = 'Market'
        market_or_national  = 'National'
        if market_vacancy > national_vacancy:
            above_or_below  = 'above'
        elif market_vacancy < national_vacancy:
            above_or_below  = 'below'
        else:
            above_or_below  = 'at'


        market_submarket_differnce = abs(market_vacancy - national_vacancy) * 100

    else:
        market_or_submarket = 'Submarket'
        market_or_national  = 'Market'
        if submarket_vacancy > market_vacancy:
            above_or_below  = 'above'
        elif submarket_vacancy < market_vacancy:
            above_or_below  = 'below'
        else:
            above_or_below  = 'at'

        market_submarket_differnce  = abs(market_vacancy - submarket_vacancy) * 100

    #Track 10 year growth in vacancy 
    try:
        lag_ammount              = -41
        lagged_date              = data_frame['Period'].iloc[lag_ammount]
        lagged_submarket_vacancy = data_frame['Vacancy Rate'].iloc[lag_ammount]
        lagged_market_vacancy    = data_frame2['Vacancy Rate'].iloc[lag_ammount]
        lagged_national_vacancy  = data_frame3['Vacancy Rate'].iloc[lag_ammount]
    except:
        lag_ammount              = 0 #if therere arent 10 years of observations, use the first available
        lagged_date              = data_frame['Period'].iloc[lag_ammount]
        lagged_submarket_vacancy = data_frame['Vacancy Rate'].iloc[lag_ammount]
        lagged_market_vacancy    = data_frame2['Vacancy Rate'].iloc[lag_ammount]
        lagged_national_vacancy  = data_frame3['Vacancy Rate'].iloc[lag_ammount]
    
    ten_year_growth = (abs(submarket_vacancy -  lagged_submarket_vacancy)) * 100
    

    if submarket_vacancy > lagged_submarket_vacancy:
        ten_year_growth_description = 'expanded'
    elif  submarket_vacancy < lagged_submarket_vacancy:
        ten_year_growth_description = 'compressed'
    else:
        ten_year_growth_description = 'stayed constant'


        

    #Calculate 10 year average
    submarket_avg_vacancy = data_frame['Vacancy Rate'].mean()
    market_avg_vacancy    = data_frame2['Vacancy Rate'].mean()

    if submarket_vacancy > submarket_avg_vacancy:
        avg_relationship_description = 'above'
    elif submarket_vacancy < submarket_avg_vacancy:
        avg_relationship_description = 'below'
    else:
        avg_relationship_description = 'at'

    #determine if vacancy "expanded", "compressed", or "stayed at" the 10 year average over the past year
    if (data_frame['Vacancy Rate'].iloc[-1] > submarket_avg_vacancy) and (data_frame['Vacancy Rate'].iloc[-5] > submarket_avg_vacancy):
        avg_relationship_change = 'stayed'

    elif (data_frame['Vacancy Rate'].iloc[-1] > submarket_avg_vacancy) and (data_frame['Vacancy Rate'].iloc[-5] < submarket_avg_vacancy):
        avg_relationship_change = 'expanded'
    
    elif (data_frame['Vacancy Rate'].iloc[-1] < submarket_avg_vacancy) and (data_frame['Vacancy Rate'].iloc[-5] > submarket_avg_vacancy):
        avg_relationship_change = 'compressed'
    
    else:
        avg_relationship_change = 'expanded/compressed'

    #Calculate total net absorption so far for the current year and how it compares to the same period last year
           
    data_frame_current_year  = data_frame.loc[data_frame['Year'] == (data_frame['Year'].max())]
    data_frame_previous_year = data_frame.loc[data_frame['Year'] == (data_frame['Year'].max() -1 )]
    current_year_total_net_absorption  = data_frame_current_year[net_absorption_var_name].sum()
    previous_year_total_net_absorption = data_frame_previous_year[net_absorption_var_name].sum()
    
    if previous_year_total_net_absorption == 0 : #Cant divide by 0
        net_absorption_so_far_this_year_percent_change = ''
    else:
        net_absorption_so_far_this_year_percent_change = ((current_year_total_net_absorption/previous_year_total_net_absorption) - 1 ) * 100
        if net_absorption_so_far_this_year_percent_change > 0:
            net_absorption_so_far_this_year_percent_change = "{:,.0f}% increase".format(abs(net_absorption_so_far_this_year_percent_change))
        else:
            net_absorption_so_far_this_year_percent_change = "{:,.0f}% decrease".format(abs(net_absorption_so_far_this_year_percent_change))

    
    #Format Variables
    submarket_avg_vacancy               = "{:,.1f}%".format(submarket_avg_vacancy)
    market_avg_vacancy                  = "{:,.1f}%".format(market_avg_vacancy)
    lagged_submarket_vacancy            = "{:,.1f}%".format(lagged_submarket_vacancy)
    ten_year_growth                     = "{:,.0f}".format(ten_year_growth)
    submarket_vacancy                   = "{:,.1f}%".format(submarket_vacancy)
    year_ago_submarket_vacancy          = "{:,.1f}%".format(year_ago_submarket_vacancy)
    market_vacancy                      = "{:,.1f}%".format(market_vacancy)
    national_vacancy                    = "{:,.1f}%".format(national_vacancy)
    market_submarket_differnce          = "{:,.0f}".format(market_submarket_differnce)
    current_year_total_net_absorption   = "{:,.0f}".format(current_year_total_net_absorption)
    

    return('Leasing activity across the '+
            market_or_submarket +
            ' has ' +
            '(slowed/accelerated/stabilized/been volatile/nonexistent)' +
            ' over the past year but has (outpaced/fallen short) of inventory growth. ' +
            'With demand (falling short of/exceeding) new supply, vacancy rates have ' +
            yoy_submarket_vacancy_growth_description +
            ' ' +
            yoy_submarket_vacancy_growth +
            ' bps over the past year from a rate of ' +
            year_ago_submarket_vacancy +
            ' to the current rate of '+
            submarket_vacancy +
            '.'
            '\n' +
            '\n' +
            'In the ' +
            quarter   +
            ' quarter, the '+
            market_or_submarket +
            ' absorbed ' +
            net_absorption  +
            ' ' +
            unit_or_sqft +
            ', representing a(n) ' +
            qoq_absorption_increase_or_decrease +
            ' from the ' +
            previous_quarter_net_absorption +
            ' ' +
            unit_or_sqft +
            ' of net absorption in the first quarter.' +
            ' With ' +
            net_absorption +
            ' ' +
            unit_or_sqft +
            ' absorbed in the second quarter, vacancy rates have ' +
            qoq_submarket_vacancy_growth_description +
            ' ' +
            qoq_submarket_vacancy_growth +
            ' bps over the last quarter.'
            ' Combined, net absorption through the first '+
            'two' +
            ' quarters of ' +
            '2021' +
            ' totaled ' +
            current_year_total_net_absorption +
            ' ' +
            unit_or_sqft  +
            '.'
            
            
        #     ', which represents a ' +
        #     net_absorption_so_far_this_year_percent_change +
        #     ' ' +
        #     ' from the same time period last year. ' +         
        # 'At ' +
        # submarket_vacancy +
        # ', vacancy rates in the ' +
        # market_or_submarket +
        # ' have ' +
        # yoy_submarket_vacancy_growth_description +
        # ' over the past twelve months as the ' +
        # market_or_submarket +
        # ' ' +
        # demand_description +
        # ' demand. '+
        # 'With net absorption totaling ' +
        # net_absorption +
        # ' ' + 
        # unit_or_sqft + 
        # ' in the ' +
        # quarter +
        # ' quarter' +
        # ', vacancy rates have ' +
        # qoq_submarket_vacancy_growth_description +
        # ' ' +
        # qoq_submarket_vacancy_growth +
        # ' bps compared to the previous quarter and ' +
        # yoy_submarket_vacancy_growth_description +
        # ' ' +
        # yoy_submarket_vacancy_growth +
        # ' bps over the past year. ' +



        '\n' +
        '\n' +
        'Going back ten years,' +
        ' vacancy rates ' +
        'have '+
        ten_year_growth_description +
        ' from ' +
        # ' by ' + 
        # ten_year_growth +
        # ' bps over the past decade from ' +
        lagged_submarket_vacancy +
        ' in ' +
        lagged_date +
        ' to the current rate of ' +
        submarket_vacancy +
        '. Over the past twelve months vacancy rates have ' +
        yoy_submarket_vacancy_growth_description +
        ' ' +
        avg_relationship_description +
        ' the 10-year average of '+
        submarket_avg_vacancy +
        ' and ' +
        'the rate falls '+
        above_or_below +
        ' the ' +
        market_or_national + 
        ' average by ' +
        market_submarket_differnce +
        ' bps.' )






def CreateRentLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    if sector == "Multifamily":
        rent_var        = 'Market Effective Rent/Unit'
        rent_growth_var = 'YoY Market Effective Rent/Unit Growth'
        qoq_rent_growth_var = 'QoQ Market Effective Rent/Unit Growth'
        unit_or_sqft      = 'unit'
    else:
        rent_var = 'Market Rent/SF'
        rent_growth_var = 'YoY Rent Growth'
        qoq_rent_growth_var = 'QoQ Rent Growth'
        unit_or_sqft      = 'SF'

    #Get current rents for submarket, market, and nation
    current_rent         = data_frame[rent_var].iloc[-1]
    primary_market_rent  = data_frame2[rent_var].iloc[-1]
    national_market_rent = data_frame3[rent_var].iloc[-1]
    
    #See how these rents compare to one another 
    primary_rent_discount = round((((current_rent/primary_market_rent) -1 ) * -1) * 100,1)
    national_rent_discount = round((((current_rent/national_market_rent) -1 ) * -1) * 100,1)

    if primary_rent_discount < 0:
        primary_rent_discount =  primary_rent_discount * -1
        chaper_or_more_expensive_primary = 'more expensive'
    else:
        chaper_or_more_expensive_primary = 'cheaper'
    
    if national_rent_discount < 0:
        national_rent_discount =  national_rent_discount * -1
        chaper_or_more_expensive_national = 'more expensive'
    else:
        chaper_or_more_expensive_national = 'cheaper'

    
    
    #Calcuate rent growth for submarket, market, and national average over past 10 years
    submarket_starting_rent             =  data_frame[rent_var].iloc[0]
    submarket_start_period              =  str(data_frame['Period'].iloc[0])
    submarket_year_ago_period           =  str(data_frame['Period'].iloc[-5])
    submarket_yoy_growth                =  data_frame[rent_growth_var].iloc[-1]
    submarket_qoq_growth                =  data_frame[qoq_rent_growth_var].iloc[-1]
    submarket_year_ago_yoy_growth       =  data_frame[rent_growth_var].iloc[-5]
    submarket_pre_pandemic_yoy_growth   =  data_frame[rent_growth_var].iloc[-6] #2020 Q1 Annual Growth if still in 2021 Q2
    submarket_decade_rent_growth        = round(((current_rent/submarket_starting_rent) - 1) * 100,1)
    submarket_decade_rent_growth_annual = submarket_decade_rent_growth/10
    
    if submarket_decade_rent_growth > 0:
        submarket_annual_growth_description = 'grown'
        submarket_annual_growth_description2 = 'increase'
    elif submarket_decade_rent_growth < 0:
        submarket_annual_growth_description = 'decreased'
        submarket_annual_growth_description2 = 'decrease'
    else:
        submarket_annual_growth_description = 'remained'
        submarket_annual_growth_description2 = '-'
    

    market_starting_rent             =  data_frame2[rent_var].iloc[0]
    market_yoy_growth                =  data_frame[rent_growth_var].iloc[-1]
    market_decade_rent_growth        = round(((primary_market_rent/market_starting_rent) - 1) * 100,1)
    market_decade_rent_growth_annual = market_decade_rent_growth/10

    if market_decade_rent_growth > 0:
        market_annual_growth_description = 'grown'
        market_annual_growth_description2 = 'increase'

    elif market_decade_rent_growth < 0:
        market_annual_growth_description = 'decreased'
        market_annual_growth_description2 = 'decrease'
    else:
        market_annual_growth_description = 'remained'
        market_annual_growth_description2 = '-'

    

    national_starting_rent             =  data_frame3[rent_var].iloc[0]
    national_decade_rent_growth        = round(((national_market_rent/national_starting_rent) - 1) * 100,1)
    national_decade_rent_growth_annual = national_decade_rent_growth/10
    
    #See if submarket grew faster than market and if market grew faster than nation
    if market_decade_rent_growth > national_decade_rent_growth:
        market_national_faster_or_slower = 'faster'
    elif market_decade_rent_growth < national_decade_rent_growth:
        market_national_faster_or_slower = 'slower'
    else:
        market_national_faster_or_slower = 'the same pace as'

    if submarket_decade_rent_growth > market_decade_rent_growth:
        submarket_market_faster_or_slower = 'faster'
    elif submarket_decade_rent_growth < market_decade_rent_growth:
        submarket_market_faster_or_slower = 'slower'
    else:
        submarket_market_faster_or_slower = 'the same pace as'

    #Describe YOY growth
    if submarket_yoy_growth < 0:
        submarket_yoy_growth_description = 'compressed'
        submarket_signal                 = 'will likely compress further' 
    elif submarket_yoy_growth > 0:
        submarket_yoy_growth_description = 'expanded'
        submarket_signal                 = 'are starting to rebound' 

    else:
        submarket_yoy_growth_description = 'remained at'
        submarket_signal                 = 'are staying put' 
    
    #Describe YOY Growth 1 year ago
    if submarket_year_ago_yoy_growth > 0:
        submarket_year_ago_yoy_growth_description = 'accelerating'
    elif submarket_year_ago_yoy_growth < 0:
        submarket_year_ago_yoy_growth_description = 'decelerating'
    else:
        submarket_year_ago_yoy_growth_description = 'stable'

    #Describe Prepandemic Growth 
    if submarket_pre_pandemic_yoy_growth > 0:
        submarket_pre_pandemic_yoy_growth_description = 'accelerating'
    elif submarket_pre_pandemic_yoy_growth < 0:
        submarket_pre_pandemic_yoy_growth_description = 'decelerating'
    else:
        submarket_pre_pandemic_yoy_growth_description = 'stable' 




    if market_yoy_growth < 0:
        market_yoy_growth_description = 'compressed'
        market_signal                 =  'will likely compress further' 
    elif market_yoy_growth > 0:
        market_yoy_growth_description = 'expanded'
        market_signal                 = 'are starting to rebound' 
    else:
        market_yoy_growth_description = 'remained at'
        market_signal                 = 'are staying put'

    
    #Variable Formatting (We use absolute value function because we already have words in variables to describe if growth is negative or positve)
    if sector == "Multifamily":
        national_rent_discount               = "{:,.0f}%".format(national_rent_discount)
        current_rent                         = "${:,.0f}".format(current_rent)
        submarket_starting_rent              = "${:,.0f}".format(submarket_starting_rent)
        market_starting_rent                 = "${:,.0f}".format(market_starting_rent)
        national_market_rent                 = "${:,.0f}".format(national_market_rent)
        submarket_decade_rent_growth         = "{:,.1f}%".format(abs(submarket_decade_rent_growth))
        submarket_decade_rent_growth_annual  = "{:,.1f}%".format(abs(submarket_decade_rent_growth_annual))
        submarket_yoy_growth                 = "{:,.1f}%".format(abs(submarket_yoy_growth))
        submarket_qoq_growth                 = "{:,.1f}%".format(submarket_qoq_growth)
        submarket_year_ago_yoy_growth        = "{:,.1f}%".format(submarket_year_ago_yoy_growth)
        submarket_pre_pandemic_yoy_growth    = "{:,.1f}%".format(submarket_pre_pandemic_yoy_growth)
        market_decade_rent_growth            = "{:,.1f}%".format(abs(market_decade_rent_growth))
        market_decade_rent_growth_annual     = "{:,.1f}%".format(abs(market_decade_rent_growth_annual))
        market_yoy_growth                    = "{:,.1f}%".format(abs(market_yoy_growth))
        national_decade_rent_growth          = "{:,.1f}%".format(national_decade_rent_growth)
        national_decade_rent_growth_annual   = "{:,.1f}%".format(abs(national_decade_rent_growth_annual))
        primary_rent_discount                = "{:,.0f}%".format(primary_rent_discount)
        primary_market_rent                  = "${:,.0f}".format(primary_market_rent)
        
        
    else:
        national_rent_discount               = "{:,.0f}%".format(national_rent_discount)
        current_rent                         = "${:,.2f}".format(current_rent)
        submarket_starting_rent              = "${:,.2f}".format(submarket_starting_rent)
        market_starting_rent                 = "${:,.2f}".format(market_starting_rent)
        national_market_rent                 = "${:,.2f}".format(national_market_rent)
        submarket_decade_rent_growth         = "{:,.1f}%".format(abs(submarket_decade_rent_growth))
        submarket_decade_rent_growth_annual  = "{:,.1f}%".format(abs(submarket_decade_rent_growth_annual))
        submarket_yoy_growth                 = "{:,.1f}%".format(abs(submarket_yoy_growth))
        submarket_qoq_growth                 = "{:,.1f}%".format(submarket_qoq_growth)
        submarket_year_ago_yoy_growth        = "{:,.1f}%".format(submarket_year_ago_yoy_growth)
        submarket_pre_pandemic_yoy_growth    = "{:,.1f}%".format(submarket_pre_pandemic_yoy_growth)
        market_decade_rent_growth            = "{:,.1f}%".format(abs(market_decade_rent_growth))
        market_decade_rent_growth_annual     = "{:,.1f}%".format(abs(market_decade_rent_growth_annual))
        market_yoy_growth                    = "{:,.1f}%".format(abs(market_yoy_growth))
        national_decade_rent_growth          = "{:,.1f}%".format(national_decade_rent_growth)
        national_decade_rent_growth_annual   = "{:,.1f}%".format(abs(national_decade_rent_growth_annual))
        primary_rent_discount                = "{:,.0f}%".format(primary_rent_discount)
        primary_market_rent                  = "${:,.2f}".format(primary_market_rent)

    #Determine if market or submarket
    if data_frame.equals(data_frame2): #Market
        market_or_submarket = 'Market'
        market_or_nation    = 'National average'
        
        
        return( 'At ' +
            current_rent +
            '/' +
            unit_or_sqft +
            ', rents in the ' +
            market_or_submarket + 
            ' are roughly ' +
            national_rent_discount +
            ' ' +
            chaper_or_more_expensive_national +
            ' than the ' +
            market_or_nation +
            ' where rents sit at ' +
            national_market_rent +
             '/' +
            unit_or_sqft +
            '. ' +
            'Rents in the ' +
            market_or_submarket +
            ' have ' +
            market_annual_growth_description +
            ' '  +
            market_decade_rent_growth +
            ' over the last decade from ' +
            market_starting_rent +
               '/' +
            unit_or_sqft +
            ' in ' +
            submarket_start_period +
            ', representing an annual ' +
           market_annual_growth_description2 +
            ' of ' +
            market_decade_rent_growth_annual +
            ', in line with/falling short of/exceeding the ' +
            market_or_nation +
            ', where rents ' +
            'expanded ' +
            national_decade_rent_growth_annual +
            ' per annum during that time. ' +
            'Prior to ' +
              'the pandemic' +
            ', rents in the '+
            market_or_submarket +
            ' were ' +
            submarket_pre_pandemic_yoy_growth_description +  
            ' with annual growth of '+
            submarket_pre_pandemic_yoy_growth +
            '. Shutdowns occurred in March, slowing demand and softening rent growth over the course of the year.' +
            ' Rent growth has ' +
            'picked up/slowed/remained steady' +
            ' over the first half of 2021' +
            ' with quarterly growth in Q2 reaching ' +
            submarket_qoq_growth +
            '. On an annual basis ' +
            market_or_submarket +
            ' rents have ' +
            market_yoy_growth_description +
            ' ' +
            market_yoy_growth +
            ', pointing to possible signs that rents ' +
           market_signal +
            ' in the near term.' 

    )   

    else: #Submarket
        market_or_submarket = 'Submarket'
        market_or_nation    = 'Market'

        return( 'At ' +
            current_rent +
            '/' +
            unit_or_sqft +
            ', rents in the ' +
            market_or_submarket + 
            ' are roughly ' +
            primary_rent_discount +
            ' ' +
            chaper_or_more_expensive_primary +
            ' than the ' +
            market_or_nation +
            ' where rents sit at ' +
            primary_market_rent +
             '/' +
            unit_or_sqft +
            '. ' +
            'Rents in the ' +
            market_or_submarket +
            ' have ' +
            submarket_annual_growth_description +
            ' '  +
            submarket_decade_rent_growth +
            ' over the last decade from ' +
            submarket_starting_rent +
               '/' +
            unit_or_sqft +
            ' in ' +
            submarket_start_period +
            ', representing an annual ' +
           submarket_annual_growth_description2 +
            ' of ' +
            submarket_decade_rent_growth_annual +
           ', in line with/falling short of/exceeding the ' +
            market_or_nation +
            ', where rents expanded ' +
            market_decade_rent_growth_annual +
            ' per annum during that time. ' +
            'Leading up to ' +
            'the pandemic' +
            ', rents in the '+
            market_or_submarket +
            ' were ' +
            submarket_pre_pandemic_yoy_growth_description +  
            ' with annual growth of '+
            submarket_pre_pandemic_yoy_growth +
            '. Despite//With shutdowns occuring in March and April 2020, demand picked up//slowed, accelerating//softening rent growth over the course of the year//temporarily.' +
            ' Rent growth has ' +
            'picked up/slowed/remained steady' +
            ' over the first half of 2021' +
            ' with quarterly growth in Q2 reaching ' +
            submarket_qoq_growth +
            ' On an annual basis ' +
            market_or_submarket +
            ' rents have ' +
            submarket_yoy_growth_description +
            ' ' +
            submarket_yoy_growth +
            ', pointing to possible signs that rents ' +
           submarket_signal +
            ' in the near term.' 

    )   

       


    


             
def CreateConstructionLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    if sector == "Multifamily":
        unit_or_sqft                        = 'units'
        under_construction                  = data_frame['Under Construction Units'].iloc[-1]
        previous_quarter_under_construction = data_frame['Under Construction Units'].iloc[-2]
        under_construction_share            = round(data_frame['Under Construction %'].iloc[-1],2)
        current_inventory      = data_frame['Inventory Units'].iloc[-1]
        decade_ago_inventory   = data_frame['Inventory Units'].iloc[0]
        
    else:
        unit_or_sqft                        = 'square feet'
        under_construction                  = data_frame['Under Construction SF'].iloc[-1]
        previous_quarter_under_construction = data_frame['Under Construction SF'].iloc[-2]
        under_construction_share            = round(data_frame['Under Construction %'].iloc[-1],2)
        current_inventory                   = data_frame['Inventory SF'].iloc[-1]
        decade_ago_inventory                = data_frame['Inventory SF'].iloc[0]

    #Determine if the supply pipeline is active or not    
    if under_construction > 0:
        active_or_inactive = 'active'
        empty_or_active    = 'active'
        upward_or_limited  = 'upward'
        well_or_poorly     = 'poorly'
    else:
        active_or_inactive = 'inactive'
        empty_or_active    = 'empty'
        upward_or_limited  = 'limited'
        well_or_poorly     = 'well'


    

    #Determine 10 year inventory growth
    inventory_growth       = current_inventory - decade_ago_inventory
    inventory_growth_pct   = round((inventory_growth/decade_ago_inventory)  * 100,2)

    if inventory_growth > 0:
        inventory_expand_or_contract = 'expanded'
        inventory_increase_or_decrease = 'increase'
    else:
        inventory_expand_or_contract = 'contracted'
        inventory_increase_or_decrease = 'decrease'
    
    if under_construction > previous_quarter_under_construction:
        elevated_or_down_compared_to_previous_quarter = 'elevated'
    elif under_construction < previous_quarter_under_construction:
         elevated_or_down_compared_to_previous_quarter = 'down'
    else:
        elevated_or_down_compared_to_previous_quarter = 'constant'



    #Format variables
    under_construction                               = "{:,.0f}".format(under_construction) 
    previous_quarter_under_construction              = "{:,.0f}".format(previous_quarter_under_construction)  
    under_construction_share                         = "{:,.1f}%".format(under_construction_share)  
    inventory_growth_pct                             = "{:,.0f}%".format(abs(inventory_growth_pct)) 
    inventory_growth                                 = "{:,}".format(abs(inventory_growth))  

    if data_frame.equals(data_frame2):
        market_or_submarket = 'Market'
    else:
        market_or_submarket = 'Submarket'
    
    try:
        return('With ' +
            under_construction +
            ' ' +
            unit_or_sqft +
            ', or the equivalent of ' +
            under_construction_share +
            ' of existing inventory, underway, ' +
        'developers are ' +
            active_or_inactive +
            ' in the ' +
            market_or_submarket +
            '. Over the past ten years, developers have ' +
           inventory_expand_or_contract +
            ' inventory by ' +
            inventory_growth +
            ' ' +
            unit_or_sqft +
            ', representing a ' +
            inventory_growth_pct +
            ' ' + inventory_increase_or_decrease +
            '. ' +
            ' Current development levels are ' +
             elevated_or_down_compared_to_previous_quarter +
             ' compared to ' +
                previous_quarter_under_construction +
                ' ' +
                unit_or_sqft +
                ' under construction in the previous quarter. ' + 
                'A few/No notable projects are set to deliver soon. ' +
                'With an ' +
                'elevated/stable/inactive' +
                ' supply pipeline, vacancies will likely '+
                'expand/see some upward pressure/see limited upward pressure' +
                ' from supply, limiting improvement in/boding well for fundamentals in the near term.' )








    #If there's a problem putting the language together, return a general paragraph we can edit
    except:
            return('Developers are currently ' +
            'active/inactive' +
            ' in the ' +
            market_or_submarket +
            '. This represents a change from the typical trend.' +
            ' In fact, the ' +
                market_or_submarket +
                ' has seen inventory expand by ' +
                'X' +
                unit_or_sqft +
                ', representing a ' +
                'X%' +
                ' increase/decrease' +
                ' over the past ten years. ' + 
                'There are currently ' +
                'X'                    +
                ' ' +
                unit_or_sqft          +
                ' , or the equivalent of ' +
                under_construction_share +
                ' of existing inventory, underway.' +
                ' This compares to ' +
                'X' +
                ' ' +
                unit_or_sqft +
                ' under construction in the previous quarter.' + 
                ' With an ' +
                'empty/active' +
                ' pipeline, vacancies will likely see some ' +
                'upward/limited' +
                ' pressure, boding ' +
                'well/poorly' +
                ' for fundamentals in the near term.'  )

    


def CreateSaleLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    if sector == "Multifamily":
        unit_or_sqft                        = 'units'
        unit_or_sqft_singular               = 'unit'
    else:
        unit_or_sqft                        = 'square feet'
        unit_or_sqft_singular               = 'SF'

    #Collapse down the data to the annual total sales info
    data_frame['n'] = 1
    data_frame2['n'] = 1
    data_frame3['n'] = 1
    
    data_frame_annual = data_frame.groupby('Year').agg(sale_volume=('Total Sales Volume', 'sum'),
                                                transaction_count=('Sales Volume Transactions', 'sum'),
                                                n = ('n','sum')
                                                )
                                                
    data_frame_annual = data_frame_annual.reset_index()
    try:
        data_frame_annual = data_frame_annual.loc[data_frame_annual['n'] == 4] #keep only years where we have 4 full quarters
        data_frame_annual = data_frame_annual.iloc[[-1,-2,-3]]          #keep the last 3 (full) years
        
        three_year_avg_sale_volume       = round(data_frame_annual['sale_volume'].mean())
        three_year_avg_sale_volume       = "${:,.0f}".format(three_year_avg_sale_volume)
        three_year_avg_transaction_count = round(data_frame_annual['transaction_count'].mean())
        three_year_avg_transaction_count = "{:,.0f}".format(three_year_avg_transaction_count)
    except:
        return('(DID NOT HAVE 3 FULL YEARS OF DATA)')

    #Now that we calculated the average per year stats, get info on latest quarter
    current_sale_volume       = data_frame['Total Sales Volume'].iloc[-1]
    try:
         current_sale_volume = round(current_sale_volume)
    except:
        pass
    current_sale_volume       = "${:,.0f}".format(current_sale_volume)
    current_transaction_count = str(round(data_frame['Sales Volume Transactions'].iloc[-1]))
    current_period = str(data_frame['Period'].iloc[-1])
    
    #Determine if investors are typically active here
    #If theres at least 1 sale per quarter, active
    if data_frame['Sales Volume Transactions'].median() >= 1:
        investors_active_or_inactive = 'active'
    else:
        investors_active_or_inactive = 'inactive'

    #Calculate the sale volume "over the last year" (last 4 quarters)
    over_last_year_sale_volume  = data_frame['Total Sales Volume'][-1:-5:-1].sum()
    over_last_year_transactions = data_frame['Sales Volume Transactions'][-1:-5:-1].sum()
    if sector == 'Multifamily':
        over_last_year_units = data_frame['Sold Units'][-1:-5:-1].sum()
    else:
        over_last_year_units = data_frame['Sold Building SF'][-1:-5:-1].sum()

    
    over_last_year_sale_volume = "${:,.0f}".format(over_last_year_sale_volume)
    over_last_year_transactions = "{:,.0f}".format(over_last_year_transactions)
    over_last_year_units        = "{:,.0f}".format(over_last_year_units) 

    #Determine the current asset value
    if sector == 'Multifamily':
        asset_value          = data_frame['Asset Value/Unit'].iloc[-1]
        asset_value          = "${:,.0f}".format(asset_value)
        asset_value_change   = data_frame['YoY Asset Value/Unit Growth'].iloc[-1]
    else:
        asset_value          = data_frame['Asset Value/Sqft'].iloc[-1]
        asset_value          = "${:,.0f}".format(asset_value)
        asset_value_change   = data_frame['YoY Asset Value/Sqft Growth'].iloc[-1]

    if asset_value_change > 0:
        asset_value_change_description = 'expanded'
    elif  asset_value_change < 0:
        asset_value_change_description = 'compressed'
    else:
        asset_value_change_description = ''
    asset_value_change          = "{:,.0f}%".format(abs(asset_value_change))
    
    #Determine if market or submarket
    if data_frame.equals(data_frame2):
        submarket_or_market = 'Market'
    else:
        submarket_or_market = 'Submarket'

    #Determine change in cap rate
    cap_rate          = data_frame['Market Cap Rate'].iloc[-1]
    cap_rate_change   = data_frame['YoY Market Cap Rate Growth'].iloc[-1]
    if cap_rate_change > 0:
        cap_rate_change_description = 'expanded'
        cap_rate_change_description2 = 'expanding'

    elif cap_rate_change < 0:
        cap_rate_change_description = 'compressed'
        cap_rate_change_description2 = 'compressing'
    else:
        cap_rate_change_description = 'remained flat'
        cap_rate_change_description2 = ''

    cap_rate_change = "{:,.0f}".format(abs(cap_rate_change))

    return('Investors are typically '  +
           investors_active_or_inactive +
            ' in this ' +
            submarket_or_market +
            '. ' +
            'Going back three years, investors have closed ' +
            three_year_avg_transaction_count +
            ' transactions per annum' +
            ' representing an annual sales average of ' +
            three_year_avg_sale_volume +
            '. ' +
           'Over the past year, there were '+
            over_last_year_transactions +
            ' transactions across ' +
            over_last_year_units +
            ' ' +
           unit_or_sqft +
           ', representing ' +
           over_last_year_sale_volume +
           ' in dollar volume.' +
            ' In ' +
            current_period +
            ', there were ' +
            current_transaction_count +
            ' sales for a total sales volume of ' +
            current_sale_volume +
            '.' +
            ' At '+
            asset_value +
            '/'+
            unit_or_sqft_singular +
            ', values in this ' +
            submarket_or_market +
            ' have ' +
            asset_value_change_description +
            ' ' +
            asset_value_change + 
            ' over the past year, ' +
            'while the market cap rate has ' +
            cap_rate_change_description +
            ' over the past year, ' +
            cap_rate_change_description2 +
             ' ' +
            cap_rate_change +
            ' bps.'
             )

def CreateOutlookLanguage(data_frame,data_frame2,data_frame3,market_title,primary_market,sector):
    if data_frame.equals(data_frame2):
        market_or_submarket = 'Market'
    else:
        market_or_submarket = 'Submarket'

    general_outlook_language = ('Current fundamentals in the ' +
                            market_or_submarket +
                            ' indicate general ' +
                            'stability/instability' +
                            ' in demand while the count of new deliverables have been ' + 
                            'expanding/steady/limited/absent' +
                            '. Together, vacancy rates have managed to ' +
                            'remain stable/expanded considerably/compressed'  +
                            ' over the course of the pandemic. ' +
                            'Rents responded by remaining ' + 
                            'stable/expanding/softening' +
                            '. The general ' +
                            'stability/instability/acceleration/deceleration' +
                            ' in fundamentals have helped improve the capital market, resulting in ' +
                            'stable/accelerating/decelerating' +
                            ' growth in property values across the sector. ' +
                            '\n' +
                            '\n' +
                            'Looking ahead over the ' +
                            '2nd half of ' + 
                            '2021' +
                            ', it is likely that demand will continue to ' +
                            'pick up/stabilize/remain muted' +
                            ' with rents ' +
                            'stabilizing/accelerating/compressing' +
                            ' further. ' +
                            'Although/However' + 
                            ', a(n) ' +
                            'empty/large' + 
                            ' supply pipeline could allow for vacancy to stabilize. ' +
                            'With fundamentals ' +
                            'improving/softening, values will likely ' +
                            'expand/compress/stabilize.')



    if sector == "Multifamily":
        sector_specific_outlook_language=('Strong economic growth and a drastically improving public health situation helped boost multifamily fundamentals in the first half of 2021. With demand and rent growth indicators surging, investors have regained confidence in the sector, and sales volume has returned to more normal levels over the past few quarters. Still, a few headwinds exist that could put upward pressure on vacancies over the next few quarters. The ' + market_or_submarket + ' still faces a robust near-term supply pipeline, and those units will deliver amid a potential slowdown in demand due to seasonality and the fading effects of fiscal stimulus that has helped thousands of people pay rent. Furthermore, single-family starts have ramped up, and the increase in new for-sale housing could draw higher-income renters away from luxury properties. Looking ahead over the next few quarters,')
    
    elif sector == "Office":
        sector_specific_outlook_language=('The first half of 2021 remained in line with pandemic-era trends in terms of office market performance. Although leasing activity has picked up slightly, it remained rather subdued. Many tenants continue to downsize and adopt hybrid work models, limiting demand and rent growth. Investment volume remains subdued, but investors are looking at alternatives such as the medical office sector or single-tenant assets with sticky tenants and lengthy leases in place. Looking ahead over the next few quarters, supply additions will be met with muted demand, limiting improvement in rents and values.')

    elif sector == "Retail":
        sector_specific_outlook_language=('The new year has delivered encouraging news for the retail sector: Retail sales activity surged as the year commenced, vaccine roll-outs are supporting strong consumer confidence metrics, and leasing activity among many tenant segments remains strong.  Such positive news does not, however, overshadow the complexity and nuance that the sector possesses. Indeed, a tale of two recoveries continues to unfold, and property performance continues to vary significantly by subtype, location, class, and tenant composition. Even with the vaccines, it is probable retailers will continue to face turbulence in the coming quarters. Those effects will likely linger for the foreseeable future, impacting demand, rent growth, and the capital markets in the process.')
    
    elif sector == "Industrial":
        sector_specific_outlook_language=("""The new year has brought much needed support to the nation's economy and to its consumers, who continue to buy record amounts of goods online. In response, industrial users continue to seek more warehouse space closer to the consumer as they evolve their supply chains to meet the demand for fast delivery times. Industrial's rent growth prospects continue to lead across sectors, as well, with both retail and office posting rent declines as multifamily gradually regains momentum after plateauing throughout much of 2020. Still, following the national theme, most markets are set to experience a deceleration in rent growth. With such strength prevailing throughout industrial's operating environment, and with other sectors and asset classes registering more volatility and relatively weaker performance, investors continue to aggressively pursue industrial acquisitions. Looking ahead over the next few quarters, demand from consumers, tenants, and investors will continue driving growth in fundamentals.""")


    return(general_outlook_language + '\n' + '\n' + sector_specific_outlook_language)
