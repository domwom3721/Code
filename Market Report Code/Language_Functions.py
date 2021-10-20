import numpy as np
import math
import os
from bs4 import BeautifulSoup

#Function that takes a number as input and writes it in words (eg: 5,000,000 ---> '5 million')
def millify(n,modifier):
    millnames = ['','k',' million',' billion',' trillion']
    
    try:
        n = float(n)
        millidx = max(0,min(len(millnames)-1,
                            int(math.floor(0 if n == 0 else math.log10(abs(n))/3))))
                            
        if n >= 1000000:
            n = modifier + '{:.1f}{}'.format(n / 10**(3 * millidx), millnames[millidx])
        elif n < 1000:
            n = "{:,.0f}".format(n)
        else:
            n =  modifier + '{:.1f}{}'.format(n / 10**(3 * millidx), millnames[millidx])

        #Do this to avoid .0 in numbers
        n = n.replace('.0','',1)
        return(n)

        
    except Exception as e:
        print(e)
        return(n)

#Function that reads in the write up from the saved html file in the CoStar Write Ups folder within the data folder
def PullCoStarWriteUp(section_names,writeup_directory):


    #Pull writeup from the CoStar Html page if we have one saved
    html_file = os.path.join(writeup_directory,'CoStar - Markets & Submarkets.html')
    if  os.path.exists(html_file):
        try:
            with open(html_file) as fp:
                soup = BeautifulSoup(fp, 'html.parser')

            narrative_bodies = soup.find_all("div", {"class": "cscc-narrative-text"})
            narrative_titles = soup.find_all("div", {"class": "cscc-detail-narrative__title"})

            for narrative,title in zip(narrative_bodies,narrative_titles):
                title_text = title.text 
                for section_name in section_names:
                    master_narrative = ''
                    if section_name in title_text:
                        for count,p in enumerate(narrative.find_all("p")):
                            text  = p.get_text()
                            text = text.replace(' 3, & 4 & 5 Star',' Class A, B, and C')
                            text = text.replace('1 & 2 and 3 Star','Class C and below') 
                            text = text.replace(' 1, & 2 & 3 Star',' Class C and below')
                            text = text.replace(' 1, 2, and 3 Star',' Class C and below')
                            text = text.replace(' 4 & 5 Star ',' Class A and B ')
                            text = text.replace('4 & 5 Stars','Class A and B')
                            text = text.replace(' 4 and 5-star  ',' Class A and B ')   
                            text = text.replace(' 4 and 5 Star ',' Class A and B ')
                            text = text.replace(' 4 or 5 Star ',' Class A and B ')
                            text = text.replace(' 4 & 5 Star',' Class A and B')
                            text = text.replace('4&5 Star','Class A and B')
                            text = text.replace('1 & 2 Star','Class C')
                            text = text.replace('2 & 3 Star','Class C')
                            text = text.replace('a 4 Star,','a Class B,')
                            text = text.replace(' 4 Star ',' Class B ')
                            text = text.replace('4 Star','Class B')
                            text = text.replace('4-Star','Class B')
                            text = text.replace(' 5 Star ',' Class A ')
                            text = text.replace('5 Star','Class A')
                            text = text.replace('5-Star','Class A')
                            text = text.replace('5-star','Class A')
                            text = text.replace(' 3 Star ',' Class C ')
                            text = text.replace('3 Star','Class C')
                            text = text.replace('3-Star','Class C')
                            text = text.replace(' 2 Star ',' Class C ')
                            text = text.replace('2 Stars','Class C')
                            text = text.replace('2 Star','Class C')
                            text = text.replace(' 1 Star ',' Class C ')
                            text = text.replace('1 Star','Class C')
                            text = text.replace('20Q1','2020 Q1')
                            text = text.replace('20Q2','2020 Q2')
                            text = text.replace('20Q3','2020 Q3')
                            text = text.replace('20Q4','2020 Q4')
                            text = text.replace('21Q1','2021 Q1')
                            text = text.replace('21Q2','2021 Q2')
                            text = text.replace('21Q3','2021 Q3')
                            text = text.replace('21Q4','2021 Q4')
                            text = text.replace('amount','number')
                            text = text.replace('2021q1','2021 Q1')
                            text = text.replace('2021q2','2021 Q2')
                            text = text.replace('2021q3','2021 Q3')
                            text = text.replace('2021q4','2021 Q4')
                            #our generated language is delayed compared to costar 
                            # text = text.replace('fourth quarter of 2021','3Q 2021')

                            #Now remove bad characters
                            for char in ['ï','»','¿','â','€']:
                                text = text.replace(char,'')


                            if count == 0:
                                master_narrative = master_narrative       + text
                            else:
                                master_narrative = master_narrative + '\n' + '\n' +text

                    if len(master_narrative) > 1:     
                        return(master_narrative)
            return('')

        except Exception as e:
            print(e)
            return('')
    else:
        return('')

#Langauge for overview section
def CreateOverviewLanguage(submarket_data_frame,market_data_frame,natioanl_data_frame,market_title,primary_market,sector,writeup_directory):

    #Pull writeup from the CoStar Html page if we have one saved
    CoStarWriteUp = PullCoStarWriteUp(section_names= ['Summary'],writeup_directory = writeup_directory)
    if CoStarWriteUp != '':
        return(CoStarWriteUp)
    
    #Section 1: Begin making variables for the overview language that come from the data: 
    if sector == 'Multifamily':
        yoy_rent_growth                 = submarket_data_frame['YoY Market Effective Rent/Unit Growth'].iloc[-1]
        qoq_rent_growth                 = submarket_data_frame['QoQ Market Effective Rent/Unit Growth'].iloc[-1]
        under_construction              = submarket_data_frame['Under Construction Units'].iloc[-1]
        under_construction_share        = submarket_data_frame['Under Construction %'].iloc[-1]
        asset_value                     = submarket_data_frame['Asset Value/Unit'].iloc[-1]         #Get current asset value
        asset_value_change              = submarket_data_frame['YoY Asset Value/Unit Growth'].iloc[-1]
        net_absorption_var_name         = 'Absorption Units'
        submarket_inventory            = submarket_data_frame['Inventory Units'].iloc[-1]
        market_inventory               = market_data_frame['Inventory Units'].iloc[-1]
        unit_or_sqft                    = 'unit'
        unit_or_sqft_singular           = 'unit'
        extra_s                         = 's'


    else: #non multifamily
        yoy_rent_growth                 = submarket_data_frame['YoY Rent Growth'].iloc[-1]
        yoy_rent_growth                 = yoy_rent_growth
        qoq_rent_growth                 = submarket_data_frame['QoQ Rent Growth'].iloc[-1]
        under_construction              = submarket_data_frame['Under Construction SF'].iloc[-1]
        under_construction_share        = submarket_data_frame['Under Construction %'].iloc[-1]
        #Get current asset value
        asset_value                     = submarket_data_frame['Asset Value/Sqft'].iloc[-1]
        asset_value_change              = submarket_data_frame['YoY Asset Value/Sqft Growth'].iloc[-1]
        net_absorption_var_name         = 'Net Absorption SF'
        #Get Submarket and market inventory and the fraction of the inventory the submarket makes up
        submarket_inventory             = submarket_data_frame['Inventory SF'].iloc[-1]
        market_inventory                = market_data_frame['Inventory SF'].iloc[-1]
        unit_or_sqft                    = 'square feet'
        unit_or_sqft_singular           = 'SF'
        extra_s                         = ''
    
    submarket_inventory_fraction        = (submarket_inventory/market_inventory) * 100
    current_sale_volume                 = submarket_data_frame['Total Sales Volume'].iloc[-1]
    current_transaction_count           = submarket_data_frame['Sales Volume Transactions'].iloc[-1]
    vacancy                             = submarket_data_frame['Vacancy Rate'].iloc[-1]
    vacancy_change                      = submarket_data_frame['YoY Vacancy Growth'].iloc[-1]
    avg_vacancy                         = submarket_data_frame['Vacancy Rate'].mean()

    #Get most recent cap rate and change in cap rate
    year_ago_cap_rate                   = submarket_data_frame['Market Cap Rate'].iloc[-5] 
    cap_rate                            = submarket_data_frame['Market Cap Rate'].iloc[-1] 
    avg_cap_rate                        = submarket_data_frame['Market Cap Rate'].mean() 
    cap_rate_yoy_change                 = submarket_data_frame['YoY Market Cap Rate Growth'].iloc[-1]


    #Section 2: Begin making variables that are conditional upon the variables created from the data itself

    #Describe YoY change in asset values
    if asset_value_change > 0:
        asset_value_change_description = 'expanded'
    elif asset_value_change < 0:
        asset_value_change_description = 'compressed'
    else:
        asset_value_change_description = 'remained steady'


    #Relationship betweeen current cap rate and the historical average

    #Cap Rate Below historical average
    if cap_rate < avg_cap_rate:
        
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'falling below'

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'remaining below'
        
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'falling below'

    #Cap Rate above historical average
    elif cap_rate > avg_cap_rate:
                
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'remaining above'
            

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'moving above'
            
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'moving above'
            

        # cap_rate_above_below_average = 'above'
    
    #Cap equals  historical average
    elif  cap_rate == avg_cap_rate:
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'falling to'

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'moving to'
        
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'remaining at'

    #Describe YoY change in cap rates
    if cap_rate_yoy_change > 0:
        cap_rate_change_description = 'expanded '
    elif cap_rate_yoy_change < 0:
        cap_rate_change_description = 'compressed '
    elif cap_rate_yoy_change == 0 :
        cap_rate_change_description = 'seen minimal movement'

   
    #Describe out change in fundamentals
    if yoy_rent_growth >= 0     and vacancy_change <= 0: #if rent is growing (or flat) and vacancy is falling (or flat) we call fundamentals improving
        fundamentals_change = 'improving'
    elif yoy_rent_growth < 0 and vacancy_change > 0 : #if rent is falling and vacancy is rising we call fundamentals softening
        fundamentals_change = 'softening'
    elif (yoy_rent_growth > 0   and vacancy_change  > 0) or (yoy_rent_growth < 0 and vacancy_change < 0 ) : #if rents are falling but vacancy is also falling OR vice versa, then mixed
        fundamentals_change = 'mixed'
    elif (yoy_rent_growth == 0 and vacancy_change == 0): #no change in rent or vacancy
        fundamentals_change = 'stable'
    else:
        fundamentals_change = '[improving/softening/mixed/stable]'

    #Determine if market or submarket
    if submarket_data_frame.equals(market_data_frame):
        market_or_submarket = 'Market'
    else:
        market_or_submarket = 'Submarket'

    #Create the sector sepecific language
    #Retail specific langauge
    if sector == "Retail":
        
        #Negative Rent Growth, positive vacancy growth
        if yoy_rent_growth < 0 and vacancy_change > 0:
            overview_sector_specific_language =  ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'The pandemic appears to have accelerated that trend in the ' +
                                market_or_submarket +
                                '. ' +
                                'This disruption has expanded vacancy rates ' + "{:,.0f} bps".format(vacancy_change) + ' to ' + "{:,.1f}%".format(vacancy) + '. ' + 'With vacancy rates expanding over the past year, rents have contracted ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + '. ') 
        
        #Negative Rent Growth, Negative vacancy growth
        elif yoy_rent_growth < 0 and vacancy_change < 0:
            overview_sector_specific_language =  ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'Despite vacancy rate compression in the ' +
                                market_or_submarket + ' over the past year, rents contracted, decreasing ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Negative rent growth, no vacancy growth
        elif  yoy_rent_growth < 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                                'Despite no change in vacancy, rents have contracted '  "{:,.1f}%".format(abs(yoy_rent_growth)) + ' over the past year. ')
        
        #Positive rent growth, positive vacancy growth
        elif yoy_rent_growth > 0 and vacancy_change > 0:
            overview_sector_specific_language =  (' Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'Despite vacancy rate expansion in the ' +
                                market_or_submarket +
                                ' over the past year, rents managed to grow, expanding ' + "{:,.1f}%".format(yoy_rent_growth) + ' since 2020 Q3. ')

        #Positive rent growth, negative vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'While these trends have continued across the retail sector as a whole, retail properties in the ' + 
                                market_or_submarket + ' have shown resounding strength since the pandemic. In fact, vacancy rates have compressed to ' + "{:,.1f}%".format(vacancy) + ' while rents have expanded ' + "{:,.1f}%".format(yoy_rent_growth) + '. ')  
        
        #Positive rent growth, no vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets.' +
                                             'Despite no change in vacancy, rents have expanded '  "{:,.1f}%".format(yoy_rent_growth) + ' over the past year. ')
                                                
        #no rent growth, negative vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'While these trends have continued across the retail sector as a whole, retail properties in the ' + 
                                market_or_submarket + ' have shown resounding strength since the pandemic. In fact, vacancy rates have compressed to ' + "{:,.1f}%".format(vacancy) + ' while rents have remained stable. ')  
        
        #no rent growth, postive vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change > 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' +
                                                 'Although rent levels have been stable over the past year, vacancy rates have expanded to ' + "{:,.1f}%".format(vacancy) + '. ' )
        
        #no rent growth, no vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets.' +
                                                'However, fundamentals have been stable with no change in rent levels or vacancy over the past year. ')

        else:
            overview_sector_specific_language = ('Prior to 2020 consumer demand was shifting from brick-and-mortar stores towards online channels, putting pressure on vacancy rates and rent growth across most markets. ' + 
                                'While these trends have continued across the retail sector as a whole, retail properties in the ' + 
                                market_or_submarket + ' have shown....')  

    #Create the Multifamily sepecific language
    if sector == "Multifamily":
        #Negative Rent Growth, positive vacancy growth
        if yoy_rent_growth < 0 and vacancy_change > 0:
            overview_sector_specific_language =  ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 
                                 sector + ' properties in the ' + market_or_submarket + 
                                """ have been negatively affected by these shifts in demand, leading to rising vacancy rates and contracting rents. """)
 	
	    #Negative Rent Growth, Negative vacancy growth
        elif yoy_rent_growth < 0 and vacancy_change < 0:
            overview_sector_specific_language =  ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' +  
                                'Despite vacancy rate compression in the ' +
                                market_or_submarket + ' over the past year, multifamily rents have contracted, decreasing ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Negative rent growth, no vacancy growth
        elif  yoy_rent_growth < 0 and vacancy_change == 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' +  
                                'Despite no change in vacancy rates for properties in the ' +
                                market_or_submarket + ' over the past year, multifamily rents have contracted, decreasing ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Positive rent growth, positive vacancy growth
        elif yoy_rent_growth > 0 and vacancy_change > 0:
            overview_sector_specific_language =  ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. The ' + 
                                market_or_submarket + 
                                ' has been positively affected by these shifts, despite an increase in vacancy over the past year. During that time, rents have managed to grow, expanding ' + "{:,.1f}%".format(yoy_rent_growth) + ' since 2020 Q3. ')

        #Positive rent growth, negative vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change < 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 
                                'These shifts have positively affected the ' + 
                                market_or_submarket + ', which has shown resounding strength since the pandemic. In fact, vacancy rates have compressed to ' + "{:,.1f}%".format(vacancy) + ' while rents have expanded ' + "{:,.1f}%".format(yoy_rent_growth) + '. ')  
        
        #Positive rent growth, no vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change == 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 
                                'These shifts have positively affected the ' + 
                                market_or_submarket + ', which has shown resounding strength despite stable vacancy rates. In fact, rents have expanded ' + "{:,.1f}%".format(yoy_rent_growth) + '. ')
        
        #no rent growth, negative vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change < 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 'While vacancy rates have compressed over the past year to ' + "{:,.1f}%".format(vacancy) + ' rents have seen no growth. ')  
        
	    #no rent growth, postive vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change > 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 
                                'These shifts have negatively affected the ' + 
                                market_or_submarket + ', where vacacny rates have expanded to ' + "{:,.1f}%".format(vacancy) + '. Despite this, rents have managed to remain stable. ')
        
        #no rent growth, no vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change == 0:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' + 
                                'These shifts have resulted in steady vacancy rates and no growth in rents. ')
        else:
            overview_sector_specific_language = ("""The unique nature of the pandemic and lockdown dramatically shifted renter preferences, reversing a multi-year trend of urbanization across many of the Nation's largest metros. """ + 
                                'Multiple factors inspired the shift, including an ability to work-from-home, the need for more affordable rents, or the desire for more space. ' ) 
                                
    #Create the Industrial sepecific language
    if sector == "Industrial": 
        
 	    #Negative Rent Growth, positive vacancy growth
        if yoy_rent_growth < 0 and vacancy_change > 0:
            overview_sector_specific_language =  ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand. ' + 
                                'Despite these macro trends, '  + sector.lower() + ' properties in the ' + market_or_submarket + 
                                ' have not felt the affects of these demand drivers, leading to softened levels of leasing activity and rent growth. ')
 	
	    #Negative Rent Growth, Negative vacancy growth
        elif yoy_rent_growth < 0 and vacancy_change < 0:
            overview_sector_specific_language =  ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand. ' + 
                                'Despite these macro trends leading to a decrease in vacancy rates, '  + sector.lower() + 'rents in the market_or_submarket have decreased ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Negative rent growth, no vacancy growth
        elif  yoy_rent_growth < 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand. ' + 
                                'Unfortunately, these macro trends have had little affect on ' + sector.lower() + ' properties in the ' + market_or_submarket + '.' + ' Despite stable vacacny rates, rents have contracted, decreasing ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Positive rent growth, positive vacancy growth
        elif yoy_rent_growth > 0 and vacancy_change > 0:
            overview_sector_specific_language =  ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand. ' + 
                                'Despite vacancy rates expanding over the past year, rents have managed to grow, expanding ' + "{:,.1f}%".format(yoy_rent_growth) + ' since 2020 Q3. ')

        #Positive rent growth, negative vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ' + 
                                'These macro trends have positively affected ' + sector.lower() + ' properties in the ' + market_or_submarket + ' With vacancy rates compressing over the past year, rents have expanded ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Positive rent growth, no vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ' + 
                                'These macro trends have positively affected ' + sector + ' properties in the ' + market_or_submarket + '. While vacancy rates have remained stable, rents have continued to expand, increasing' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #no rent growth, negative vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ' +
                                'Despite these macro trends leading to compressing vacancy rates, industrial rents have seen no growth over the past year. ')
	
        #no rent growth, postive vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change > 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ' +
                                'Despite these macro trends, vacancy rates have expanded over the past year. Fortunately, rents have managed to stay put, but at 0%, are close to moving into negative territory. ' )

        #no rent growth, no vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ')
        else:
            overview_sector_specific_language = ('Industrial enters the fourth quarter in among the best shape of any of the major property types. ' + 
                                'A pandemic driven spike in e-commerce sales along with significant growth in third-party logistics providers continues to drive demand for industrial space. ')

    #Create the Office sepecific language
    if sector == "Office": 

        #Negative Rent Growth, positive vacancy growth
        if yoy_rent_growth < 0 and vacancy_change > 0:
            overview_sector_specific_language =  ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Vacancy and availability rates have expanded as increasingly more businesses and tenants adopt remote work policies. While some markets and submarkets have fared better than others, ' + 
                                  sector.lower() + ' properties in the ' + market_or_submarket + 
                                ' have not. With vacancy rates rising over the year, annual rent growth remains in negative territory. ')
 	
	    #Negative Rent Growth, Negative vacancy growth
        elif yoy_rent_growth < 0 and vacancy_change < 0:
            overview_sector_specific_language =  ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Vacancy and availability rates have expanded as increasingly more businesses and tenants adopt remote work policies. While some markets and submarkets have fared better than others, ' + 
                                  sector.lower() + ' properties in the ' + market_or_submarket +  
                                ' continue to see negative rent growth despite vacancy rate compression. In fact, office rents have decreaseed ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Negative rent growth, no vacancy growth
        elif  yoy_rent_growth < 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Vacancy and availability rates have expanded as increasingly more businesses and tenants adopt remote work policies. ' + 
                                'Despite stable vacancy rates for ' +  sector.lower() + ' properties in the ' + market_or_submarket +  ' rents have contracted, decreasing ' + "{:,.1f}%".format(abs(yoy_rent_growth)) + ' since 2020 Q3. ')
        
        #Positive rent growth, positive vacancy growth
        elif yoy_rent_growth > 0 and vacancy_change > 0:
            overview_sector_specific_language =  ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Vacancy and availability rates have expanded as increasingly more businesses and tenants adopt remote work policies. Some markets and submarkets have fared better than others. ' 
                                'Despite vacancy rates expanding over the past year, rents have managed to grow, expanding ' + "{:,.1f}%".format(yoy_rent_growth) + ' since 2020 Q3. ')

        #Positive rent growth, negative vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Adverse market trends that plagued the office sector during the pandemic are no longer affecting the ' + market_or_submarket + '. ' +
				'With vacancy rates compressing over the year, annual rent growth is in positive territory. In fact, vacancy rates have compressed to ' + "{:,.1f}%".format(vacancy) + ' while rents have expanded ' + "{:,.1f}%".format(yoy_rent_growth) + '. ')  
        
        #Positive rent growth, no vacancy growth
        elif  yoy_rent_growth > 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Vacancy and availability rates have expanded as increasingly more businesses and tenants adopt remote work policies. Some markets and submarkets have fared better than others. ' 
                                'With stable vacancy rates over the past year, rents have managed to grow, expanding ' + "{:,.1f}%".format(yoy_rent_growth) + ' since 2020 Q3. ')
        
        #no rent growth, negative vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change < 0:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Despite vacancy rates compressing over the past year, rents have not expanded, although they do remain stable, which is a welcome sign as well. ')

	    #no rent growth, postive vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change > 0:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ' + 
                                'Despite vacancy rates expanding over the past year, rents have managed to remain stable. ' ) 
        
        #no rent growth, no vacancy growth
        elif  yoy_rent_growth == 0 and vacancy_change == 0:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ')
        
        else:
            overview_sector_specific_language = ('Heading into Q4 2021, some of the adverse market trends established during the pandemic continue to plague the office sector. ')

  
    #Section 3: Format Variables
    under_construction                  = millify(under_construction,'')     
    under_construction_share            = "{:,.0f}%".format(under_construction_share)
    submarket_inventory                 = millify(submarket_inventory,'') 
    market_inventory                    = millify(market_inventory,'') 
    submarket_inventory_fraction        = "{:,.1f}%".format(submarket_inventory_fraction) 
    asset_value                         = "${:,.0f}/". format(asset_value)
    yoy_rent_growth                     = "{:,.1f}%".format(abs(yoy_rent_growth))
    current_sale_volume                 = millify(current_sale_volume,'$')
    current_transaction_count           = "{:,.0f}".format(current_transaction_count) 
    vacancy                             = "{:,.1f}%".format(vacancy)
    avg_vacancy                         = "{:,.1f}%".format(avg_vacancy)
    cap_rate                            = "{:,.1f}%".format(cap_rate)
    cap_rate_yoy_change                 = "{:,.0f} bps".format(abs(cap_rate_yoy_change))
    if cap_rate_yoy_change              == '0 bps':
        cap_rate_yoy_change             = ''
    


    #Section 4: Begin putting sentances together

    #Section 4.1: Create the first subsection (overview_intro_language)
    #Market
    if  market_or_submarket == 'Market':
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
         ' space. '
                                )  

    #Submarket
    else:
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
         ' of the Market’s total inventory. '
                                    )  
      

    #Create the construction sentance
    construcion_sentance = (
                'There are currently ' +
                under_construction +
                ' ' +
                unit_or_sqft +
                extra_s +
                ' underway representing an inventory expansion of ' +
                 under_construction_share +
                '.  '                   
                           )
    #If there is no active construction, change the costruction sentance to be less robotic
    if (construcion_sentance == 'There are currently 0 square feet underway representing an inventory expansion of 0%.  ') or (' 0 units underway' in construcion_sentance):
        construcion_sentance = 'There is no active construction currently underway.  '

    #Section 4.2: Create the conclusion of the overivew language
    overview_conclusion_language = (
                ' '                                +
                "{with_or_despite}".format(with_or_despite = "Despite " if ((asset_value_change_description == 'compressed' and fundamentals_change == 'improving') or (asset_value_change_description == 'expanded' and fundamentals_change == 'softening') )  else "With ") +     
                fundamentals_change                +
                ' fundamentals '                   +
                 ' for '                           +
                 sector.lower()                    +
                 ' properties in the '             +
                 market_or_submarket               +
                ', values have '                   +
                asset_value_change_description     +
                ' over the past year to '          +
                asset_value                        +
                unit_or_sqft_singular              +
                ' and cap rates have '             +
                cap_rate_change_description        +
                cap_rate_yoy_change                +
                ' to a rate of '                   +
                cap_rate                           +
                ', '                               +
                cap_rate_above_below_average       +
                ' the long-term average.'
                                    )

    #Section 4.3: Combine the 3 langauge variables together to form the overview paragraph and return it
    overview_language = (overview_intro_language     + overview_sector_specific_language + overview_conclusion_language)
    return(overview_language)    
    
    
    #Unused code (old below)

    # demand_change                       = data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-1] - data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-5]


     # #Describe change in demand over the last year
    # if demand_change > 0:
    #     demand_change = 'accelerate'
    # elif demand_change < 0:
    #     demand_change = 'slow'
    # elif demand_change == 0:
    #     demand_change = 'remain steady'
    # else:
    #      demand_change = '[accelerate/slow/remained steady]'
    


    # #Describe YoY change in vacancy rates
    # if vacancy_change > 0:
    #     vacancy_change_description = 'expand'
    # elif vacancy_change < 0:
    #     vacancy_change_description = 'compress'
    # elif vacancy_change == 0:
    #     vacancy_change_description = 'remained steady'
    # else:
    #     vacancy_change_description = ''

        # #Describe relationship between change in demand and change in vacancy
    # if demand_change == 'accelerate' and vacancy_change_description == 'compress':
    #     demand_change_vacancy_relationship = 'causing'                          +  ' vacancy rates to '                + vacancy_change_description
    # elif demand_change == 'slow' and vacancy_change_description == 'expand':
    #     demand_change_vacancy_relationship = 'causing'                          +  ' vacancy rates to '                + vacancy_change_description
    
    # #mismatch between demand change and vacancy rates change 
    # elif demand_change == 'slow' and vacancy_change_description == 'compress':
    #     demand_change_vacancy_relationship = 'but'                          +  ' vacancy rates '                + vacancy_change_description
    # elif demand_change == 'accelerate' and vacancy_change_description == 'expand':
    #     demand_change_vacancy_relationship = 'but'                          +  ' vacancy rates '                + vacancy_change_description
    # else:
    #     demand_change_vacancy_relationship = 'causing'                          +  ' vacancy rates to '                + vacancy_change_description

                   

     # #Describe vacancy rates relative to the historical average
    # if vacancy > avg_vacancy:
    #     vacancy_avg_above_or_below = 'above'
    # elif vacancy < avg_vacancy:
    #     vacancy_avg_above_or_below = 'below'
    # elif vacancy == avg_vacancy:
    #     vacancy_avg_above_or_below = 'at'
    # else:
    #     ''

    #Describe cap rates relative to the historical average
    # #Describe YoY rent growth
    # if yoy_rent_growth > 0:
    #     rent_growth_description = 'expanded'
    # elif yoy_rent_growth < 0:
    #     rent_growth_description = 'compressed'
    # else:
    #     rent_growth_description = 'remained steady'

    # #Get Language for rent trends
    # if yoy_rent_growth > 0 and qoq_rent_growth < 0:
    #     rent_growth_description = 'have expanded over the past year but compressed in the past quarter'

    # elif yoy_rent_growth < 0 and qoq_rent_growth < 0:
    #     rent_growth_description = 'have contracted over the past year and continue to soften'


    
        # #Create the capital markets sentance
    # #Write first half of the capital markets section
    # if current_transaction_count == '1':
    #     number_sales_sentanece_fragment  = ('There was only '                +
    #                                         current_transaction_count        +
    #                                         ' sale this quarter'                              
    #                                     )

    # elif current_transaction_count != '0':
    #     number_sales_sentanece_fragment  = ('There were '                     +
    #                                         current_transaction_count        +
    #                                         ' sales this quarter'                              
    #                                     )

    # elif current_transaction_count == '0':
    #     number_sales_sentanece_fragment = 'There were no transactions this quarter'

    # #Write second half of the capital markets section
    # if current_sale_volume != '$0':
    #     sales__volume_sentanece_fragment = (
    #                                     ' for a total sales volume of '       +
    #                                         current_sale_volume                                
    #                                         )
    # else:
    #     sales__volume_sentanece_fragment = ''

    # capital_markets_sentance             =  number_sales_sentanece_fragment +  sales__volume_sentanece_fragment + '.  ' 


    #    ' Over the past twelve months, the ' +
    #     market_or_submarket                +
    #     ' has seen demand '                +
    #     demand_change                      +
    #     ' '                                +
    #     demand_change_vacancy_relationship +
    #     ' to the current rate of '         +
    #     vacancy                            +
    #     '.'                                +
    #     ' Meanwhile, rents '               +
    #     rent_growth_description            +
    #     ' at an annual rate of '           +
    #     yoy_rent_growth                    +
    #     " as of "                          +
    #     latest_quarter                     +
    #     '. '                               +
    #     construcion_sentance               +       
    #     capital_markets_sentance           +

#Language for Supply and Demand Section
def CreateDemandLanguage(submarket_data_frame,market_data_frame,natioanl_data_frame,market_title,primary_market,sector,writeup_directory):
    
    #Pull writeup from the CoStar Html page if we have one saved
    CoStarWriteUp = PullCoStarWriteUp(section_names= ['Vacancy','Supply and Demand', 'Leasing'],writeup_directory = writeup_directory)
    if CoStarWriteUp != '':
        return(CoStarWriteUp)

    #Section 1: Begin making variables for the supply and demand language that come from the data: 
    if sector == 'Multifamily':
        unit_or_sqft                    = 'units'
        net_absorption_var_name         = 'Absorption Units'
        inventory_var_name              = 'Inventory Units'


        net_absorption                  =  submarket_data_frame['Absorption Units'].iloc[-1]
        previous_quarter_net_absorption =  submarket_data_frame['Absorption Units'].iloc[-2]
        covid_quarter_net_absorption    =  submarket_data_frame['Absorption Units'].iloc[-6] #change_each_Q
        # firsthalf2020_net_absorption  =  submarket_data_frame['Absorption Units']
        # year_ago_net_absorption       = submarket_data_frame['Absorption Units'].iloc[-5] #change_each_Q
        #over_last_year_units           = submarket_data_frame['Sold Units'][-1:-5:-1].sum()

    else:
        unit_or_sqft                    = 'square feet'
        net_absorption_var_name         = 'Net Absorption SF'
        inventory_var_name              = 'Inventory SF'
        net_absorption                  =  submarket_data_frame['Net Absorption SF'].iloc[-1]
        previous_quarter_net_absorption =  submarket_data_frame['Net Absorption SF'].iloc[-2]
        covid_quarter_net_absorption    =  submarket_data_frame['Net Absorption SF'].iloc[-6] #change_each_Q
        #availability rate              =  submarket_data_frame['Availability Rate'].iloc[-1]    
        # year_ago_net_absorption       = submarket_data_frame['Net Absorption SF'].iloc[-5] #change_each_Q
        # year_ago_leasing_activity     = submarket_data_frame['Leasing Activity'].iloc[-5] #change_each_Q

    #inventory levels
    submarket_inventory                 = submarket_data_frame[inventory_var_name].iloc[-1]
    decade_ago_inventory                = submarket_data_frame[inventory_var_name].iloc[0]
    ten_year_inventory_growth           = submarket_inventory - decade_ago_inventory       
    inventory_growth_pct                = round(((ten_year_inventory_growth)/decade_ago_inventory)  * 100,2)


    #Get latest quarter and year
    latest_quarter                      = str(submarket_data_frame['Period'].iloc[-1])
    latest_year                         = str(submarket_data_frame['Year'].iloc[-1])
    previous_quarter                    = str(submarket_data_frame['Period'].iloc[-2])

    #Get the current vacancy rates
    previous_year_submarket_vacancy     = submarket_data_frame['Vacancy Rate'].iloc[-5]
    previous_quarter_submarket_vacancy  = submarket_data_frame['Vacancy Rate'].iloc[-2]
    submarket_vacancy                   = submarket_data_frame['Vacancy Rate'].iloc[-1]
    market_vacancy                      = market_data_frame['Vacancy Rate'].iloc[-1]
    national_vacancy                    = natioanl_data_frame['Vacancy Rate'].iloc[-1]
    year_ago_submarket_vacancy          = submarket_data_frame['Vacancy Rate'].iloc[-5]

    #Determine if vacancy has grown or compressed
    yoy_submarket_vacancy_growth        = submarket_data_frame['YoY Vacancy Growth'].iloc[-1]
    yoy_market_vacancy_growth           = market_data_frame['YoY Vacancy Growth'].iloc[-1]
    qoq_submarket_vacancy_growth        = submarket_data_frame['QoQ Vacancy Growth'].iloc[-1]
    qoq_market_vacancy_growth           = market_data_frame['QoQ Vacancy Growth'].iloc[-1]

    #Calculate 10 year average, trough, and peak
    submarket_avg_vacancy               = submarket_data_frame['Vacancy Rate'].mean()
    market_avg_vacancy                  = market_data_frame['Vacancy Rate'].mean()
    # submarket_trough_vacancy            = submarket_data_frame['Vacancy Rate'].min()
    # market_trough_vacancy               = market_data_frame['Vacancy Rate'].min()
    # submarket_peak_vacancy              = submarket_data_frame['Vacancy Rate'].max()
    # market_peak_vacancy                 = market_data_frame['Vacancy Rate'].max()

    # leasing_activity12mo                = submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-1] 
    leasing_change                      = submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-1] -  submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-5]
    inventory_change                    = submarket_data_frame[inventory_var_name].iloc[-1] -  submarket_data_frame[inventory_var_name].iloc[-5]


    #Track 10 year growth in vacancy 
    try:
        lag_ammount                     = -41
        lagged_submarket_vacancy        = submarket_data_frame['Vacancy Rate'].iloc[lag_ammount]
        # lagged_date                     = submarket_data_frame['Period'].iloc[lag_ammount]
        # lagged_market_vacancy           = market_data_frame['Vacancy Rate'].iloc[lag_ammount]
        # lagged_national_vacancy         = natioanl_data_frame['Vacancy Rate'].iloc[lag_ammount]
    except:
        lag_ammount                     = 0 #if therere arent 10 years of observations, use the first available
        lagged_submarket_vacancy        = submarket_data_frame['Vacancy Rate'].iloc[lag_ammount]
        # lagged_date                     = submarket_data_frame['Period'].iloc[lag_ammount]
        # lagged_market_vacancy           = market_data_frame['Vacancy Rate'].iloc[lag_ammount]
        # lagged_national_vacancy         = natioanl_data_frame['Vacancy Rate'].iloc[lag_ammount]
    
    ten_year_growth                     = (abs(submarket_vacancy -  lagged_submarket_vacancy)) * 100
    


    #Section 2: Begin making variables that are conditional upon the variables created from the data itself:

    #Describe quarter over quarter change
    if net_absorption > previous_quarter_net_absorption:
        qoq_absorption_increase_or_decrease = 'an increase'

    elif net_absorption < previous_quarter_net_absorption:
        qoq_absorption_increase_or_decrease = 'a decrease'
    
    else:
        qoq_absorption_increase_or_decrease = 'no change'
    
    #describe current quarter net absorption (vacated if negative, absorbed if positve)
    if net_absorption < 0:        
        net_absorption_description = ' vacated '
    else:
        net_absorption_description = ' absorbed '


    #Get the word to decribe the quarter (first, 2nd, third, fourth)
    if 'Q1' in latest_quarter:
        quarter = 'first'
        number_of_quarters = 'the first quarter of '
        
    elif 'Q2' in latest_quarter:
        quarter = '2nd'
        number_of_quarters = 'the first two quarters of ' 

    elif 'Q3' in latest_quarter:
        quarter = 'third'
        number_of_quarters = 'the first three quarters of ' 

    elif 'Q4' in latest_quarter:
        quarter = 'fourth'
        number_of_quarters = '' 

    #Describe change in vacancy over the past year
    if yoy_submarket_vacancy_growth > 0:
        yoy_submarket_vacancy_growth_description  = 'expanded'
    elif yoy_submarket_vacancy_growth < 0:
        yoy_submarket_vacancy_growth_description  = 'compressed'
    else:
        yoy_submarket_vacancy_growth_description  = 'remained flat'


    #Describe change in vacancy over the past quarter
    if qoq_submarket_vacancy_growth > 0:
        qoq_submarket_vacancy_growth_description = 'expanded'
            
    elif qoq_submarket_vacancy_growth < 0:
        qoq_submarket_vacancy_growth_description = 'compressed'

    else:
        qoq_submarket_vacancy_growth_description = 'remained flat'


    #Determine if market or submarket
    if submarket_data_frame.equals(market_data_frame):
        market_or_submarket = 'Market'
        
        if natioanl_data_frame['Geography Name'].iloc[0]  == 'New York - NY' :
            market_or_national  = 'New York Metro' 
        
        elif natioanl_data_frame['Geography Name'].iloc[0]  == 'United States of America':
            market_or_national    = 'National'

        else:
            market_or_national    = natioanl_data_frame['Geography Name'].iloc[0] 



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

    #Check if vacancy is above or below the historical average

    #Submarket current vacancy above historical average
    if submarket_vacancy > submarket_avg_vacancy:
        #Last year's vacacany rate below historical average
        if previous_year_submarket_vacancy < submarket_avg_vacancy:
            avg_relationship_description = 'moving above'

        #Last year's vacacany rate above historical average
        elif previous_year_submarket_vacancy > submarket_avg_vacancy:
            avg_relationship_description = 'remaining above'

        #Last year's vacacany equal to historical average
        elif  previous_year_submarket_vacancy == submarket_avg_vacancy:
            avg_relationship_description = 'moving above'

    
    #Submarket current vacancy below historical average
    elif submarket_vacancy < submarket_avg_vacancy:
        #Last year's vacacany rate below historical average
        if previous_year_submarket_vacancy < submarket_avg_vacancy:
            avg_relationship_description = 'remaining below'

        #Last year's vacacany rate above historical average
        elif previous_year_submarket_vacancy > submarket_avg_vacancy:
            avg_relationship_description = 'moving below'

        #Last year's vacacany equal to historical average
        elif  previous_year_submarket_vacancy == submarket_avg_vacancy:
            avg_relationship_description = 'moving below'


    #Submarket current vacancy equal to historical average
    elif submarket_vacancy == submarket_avg_vacancy:
        #Last year's vacacany rate below historical average
        if previous_year_submarket_vacancy < submarket_avg_vacancy:
            avg_relationship_description = 'moved to'

        #Last year's vacacany rate above historical average
        elif previous_year_submarket_vacancy > submarket_avg_vacancy:
            avg_relationship_description = 'fell to'

        #Last year's vacacany equal to historical average
        elif  previous_year_submarket_vacancy == submarket_avg_vacancy:
            avg_relationship_description = 'remained at'



        avg_relationship_description = 'at'

    #Calculate total net absorption so far for the current year and how it compares to the same period last year
    data_frame_current_year  = submarket_data_frame.loc[submarket_data_frame['Year'] == (submarket_data_frame['Year'].max())]
    data_frame_previous_year = submarket_data_frame.loc[submarket_data_frame['Year'] == (submarket_data_frame['Year'].max() -1 )]
    data_frame_previous_year = data_frame_previous_year.iloc[0:len(data_frame_current_year)] #make sure we are comparing the same period from the currnet year to the period from last year
    assert (list(data_frame_previous_year['Quarter'])) ==  list((data_frame_current_year['Quarter']))

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


    #This is the first part of the first sentance and explains why vacancy changed
    #Inventory expanded over past year
    if inventory_change > 0:

        #Vacancy increased
        if yoy_submarket_vacancy_growth > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite demand picking up, with rising inventory levels, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'With falling demand and rising inventory levels, vacancy rates have '

               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite no change in demand, with rising inventory levels, vacancy rates have '

            else:
                leasing_activity_intro_clause = 'Vacancy rates have '
                
        #Vacancy decreased
        elif yoy_submarket_vacancy_growth < 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite growing inventory levels, with demand picking up, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling demand and rising inventory levels, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite no change in demand and rising inventory levels, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        #Vacancy flat
        elif yoy_submarket_vacancy_growth == 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite demand picking up, with rising inventory levels, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling demand and rising inventory levels, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite rising inventory levels, with no change in demand, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        else:
            leasing_activity_intro_clause = 'Vacancy rates have '

    #Inventory contracted over the past year
    elif inventory_change < 0:

        #Vacancy increased
        if yoy_submarket_vacancy_growth > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite falling inventory levels and growing demand, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling inventory levels, with falling demand, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite falling demand and no change in demand, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        #Vacancy decreased
        elif yoy_submarket_vacancy_growth < 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'With falling inventory levels and growing demand, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling demand, with falling inventory levels, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'With falling inventory levels and no change in demand, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        #Vacancy flat
        elif yoy_submarket_vacancy_growth == 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite falling inventory and growing demand, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling inventory levels, with demand falling, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite falling inventory levels and no change in demand, vacancy rates have '
        
        else:
            leasing_activity_intro_clause = 'Vacancy rates have '

    #Inventory flat over the past year
    elif inventory_change == 0:

        #Vacancy increased
        if yoy_submarket_vacancy_growth > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite a lack of inventory growth and accelerating demand, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'With no inventory growth, but falling demand, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'Despite a lack of inventory growth and no change in net absorption over the previous 12 months, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        #Vacancy decreased
        elif yoy_submarket_vacancy_growth < 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'With demand picking up in the absence of inventory growth, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Although demand has declined, in the absence of inventory growth, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'With demand and inventory levels flat, vacancy rates have '
            
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '

        

        #Vacancy flat
        elif yoy_submarket_vacancy_growth == 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                leasing_activity_intro_clause = 'Despite rising demand and the absence of inventory growth, vacancy rates have '

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                leasing_activity_intro_clause = 'Despite falling demand, with no inventory growth, vacancy rates have '
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                leasing_activity_intro_clause = 'With demand and inventory levels flat, vacancy rates have '
            else:
                leasing_activity_intro_clause = 'Vacancy rates have '
        
        else:
            leasing_activity_intro_clause = 'Vacancy rates have '
    
    else:
            leasing_activity_intro_clause = 'Vacancy rates have '


    #Section 3: Format Variables
    net_absorption                      = millify(abs(net_absorption),'')
    previous_quarter_net_absorption     = millify(previous_quarter_net_absorption,'')
    covid_quarter_net_absorption        = "{:,.0f}".format(covid_quarter_net_absorption)
    yoy_submarket_vacancy_growth        = "{:,.0f}".format(abs(yoy_submarket_vacancy_growth))
    yoy_market_vacancy_growth           = "{:,.0f}".format(abs(yoy_market_vacancy_growth))
    qoq_submarket_vacancy_growth        = "{:,.0f}".format(abs(qoq_submarket_vacancy_growth))
    qoq_market_vacancy_growth           = "{:,.0f}".format(abs(qoq_market_vacancy_growth))
    submarket_avg_vacancy               = "{:,.1f}%".format(submarket_avg_vacancy)
    market_avg_vacancy                  = "{:,.1f}%".format(market_avg_vacancy)
    lagged_submarket_vacancy            = "{:,.1f}%".format(lagged_submarket_vacancy)
    ten_year_growth                     = "{:,.0f}".format(ten_year_growth)
    submarket_vacancy                   = "{:,.1f}%".format(submarket_vacancy)
    year_ago_submarket_vacancy          = "{:,.1f}%".format(year_ago_submarket_vacancy)
    market_vacancy                      = "{:,.1f}%".format(market_vacancy)
    national_vacancy                    = "{:,.1f}%".format(national_vacancy)
    market_submarket_differnce          = "{:,.0f}".format(market_submarket_differnce)
    current_year_total_net_absorption   = millify(current_year_total_net_absorption,'')
    submarket_inventory                 = millify(submarket_inventory,'')
    # inventory_growth_pct                = "{:,.1f}%".format(inventory_growth_pct)
    # ten_year_inventory_growth           = millify(ten_year_inventory_growth,'')

    #Section 4: Put together the variables we have created into the supply and demand language and return it
    return(
            #Sentance 1 (Supply sentence)
            'The '                                                      +
            market_or_submarket                                         +
            ' is home to '                                              +
            submarket_inventory                                         +
            ' '                                                         +
            unit_or_sqft                                                +
            ' of '                                                      +
            sector.lower()                                              +
            ' space,'                                                   + 
            ' and developers have '                                     +
            "{growth_description}".format(growth_description = "added, net of demolitions," if ten_year_inventory_growth >= 0  else "demolished") +     
            ' '                                                         +
            millify(abs(ten_year_inventory_growth),'')                       +
            ' '                                                         +
            unit_or_sqft                                                +
            ' over the past ten years. '                                +
            'These '                                                    +
            "{development_or_demolitions}".format(development_or_demolitions = "developments" if ten_year_inventory_growth >= 0  else "demolitions") +     
            ' have '                                                     +
            "{expanded_or_reduced}".format(expanded_or_reduced = "expanded" if inventory_growth_pct >= 0  else "reduced") +     
            ' inventory by '                                            +
            "{:,.1f}%".format(abs(inventory_growth_pct))                +
            '. '                                                        +
            


            #Sentence 2
            leasing_activity_intro_clause                               +
            yoy_submarket_vacancy_growth_description                    +
            ' '                                                         +
            yoy_submarket_vacancy_growth                                +
            ' bps over the past year from '                             +
            year_ago_submarket_vacancy                                  +
            ' to '                                                      +
            submarket_vacancy                                           +
            ', '                                                        +
            avg_relationship_description                                +
            ' the 10-year average of '                                  +
            submarket_avg_vacancy                                       +
            ' and '                                                     +
            above_or_below                                              +
            ' the '                                                     +
            market_or_national                                          +     
            ' average'                                                  +
            "{by_bps}".format(by_bps = "" if market_submarket_differnce == '0'  else ( ' by ' + market_submarket_differnce + ' bps') ) +     

            #Sentence 3
            '. In the '                                                 +
            quarter                                                     +
            ' quarter, '                                                +
            sector.lower()                                              + 
            ' tenants in the '                                          +
            market_or_submarket                                         +
            net_absorption_description                                  +
            net_absorption                                              +
            ' '                                                         +
            unit_or_sqft                                                +
            ', '                                                        +
            qoq_absorption_increase_or_decrease                         +
            ' from the '                                                +
            previous_quarter_net_absorption                             +
            ' '                                                         +
            unit_or_sqft                                                +
            ' of net absorption in '                                    +
            previous_quarter                                            + 
            '. '                                                        +
            
            #Sentence 4
            ' With '                                                    +
            net_absorption                                              +
            ' '                                                         +
            unit_or_sqft                                                +
            net_absorption_description                                  +
            'in the '                                                   +
            quarter                                                     +
            ' quarter, vacancy rates have '                             +
            qoq_submarket_vacancy_growth_description                    +
            ' '                                                         +
            qoq_submarket_vacancy_growth                                +
            ' bps since '                                               +
            previous_quarter[5:]                                        +
            '. '                                                        +
            
            #Sentence 5
            'Combined, net absorption through '                         +
            number_of_quarters                                          +
            latest_year                                                 +
            ' totaled '                                                 +
            current_year_total_net_absorption                           +
            ' '                                                         +
            unit_or_sqft                                                +
            '. '
            )  

        # #Describe leasing activity/net abosorption over the past year relative to inventory growth
    # if leasing_change > 0:
    #     leasing_activity_change = 'picked up'
    # elif leasing_change < 0:
    #     leasing_activity_change = 'slowed'
    # elif submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-1] == 0:
    #     leasing_activity_change = 'been nonexistent'
    # else:
    #     leasing_activity_change                          = '[slowed/accelerated/stabilized/been volatile/nonexistent]'

    # if leasing_activity12mo > inventory_change:
    #     demand_fallenshort_or_exceeding_inventorygrowth  = 'exceeded'
    #     # demand_fallingshort_or_exceeding_inventorygrowth = 'exceeding'

    # elif leasing_activity12mo < inventory_change:
    #     demand_fallenshort_or_exceeding_inventorygrowth  = 'fallen short of'
    #     # demand_fallingshort_or_exceeding_inventorygrowth = 'falling short of'
    # else:
    #     demand_fallenshort_or_exceeding_inventorygrowth  = '[fallen short of/exceeded]'
    #     # demand_fallingshort_or_exceeding_inventorygrowth = '[falling short of/exceeding]'
    

        # #Determine conjunction (and or but)
    # if leasing_activity_change == 'picked up' and demand_fallenshort_or_exceeding_inventorygrowth == 'exceeded':
    #     demand_inventory_growth_and_or_but               = 'and' 
    # elif leasing_activity_change == 'slowed' and demand_fallenshort_or_exceeding_inventorygrowth == 'exceeded':
    #     demand_inventory_growth_and_or_but               = 'but'
    # elif leasing_activity_change == 'picked up' and demand_fallenshort_or_exceeding_inventorygrowth == 'fallen short of':
    #     demand_inventory_growth_and_or_but               = 'but'
    # elif leasing_activity_change == 'slowed' and demand_fallenshort_or_exceeding_inventorygrowth == 'fallen short of':
    #     demand_inventory_growth_and_or_but               = 'and'
    # else:
    #     demand_inventory_growth_and_or_but               = '[and/but]'

    # if submarket_vacancy > lagged_submarket_vacancy:
    #     ten_year_growth_description = 'expanded'
    # elif  submarket_vacancy < lagged_submarket_vacancy:
    #     ten_year_growth_description = 'compressed'
    # else:
    #     ten_year_growth_description = 'stayed stead'



    #Old Code (Unused) below:
    #           
    # #Sentance 1
    # 'Leasing activity in the '                                  +
    # market_or_submarket                                         +
    # ' has '                                                     +
    # leasing_activity_change                                     +
    # ' over the past year '                                      +
    # demand_inventory_growth_and_or_but                          +
    # ' has '                                                     + 
    # demand_fallenshort_or_exceeding_inventorygrowth             + 
    # ' inventory growth. '                                       +
    
    # #Sentance 2
    # 'With demand '                                              +
    # demand_fallingshort_or_exceeding_inventorygrowth            +
    # ' new supply, '                                             +

    #determine if vacancy "expanded", "compressed", or "stayed at" the 10 year average over the past year
    # if (submarket_data_frame['Vacancy Rate'].iloc[-1] > submarket_avg_vacancy) and (submarket_data_frame['Vacancy Rate'].iloc[-5] > submarket_avg_vacancy):
    #     avg_relationship_change = 'stayed'

    # elif (submarket_data_frame['Vacancy Rate'].iloc[-1] > submarket_avg_vacancy) and (submarket_data_frame['Vacancy Rate'].iloc[-5] < submarket_avg_vacancy):
    #     avg_relationship_change = 'expanded'
    
    # elif (submarket_data_frame['Vacancy Rate'].iloc[-1] < submarket_avg_vacancy) and (submarket_data_frame['Vacancy Rate'].iloc[-5] > submarket_avg_vacancy):
    #     avg_relationship_change = 'compressed'
    
    # else:
    #     avg_relationship_change = 'expanded/compressed'
                                                    
#Language for rent section
def CreateRentLanguage(submarket_data_frame,market_data_frame,natioanl_data_frame,market_title,primary_market,sector,writeup_directory):

    #Pull writeup from the CoStar Html page if we have one saved
    CoStarWriteUp = PullCoStarWriteUp(section_names= ['Rent',],writeup_directory = writeup_directory)
    if CoStarWriteUp != '':
        return(CoStarWriteUp)

    #Section 1: Begin making variables for the overview language that come from the data: 
    if sector == "Multifamily":
        rent_var                = 'Market Effective Rent/Unit'
        rent_growth_var         = 'YoY Market Effective Rent/Unit Growth'
        qoq_rent_growth_var     = 'QoQ Market Effective Rent/Unit Growth'
        unit_or_sqft            = 'unit'
    else:
        rent_var                = 'Market Rent/SF'
        rent_growth_var         = 'YoY Rent Growth'
        qoq_rent_growth_var     = 'QoQ Rent Growth'
        unit_or_sqft            = 'SF'

    #Get current rents for submarket, market, and nation
    current_rent                     = submarket_data_frame[rent_var].iloc[-1]
    primary_market_rent              = market_data_frame[rent_var].iloc[-1]
    national_market_rent             = natioanl_data_frame[rent_var].iloc[-1]
    
    #See how these rents compare to one another 
    primary_rent_discount            = round((((current_rent/primary_market_rent) -1 ) * -1) * 100,1)
    national_rent_discount           = round((((current_rent/national_market_rent) -1 ) * -1) * 100,1)
    market_starting_rent             =  market_data_frame[rent_var].iloc[0]
    market_yoy_growth                =  submarket_data_frame[rent_growth_var].iloc[-1]
    market_decade_rent_growth        = round(((primary_market_rent/market_starting_rent) - 1) * 100,1)
    market_decade_rent_growth_annual = market_decade_rent_growth/10
    current_period                   = str(submarket_data_frame['Period'].iloc[-1]) #Get most recent quarter

    #Calcuate rent growth for submarket, market, and national average over past 10 years
    submarket_starting_rent                =  submarket_data_frame[rent_var].iloc[0]
    submarket_previous_quarter_yoy_growth  =  submarket_data_frame[rent_growth_var].iloc[-2]
    submarket_yoy_growth                   =  submarket_data_frame[rent_growth_var].iloc[-1]
    submarket_qoq_growth                   =  submarket_data_frame[qoq_rent_growth_var].iloc[-1]
    submarket_year_ago_yoy_growth          =  submarket_data_frame[rent_growth_var].iloc[-5]
    submarket_decade_rent_growth           = round(((current_rent/submarket_starting_rent) - 1) * 100,1)
    submarket_decade_rent_growth_annual    = submarket_decade_rent_growth/10
    national_starting_rent                 =  natioanl_data_frame[rent_var].iloc[0]
    national_decade_rent_growth            = round(((national_market_rent/national_starting_rent) - 1) * 100,1)
    national_decade_rent_growth_annual     = national_decade_rent_growth/10
    
    #Current and past rent growth rates and rent levels
    submarket_pre_2020_average_yoy_rent_growth         = submarket_data_frame.loc[submarket_data_frame['Year'] < 2020][rent_growth_var].mean() #average year over year rent growth before 2020
    submarket_2019Q3_yoy_growth                        = submarket_data_frame.loc[submarket_data_frame['Period'] == '2019 Q3'][rent_growth_var].iloc[-1]  
    submarket_2019Q4_yoy_growth                        = submarket_data_frame.loc[submarket_data_frame['Period'] == '2019 Q4'][rent_growth_var].iloc[-1]  
    submarket_2019Q4_rent                              = submarket_data_frame.loc[submarket_data_frame['Period'] == '2019 Q4'][rent_var].iloc[-1]  
    submarket_2020Q4_yoy_growth                        = submarket_data_frame.loc[submarket_data_frame['Period'] == '2020 Q4'][rent_growth_var].iloc[-1]  
    submarket_2020Q1_qoq_growth                        = submarket_data_frame.loc[submarket_data_frame['Period'] == '2020 Q1'][qoq_rent_growth_var].iloc[-1]  
    submarket_2020Q2_qoq_growth                        = submarket_data_frame.loc[submarket_data_frame['Period'] == '2020 Q2'][qoq_rent_growth_var].iloc[-1]  
    


    #Section 2: Create variables that are conditional on the variables we pulled from the data

    #Describe the relationship between the submarket rent levels compared to the market rent levels
    if primary_rent_discount < 0:
        primary_rent_discount             =  primary_rent_discount * -1
        cheaper_or_more_expensive_primary = 'more expensive'
    else:
        cheaper_or_more_expensive_primary = 'cheaper'

    #Describe the relationship between the market rent levels compared to national rent levels
    if national_rent_discount < 0:
        national_rent_discount             =  national_rent_discount * -1
        cheaper_or_more_expensive_national = 'more expensive'
    else:
        cheaper_or_more_expensive_national = 'cheaper'

    
    #Describe rent growth in the submarket over the past decade
    if submarket_decade_rent_growth > 0:
        submarket_annual_growth_description = 'grown'
    elif submarket_decade_rent_growth < 0:
        submarket_annual_growth_description = 'decreased'
    else:
        submarket_annual_growth_description = 'remained'
    
    #Describe rent growth in the market over the past decade
    if market_decade_rent_growth > 0:
        market_annual_growth_description = 'grown'

    elif market_decade_rent_growth < 0:
        market_annual_growth_description = 'decreased'
    else:
        market_annual_growth_description = 'remained'

    #Describe relationship between quarterly growth and annual rent growth
    if submarket_previous_quarter_yoy_growth > submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'contracting annual growth to'

    elif submarket_previous_quarter_yoy_growth < submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'pushing annual growth to'
    
    elif submarket_previous_quarter_yoy_growth == submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'keeping annual growth at'
    else:
        qoq_pushing_or_contracting_annual_growth = '[contracting/pushing] annual growth to'

      
    #Describe Prepandemic Growth 
    #There's 3 possible starting situations, YoY rent growth in 2019 Q4 was higher than 2019 Q3, lower, or the same, next we need to determine if the growth rate is higher, lower, or in line with the
    #historical average (pre 2020 average)
    if submarket_2019Q4_yoy_growth > submarket_2019Q3_yoy_growth: #rent growth accelerated
        if submarket_2019Q4_yoy_growth > submarket_pre_2020_average_yoy_rent_growth: #above historical average
            submarket_pre_pandemic_yoy_growth_description = 'accelerated above the previous quarter, and was above the historical average,'
        
        elif submarket_2019Q4_yoy_growth < submarket_pre_2020_average_yoy_rent_growth:  #below historical average
            submarket_pre_pandemic_yoy_growth_description = 'accelerated above the previous quarter, but remained below the historical average,'

        elif submarket_2019Q4_yoy_growth == submarket_pre_2020_average_yoy_rent_growth: #equal to historical average
            submarket_pre_pandemic_yoy_growth_description = 'accelerated above the previous quarter, and was in line with the historical average,'

    
    elif submarket_2019Q4_yoy_growth < submarket_2019Q3_yoy_growth: #rent growth softend
        if submarket_2019Q4_yoy_growth > submarket_pre_2020_average_yoy_rent_growth:  #above historical average
            submarket_pre_pandemic_yoy_growth_description = 'softened below the previous quarter, but was above the historical average,'

        elif submarket_2019Q4_yoy_growth < submarket_pre_2020_average_yoy_rent_growth: #below historical average
            submarket_pre_pandemic_yoy_growth_description = 'softened below the previous quarter, and was below the historical average,'

        elif submarket_2019Q4_yoy_growth == submarket_pre_2020_average_yoy_rent_growth: #equal to historical average
            submarket_pre_pandemic_yoy_growth_description = 'softened below the previous quarter, but was in line with the historical average,'

    
    elif submarket_2019Q4_yoy_growth == submarket_2019Q3_yoy_growth: #rent growth constant
        if submarket_2019Q4_yoy_growth > submarket_pre_2020_average_yoy_rent_growth:  #above historical average
            submarket_pre_pandemic_yoy_growth_description = 'remained stable, and was above the historical average,'

        elif submarket_2019Q4_yoy_growth < submarket_pre_2020_average_yoy_rent_growth: #below historical average
            submarket_pre_pandemic_yoy_growth_description = 'remained stable, but was below the historical average,'

        elif submarket_2019Q4_yoy_growth == submarket_pre_2020_average_yoy_rent_growth: #equal to historical average
            submarket_pre_pandemic_yoy_growth_description = 'remained stable and in line with the historical average'
    else: submarket_pre_pandemic_yoy_growth_description = '[accelerated/softend/remained stable]'



    if submarket_data_frame.equals(market_data_frame): #Market
        market_or_submarket = 'Market'

        if natioanl_data_frame['Geography Name'].iloc[0]  == 'New York - NY' :
            market_or_nation  = 'New York Metro average' 
        
        elif natioanl_data_frame['Geography Name'].iloc[0]  == 'United States of America':
            market_or_nation    = 'National average'
        else:
            market_or_nation    = natioanl_data_frame['Geography Name'].iloc[0] + ' average'

        #Check if market decade growth was slower or faster than national growth
        if market_decade_rent_growth_annual > national_decade_rent_growth_annual:
              ten_year_growth_inline_or_exceeding = 'exceeding'
        elif market_decade_rent_growth_annual < national_decade_rent_growth_annual:
            ten_year_growth_inline_or_exceeding = 'falling short of'
        else:
            ten_year_growth_inline_or_exceeding = 'in line with'
    else:
        market_or_submarket = 'Submarket'
        market_or_nation    = 'Market'

        #Check if submakret decade growth was slower or faster than market growth
        if submarket_decade_rent_growth_annual > market_decade_rent_growth_annual:
              ten_year_growth_inline_or_exceeding = 'exceeding'
        elif submarket_decade_rent_growth_annual < market_decade_rent_growth_annual:
            ten_year_growth_inline_or_exceeding = 'falling short of'
        else:
            ten_year_growth_inline_or_exceeding = 'in line with'




    #Section 3: Format Variables
    if sector == "Multifamily":
        national_rent_discount               = "{:,.0f}%".format(national_rent_discount)
        current_rent                         = "${:,.0f}".format(current_rent)
        submarket_2019Q4_rent                = "${:,.0f}".format(submarket_2019Q4_rent)
        submarket_starting_rent              = "${:,.0f}".format(submarket_starting_rent)
        market_starting_rent                 = "${:,.0f}".format(market_starting_rent)
        national_market_rent                 = "${:,.0f}".format(national_market_rent)
        submarket_decade_rent_growth         = "{:,.1f}%".format(abs(submarket_decade_rent_growth))
        submarket_decade_rent_growth_annual  = "{:,.1f}%".format(abs(submarket_decade_rent_growth_annual))
        submarket_yoy_growth                 = "{:,.1f}%".format(abs(submarket_yoy_growth))
        submarket_qoq_growth                 = "{:,.1f}%".format(submarket_qoq_growth)
        submarket_year_ago_yoy_growth        = "{:,.1f}%".format(submarket_year_ago_yoy_growth)
        submarket_2019Q4_yoy_growth          = "{:,.1f}%".format(submarket_2019Q4_yoy_growth)
        # submarket_2020Q4_yoy_growth          = "{:,.1f}%".format(submarket_2020Q4_yoy_growth)
        # submarket_2020Q2_qoq_growth          = "{:,.1f}%".format(submarket_2020Q2_qoq_growth)
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
        submarket_2019Q4_rent                = "${:,.2f}".format(submarket_2019Q4_rent)
        submarket_starting_rent              = "${:,.2f}".format(submarket_starting_rent)
        market_starting_rent                 = "${:,.2f}".format(market_starting_rent)
        national_market_rent                 = "${:,.2f}".format(national_market_rent)
        submarket_decade_rent_growth         = "{:,.0f}%".format(abs(submarket_decade_rent_growth))
        submarket_decade_rent_growth_annual  = "{:,.1f}%".format(submarket_decade_rent_growth_annual)
        submarket_yoy_growth                 = "{:,.1f}%".format(submarket_yoy_growth)
        submarket_qoq_growth                 = "{:,.1f}%".format(submarket_qoq_growth)
        submarket_year_ago_yoy_growth        = "{:,.1f}%".format(submarket_year_ago_yoy_growth)
        submarket_2019Q4_yoy_growth          = "{:,.1f}%".format(submarket_2019Q4_yoy_growth)
        # submarket_2020Q4_yoy_growth          = "{:,.1f}%".format(submarket_2020Q4_yoy_growth)
        # submarket_2020Q2_qoq_growth          = "{:,.1f}%".format(submarket_2020Q2_qoq_growth)
        market_decade_rent_growth            = "{:,.0f}%".format(abs(market_decade_rent_growth))
        market_decade_rent_growth_annual     = "{:,.1f}%".format(abs(market_decade_rent_growth_annual))
        market_yoy_growth                    = "{:,.1f}%".format(market_yoy_growth)
        national_decade_rent_growth          = "{:,.0f}%".format(national_decade_rent_growth)
        national_decade_rent_growth_annual   = "{:,.1f}%".format(abs(national_decade_rent_growth_annual))
        primary_rent_discount                = "{:,.0f}%".format(primary_rent_discount)
        primary_market_rent                  = "${:,.2f}".format(primary_market_rent)
    

    #Section 4: Put togther our rent langauge for either a market or submarket and return it
    if market_or_submarket == 'Market': #Market
        return( 'At ' +
            current_rent +
            '/' +
            unit_or_sqft +
            ', the rents in the ' +
            market_or_submarket + 
            ' are '             +
            "{market_nation_rent_relationship}".format(market_nation_rent_relationship = "equal to " if (primary_market_rent == national_market_rent )  else   ('roughly ' + national_rent_discount + ' ' + cheaper_or_more_expensive_national + ' than ')) +
            'the ' +
            market_or_nation +            
            "{national_rent_level}".format(national_rent_level = "" if (primary_market_rent == national_market_rent)  else     (' where rents sit at ' + national_market_rent + '/' + unit_or_sqft )) +
            '. ' +
            'Rents in the ' +
            market_or_submarket +
            ' have ' +
            market_annual_growth_description +
            ' '                             +
            market_decade_rent_growth_annual +
            ' per annum over the past decade, '+
            ten_year_growth_inline_or_exceeding +
            ' the ' +
            market_or_nation +
            ', where rents ' +
            'expanded ' +
            national_decade_rent_growth_annual +
            ' per annum during that time. ' +
            'In 2019 Q4, annual rent growth in the '+
            market_or_submarket +
            ' ' +
            submarket_pre_pandemic_yoy_growth_description + 
            ' with annual growth of '+
            submarket_2019Q4_yoy_growth +
            '. '                        +
            'In 2020 Q2, quarterly rent growth '              +
            "{growth_description}".format(growth_description = "reached " if submarket_2020Q2_qoq_growth >= submarket_2020Q1_qoq_growth  else "fell to ") +
            "{:,.1f}%".format(submarket_2020Q2_qoq_growth) +
            
            '. By the end of 2020, rents had '             +
            "{growth_description}".format(growth_description = "grown " if  submarket_2020Q4_yoy_growth >= 0  else "fallen ") +                                           
            "{:,.1f}%".format(abs(submarket_2020Q4_yoy_growth)) +
            ' from the 2019 Q4 rent level of '             +
            submarket_2019Q4_rent                          +
            '/'                                            +
            unit_or_sqft +
            '. '
            'Quarterly growth in '                     +
            current_period                              +
            ' reached ' +
            submarket_qoq_growth +
            ', '                 +
           qoq_pushing_or_contracting_annual_growth +
           ' ' +
            market_yoy_growth +
            '.' 
            )   

    elif market_or_submarket == 'Submarket':
        return( 'At ' +
            current_rent +
            '/' +
            unit_or_sqft +
            ', rents in the ' +
            market_or_submarket + 
            ' are '             +
            "{submarket_market_rent_relationship}".format(submarket_market_rent_relationship = "equal to " if (current_rent == primary_market_rent)  else   ('roughly ' + primary_rent_discount + ' ' + cheaper_or_more_expensive_primary + ' than ')) +
            'the ' +
            market_or_nation +
            "{market_rent_level}".format(market_rent_level = "" if (current_rent == primary_market_rent)  else     (' where rents sit at ' + primary_market_rent + '/' + unit_or_sqft )) +
            '. ' +
            
            
            'Rents in the ' +
            market_or_submarket +
            ' have ' +
            submarket_annual_growth_description +
            ' '  +
            submarket_decade_rent_growth_annual +
           ' per annum over the past decade, ' +
           ten_year_growth_inline_or_exceeding +
           ' the ' +
            market_or_nation +
            ', where rents expanded ' +
            market_decade_rent_growth_annual +
            ' per annum during that time. ' +
            'In 2019 Q4, annual rent growth in the '+
            market_or_submarket +
            ' ' +
            submarket_pre_pandemic_yoy_growth_description +  
            ' with annual growth of '                     +
            submarket_2019Q4_yoy_growth                   +   
            '. '                                          +
            'In 2020 Q2, quarterly rent growth '              +
            "{growth_description}".format(growth_description = "reached " if submarket_2020Q2_qoq_growth >= submarket_2020Q1_qoq_growth  else "fell to ") +
            "{:,.1f}%".format(submarket_2020Q2_qoq_growth) +
            
            '. By the end of 2020, rents had '             +
            "{growth_description}".format(growth_description = "grown " if  submarket_2020Q4_yoy_growth >= 0  else "fallen ") +                                           
            "{:,.1f}%".format(abs(submarket_2020Q4_yoy_growth)) +
            ' from the 2019 Q4 rent level of '             +
            submarket_2019Q4_rent                          +
            '/'                                            +
            unit_or_sqft +
            '. '
            'Quarterly growth in ' +
            current_period          +
            ' reached ' +
            submarket_qoq_growth +
            ', '                 +
            qoq_pushing_or_contracting_annual_growth +
             ' ' +
            submarket_yoy_growth +
            '.' 
            )   

        #Old code below
        # submarket_annual_rent_growth_peak   = submarket_data_frame[rent_growth_var].max()


        # submarket_start_period              =  str(submarket_data_frame['Period'].iloc[0])
        # submarket_year_ago_period           =  str(submarket_data_frame['Period'].iloc[-5])
        #market_ytd_rent_growth           = 
        
        # #Describe YOY Growth 1 year ago
        # if submarket_year_ago_yoy_growth > 0:
        #     submarket_year_ago_yoy_growth_description = 'accelerating'
        # elif submarket_year_ago_yoy_growth < 0:
        #     submarket_year_ago_yoy_growth_description = 'decelerating'
        # else:
        #     submarket_year_ago_yoy_growth_description = 'stable'


        # if submarket_decade_rent_growth > market_decade_rent_growth and primary_rent_discount > 0:
        #         decade_rent_and_rent_discount = ' Despite elevated rents compared to the Market, landlords have had no issue pushing rents over the past ten years. '

           # #Calculate 10 year average, trough, and peak
        # if sector == "Multifamily":
        #     submarket_trough_rents               = submarket_data_frame['Market Effective Rent/Unit'].min()
        #     market_trough_rents                 = market_data_frame['Market Effective Rent/Unit'].min()
        #     submarket_peak_rents               = submarket_data_frame['Market Effective Rent/Unit'].max()
        #     market_peak_rents                  = market_data_frame['Market Effective Rent/Unit'].max()
        
        # else:
        #     submarket_trough_rents               = submarket_data_frame['Market Rent/SF'].min()
        #     market_trough_rents                 = market_data_frame['Market Rent/SF'].min()
        #     submarket_peak_rents               = submarket_data_frame['Market Rent/SF'].max()
        #     market_peak_rents                  = market_data_frame['Market Rent/SF'].max()    


        # #See if submarket grew faster than market and if market grew faster than nation
        # if market_decade_rent_growth > national_decade_rent_growth:
        #     market_national_faster_or_slower = 'faster'
        # elif market_decade_rent_growth < national_decade_rent_growth:
        #     market_national_faster_or_slower = 'slower'
        # else:
        #     market_national_faster_or_slower = 'the same pace as'

        # if submarket_decade_rent_growth > market_decade_rent_growth:
        #     submarket_market_faster_or_slower = 'faster'
        # elif submarket_decade_rent_growth < market_decade_rent_growth:
        #     submarket_market_faster_or_slower = 'slower'
        # else:
        #     submarket_market_faster_or_slower = 'the same pace as'

        # #Describe YOY growth for submarket
        # if submarket_yoy_growth < 0:
        #     submarket_yoy_growth_description = 'compressed'
        #     submarket_signal                 = 'will likely compress further' 
        # elif submarket_yoy_growth > 0:
        #     submarket_yoy_growth_description = 'expanded'
        #     submarket_signal                 = 'are starting to rebound' 
        # else:
        #     submarket_yoy_growth_description = 'remained at'
        #     submarket_signal                 = 'are staying put' 

#Language for construction section
def CreateConstructionLanguage(submarket_data_frame, market_data_frame, natioanl_data_frame, market_title, primary_market, sector,writeup_directory):
    #Pull writeup from the CoStar Html page if we have one saved
    CoStarWriteUp = PullCoStarWriteUp(section_names= ['Construction',],writeup_directory = writeup_directory)
    if CoStarWriteUp != '':
        return(CoStarWriteUp)
    
    #Section 1: Begin making variables for the overview language that come from the data:     
    if sector == "Multifamily":
        unit_or_sqft                        = 'units'
        under_construction                  = submarket_data_frame['Under Construction Units'].iloc[-1]
        median_construction_level           = submarket_data_frame['Under Construction Units'].median()
        under_construction_share            = round(submarket_data_frame['Under Construction %'].iloc[-1],2)
        current_inventory                   = submarket_data_frame['Inventory Units'].iloc[-1]
        decade_ago_inventory                = submarket_data_frame['Inventory Units'].iloc[0]
        delivered_inventory                 = submarket_data_frame['Gross Delivered Units'].sum()
        demolished_inventory                = submarket_data_frame['Demolished Units'].sum()                        
        # previous_quarter_under_construction = data_frame['Under Construction Units'].iloc[-2]

    else:
        unit_or_sqft                        = 'square feet'
        under_construction                  = submarket_data_frame['Under Construction SF'].iloc[-1]
        median_construction_level           = submarket_data_frame['Under Construction SF'].median()
        under_construction_share            = round(submarket_data_frame['Under Construction %'].iloc[-1],2)
        current_inventory                   = submarket_data_frame['Inventory SF'].iloc[-1]
        decade_ago_inventory                = submarket_data_frame['Inventory SF'].iloc[0]
        delivered_inventory                 = submarket_data_frame['Gross Delivered SF'].sum()
        demolished_inventory                = submarket_data_frame['Demolished SF'].sum()
        # previous_quarter_under_construction = data_frame['Under Construction SF'].iloc[-2]
    
    yoy_submarket_vacancy_growth        = submarket_data_frame['YoY Vacancy Growth'].iloc[-1]

    if submarket_data_frame.equals(market_data_frame):
        market_or_submarket                 = 'Market'
    else:
        market_or_submarket                 = 'Submarket'

    inventory_growth                        = current_inventory - decade_ago_inventory
    inventory_growth_pct                    = round((inventory_growth/decade_ago_inventory)  * 100,2)
    
    #Section 2: Begin making varaiables that are conditional on the variables we have created in section 1

    #Section 3: Format variables
    # inventory_growth_pct                        = "{:,.1f}%".format(inventory_growth_pct)

    
    #Section 4: Put together our variables into sentances and return the language
    #Determine if developers are historically active here
    #If theres at least 1 deliverable per quarter, active
    if median_construction_level >= 1:
        developers_historically_active_or_inactive = ('Developers have been active for much of the past ten years. In fact, they have added ' + 
                                        millify(delivered_inventory,'')  +
                                        ' '                 +
                                        unit_or_sqft        +
                                        ' to the '          +
                                         market_or_submarket + 
                                        ' over that time, '  +
                                        "{inventory_expand_contract}".format(inventory_expand_contract = "expanding inventory by " if  inventory_growth_pct >= 0  else "but inventory contracted by ") +                                           
                                         "{:,.1f}%".format(abs(inventory_growth_pct)) +
                                          '. '
                                        )
    #Inactive devlopers
    else:
        developers_historically_active_or_inactive = ('Developers have been inactive for much of the past ten years. ')
        
        #If they've added to inventory, add a sentance about that
        if delivered_inventory > 0:
            developers_historically_active_or_inactive = developers_historically_active_or_inactive +  (
                                        'In fact, they have added just ' + 
                                        millify(delivered_inventory,'') + 
                                        ' '                 +
                                        unit_or_sqft        +
                                        ' to the '          +
                                        market_or_submarket + 
                                        ' over that time. '
                                                                                                        )
            #If they've demolished space, add a sentance about that
            if demolished_inventory > 0:
                developers_historically_active_or_inactive = developers_historically_active_or_inactive + ('Developers have also removed space for higher and better use, removing ' + 
                                            millify(demolished_inventory,'') + 
                                            ' ' +
                                            unit_or_sqft + 
                                            '. '
                                            )
        #If developers haven't added to inventory, we don't add that sentance
        else:

            #If they've demolished space, add a sentance about that
            if demolished_inventory > 0:
                developers_historically_active_or_inactive = developers_historically_active_or_inactive +  ('They have removed space for higher and better use, removing ' + 
                                            millify(demolished_inventory,'') + 
                                            ' ' +
                                            unit_or_sqft + 
                                            '. '
                                            )



    
    
    #Determine if the supply pipeline is active or not    
    if under_construction > 0:
        currently_active_or_inactive = 'Developers are currently active in the ' + market_or_submarket + ' with ' + millify(under_construction,'') + ' ' + unit_or_sqft + ', or the equivalent of ' + "{:,.1f}%".format(under_construction_share)   + ' of existing inventory, underway. '
        if yoy_submarket_vacancy_growth > 0:
            pipeline_vacancy_pressure    = 'The active pipeline will likely add upward pressure to vacancy rates in the near term.'
        elif yoy_submarket_vacancy_growth <= 0:
            pipeline_vacancy_pressure    = 'Demand in the ' + market_or_submarket + ' has outpaced new deliverables but could slow in the near term due to seasonal trends.'

    elif under_construction <= 0 :
        currently_active_or_inactive = 'Developers are not currently active in the ' + market_or_submarket + '. ' 
        pipeline_vacancy_pressure    = 'The empty pipeline will likely limit supply pressure on vacancies, boding well for fundamentals in the near term. '

    
    return(developers_historically_active_or_inactive + currently_active_or_inactive + pipeline_vacancy_pressure)












    #Old Code below:
    #Determine 10 year inventory growth   
    #if inventory_growth > 0 and under_construction > 0:
    #    inventory_expand_or_contract = 'In fact, over the past ten years, developers have added ' +  millify(inventory_growth,'') + ' '  + unit_or_sqft + ', expanding inventory by ' + inventory_growth_pct + '.'
    #
    #elif inventory_growth > 0 and under_construction <= 0 :
    #    inventory_expand_or_contract = 'However, over the past ten years, developers have added ' +  millify(inventory_growth,'') + ' '  + unit_or_sqft + ', expanding inventory by ' + inventory_growth_pct + '.'
    #
    #elif inventory_growth < 0:
    #    inventory_expand_or_contract = 'In fact, inventory has contracted ' +  millify(abs(inventory_growth),'')   + ' ' + unit_or_sqft + ', a ' + inventory_growth_pct + ' change. '
    # 
    #elif inventory_growth == 0:
    #    inventory_expand_or_contract = 'Over the past ten years, inventory levels have remained constant in the ' + market_or_submarket + '.'


    # #Determine qoq trends
    # if under_construction > 0 and previous_quarter_under_construction == 0:
    #     elevated_or_down_compared_to_previous_quarter = ' Developers have resumed activity after a brief pause. With ' +  "{:,.0f}".format(under_construction) + ' ' + unit_or_sqft + ' underway, inventory will expand by ' + "{:,.1f}%".format(under_construction_share) + '. While the pipeline is active, projects will not likely deliver over the 2nd half of the year, limiting supply pressure on vacancy rates.'
    
    # elif under_construction > 0 and previous_quarter_under_construction > 0 and under_construction == previous_quarter_under_construction:     
    #     elevated_or_down_compared_to_previous_quarter = """ Developers have remained active with the same level of construction underway in the previous quarter."""
    
    # elif under_construction > 0 and previous_quarter_under_construction > 0 and under_construction > previous_quarter_under_construction:     
    #     elevated_or_down_compared_to_previous_quarter = """ Developers have remained active with current construction levels surpassing the previous quarter's."""
    
    # elif under_construction > 0 and previous_quarter_under_construction > 0 and under_construction < previous_quarter_under_construction:     
    #     elevated_or_down_compared_to_previous_quarter = """ Developers have remained active, but current construction levels are below the previous quarter's."""

    # elif under_construction > 0 and previous_quarter_under_construction > 0:     
    #     elevated_or_down_compared_to_previous_quarter = ' Developers have remained active with ' +  "{:,.0f}".format(under_construction) + ' ' + unit_or_sqft + ' underway.'
    
    # elif under_construction <= 0 and previous_quarter_under_construction > 0:
    #      elevated_or_down_compared_to_previous_quarter = ' After activity in the previous quarter, developers have paused and nothing is currently underway. The empty pipeline will likely limit supply pressure on vacancies, boding well for fundamentals in the near term.'
    
    # elif under_construction == previous_quarter_under_construction == 0:
    #     elevated_or_down_compared_to_previous_quarter = ' Development activity has been steady with nothing underway in the current or previous quarter.' 

#Language for sales section
def CreateSaleLanguage(submarket_data_frame,market_data_frame,natioanl_data_frame,market_title,primary_market,sector,writeup_directory):
    #Pull writeup from the CoStar Html page if we have one saved
    CoStarWriteUp = PullCoStarWriteUp(section_names= ['Sales','Capital Markets'],writeup_directory = writeup_directory)
    if CoStarWriteUp != '':
        return(CoStarWriteUp)

    #Section 1: Begin making variables for the sales language that come from the data:     
    if sector == "Multifamily":
        unit_or_sqft                        = 'units'
        unit_or_sqft_singular               = 'unit'
        asset_value                         = submarket_data_frame['Asset Value/Unit'].iloc[-1]
        asset_value_change                  = submarket_data_frame['YoY Asset Value/Unit Growth'].iloc[-1]
        over_last_year_units                = submarket_data_frame['Sold Units'][-1:-5:-1].sum()

    else:
        unit_or_sqft                        = 'square feet'
        unit_or_sqft_singular               = 'SF'
        asset_value                         = submarket_data_frame['Asset Value/Sqft'].iloc[-1]
        asset_value_change                  = submarket_data_frame['YoY Asset Value/Sqft Growth'].iloc[-1]
        over_last_year_units                = submarket_data_frame['Sold Building SF'][-1:-5:-1].sum()

    current_sale_volume                     = submarket_data_frame['Total Sales Volume'].iloc[-1]
    current_transaction_count               = submarket_data_frame['Sales Volume Transactions'].iloc[-1]
    current_period                          = str(submarket_data_frame['Period'].iloc[-1])
    cap_rate                                = submarket_data_frame['Market Cap Rate'].iloc[-1]
    cap_rate_change                         = submarket_data_frame['YoY Market Cap Rate Growth'].iloc[-1]

    #Calculate the sale volume "over the last year" (last 4 quarters)
    over_last_year_sale_volume              = submarket_data_frame['Total Sales Volume'][-1:-5:-1].sum()
    over_last_year_transactions             = submarket_data_frame['Sales Volume Transactions'][-1:-5:-1].sum()
    
    #Collapse down the data to the annual total sales info
    data_frame_annual                    = submarket_data_frame
    data_frame_annual['n']               = 1
    data_frame_annual['n']               = 1
    data_frame_annual['n']               = 1
    data_frame_annual                    = data_frame_annual.groupby('Year').agg(sale_volume=('Total Sales Volume', 'sum'),
                                                transaction_count=('Sales Volume Transactions', 'sum'),
                                                n = ('n','sum')
                                                ).reset_index()
                                                
 
    try:
        data_frame_annual                   = data_frame_annual.loc[data_frame_annual['n'] == 4] #keep only years where we have 4 full quarters
        data_frame_annual                   = data_frame_annual.iloc[[-1,-2,-3]]          #keep the last 3 (full) years
        three_year_avg_sale_volume          = round(data_frame_annual['sale_volume'].mean())
        three_year_avg_transaction_count    = round(data_frame_annual['transaction_count'].mean())
    except:
        return('(DID NOT HAVE 3 FULL YEARS OF DATA)')


    
    #Section 2: Begin making varaiables that are conditional on the variables we have created in section 1

    #Determine if investors are typically active here
    #If theres at least 1 sale per quarter, active
    if submarket_data_frame['Sales Volume Transactions'].median() >= 1:
        investors_active_or_inactive = 'Buyers have shown steady interest and have been busily acquiring assets over the years. '
    else:
        investors_active_or_inactive = 'Buyers have not shown much interest in acquiring assets over the years. '


    #Describe change in asset values
    if asset_value_change > 0:
        asset_value_change_description = 'expanded'
    elif  asset_value_change < 0:
        asset_value_change_description = 'compressed'
    elif  asset_value_change == 0:
        asset_value_change_description = 'remained stable'
    else:
        asset_value_change_description = ''
    
    
    
    #Determine if market or submarket
    if submarket_data_frame.equals(market_data_frame):
        submarket_or_market           = 'Market'
    else:
        submarket_or_market           = 'Submarket'

    #Determine change in cap rate
    if cap_rate_change > 0:
        cap_rate_change_description  = 'expanded'
        cap_rate_change_description_to_or_at = 'to'

    elif cap_rate_change < 0:
        cap_rate_change_description  = 'compressed'
        cap_rate_change_description_to_or_at = 'to'

    else:
        cap_rate_change_description  = 'remained stable'
        cap_rate_change_description_to_or_at = 'at'


    if current_sale_volume > 0 :
        for_a_sale_volume_of = ' for a total sales volume of ' + millify(current_sale_volume,'$')                              
    else:
        for_a_sale_volume_of = ''

    if current_transaction_count > 1 or current_transaction_count == 0:
        sales_count_was_or_were      = 'were'
        sales_count_sale_or_sales    = 'sales'
        if current_transaction_count ==0:
            sales_count_sale_or_sales    = 'transactions'


        if current_transaction_count == 0:
            current_transaction_count = 'no'
    else:
        sales_count_was_or_were      = 'was'
        sales_count_sale_or_sales    = 'sale'
    
    
    if over_last_year_transactions > 1 or over_last_year_transactions == 0:
        over_last_year_transactions_or_transaction  = 'transactions'
        over_last_year_was_or_were                  = 'were'

    else:
        over_last_year_transactions_or_transaction  = 'transaction'
        over_last_year_was_or_were                  = 'was'

    if three_year_avg_transaction_count > 1 or three_year_avg_transaction_count == 0:
        three_year_avg_transaction_or_transactions  = 'transactions'

    else:
        three_year_avg_transaction_or_transactions  = 'transaction'



    #Section 3: Format variables
    cap_rate                         = "{:,.1f}%".format(cap_rate)

    #Format cap rate change variable
    if cap_rate_change == 0 :
            cap_rate_change          = ''
    else:
            cap_rate_change          = "{:,.0f} bps".format(abs(cap_rate_change))

    #format Asset value chagne
    if asset_value_change != 0:
        asset_value_change               = "{:,.0f}%".format(abs(asset_value_change))
        if asset_value_change == '0%':
            asset_value_change = 'slightly'
    else:
        asset_value_change = ''


    over_last_year_sale_volume       = millify(over_last_year_sale_volume,'$')
    over_last_year_transactions      = "{:,.0f}".format(over_last_year_transactions,'$')
    over_last_year_units             = millify(over_last_year_units,'')
    three_year_avg_sale_volume       = millify(three_year_avg_sale_volume,'$')
    three_year_avg_transaction_count = "{:,.0f}".format(three_year_avg_transaction_count)
    asset_value                      = "${:,.0f}".format(asset_value)

    current_sale_volume              = millify(current_sale_volume,'$')
    if current_transaction_count != 'no':
        current_transaction_count    = millify(current_transaction_count,'')


    #Section 4: Put together our variables into a pargaraph and return the sales language
    return(investors_active_or_inactive                      +
            'Going back three years, investors have closed, on average, ' +
            three_year_avg_transaction_count                 +
            ' '                                              +
            three_year_avg_transaction_or_transactions       +
            ' per year'                                      +
            ' with an annual average sales volume of '       +
            three_year_avg_sale_volume                       +
            '. '                                             +
           'Over the past year, there '                      +
            over_last_year_was_or_were                       +
           ' '                                               +
            over_last_year_transactions                      +
            ' closed '                                        + 
            over_last_year_transactions_or_transaction       +
            ' across '                                       +
            over_last_year_units                             +
            ' '                                              +
           unit_or_sqft                                      +
           ', representing '                                 +
           over_last_year_sale_volume                        +
           ' in dollar volume.'                              +
            ' In '                                           +
            current_period                                   +
            ', there '                                       + 
            sales_count_was_or_were                          +
             ' '                                             +
            current_transaction_count                        +
            ' '                                              +
            sales_count_sale_or_sales                        +
            "{recorded}".format(recorded = " recorded" if current_transaction_count == 'no'  else "") +
            for_a_sale_volume_of                             +
            '.'                                              +
            
            #Second paragraph
            ' Market pricing, based on the estimated price movement of all properties in the ' +
            submarket_or_market                              +
            ', sat at '                                      +
            asset_value                                      +
            '/'                                              +
            unit_or_sqft_singular                            +
            ' and has '                                     +
            asset_value_change_description                   +
            ' '                                              +
            asset_value_change                               + 
            ' over the past year, '                          +
            'while the market cap rate has '                 +
            cap_rate_change_description                      +
            ' '                                              +
             cap_rate_change                                 +
            ' over the past year '                           +
           cap_rate_change_description_to_or_at              +
           ' '                                               +
            cap_rate                                         +
            '. '                                             +
            ' Although capital markets have held up relatively well, uncertainty still remains. ' +
            ' Some investors may need to see signs of sustained economic growth before engaging. '
            )

        #Old unused code below
        # #determine change in investment volume over the last three years and the past year
        # if over_last_year_transactions > three_year_avg_transaction_count:
        #     investment_volume_change = 'Despite concerns over the pandemic, the number of closed transactions has picked up. In the past twelve months, investors have closed '
        # else:     
        #     investment_volume_change = 'With uncertainty surrounding the pandemic, transaction activity has slowed. '

#Language for outlook section
def CreateOutlookLanguage(submarket_data_frame,market_data_frame,natioanl_data_frame,market_title,primary_market,sector,writeup_directory):

    #Section 1: Begin making variables for the overview language that come from the data:
    if sector == "Multifamily":
        rent_var                = 'Market Effective Rent/Unit'
        rent_growth_var         = 'YoY Market Effective Rent/Unit Growth'
        qoq_rent_growth_var     = 'QoQ Market Effective Rent/Unit Growth'
        unit_or_sqft            = 'unit'
        net_absorption_var_name = 'Absorption Units'
        inventory_var_name      = 'Inventory Units'
        asset_value             = submarket_data_frame['Asset Value/Unit'].iloc[-1]
        asset_value_change      = submarket_data_frame['YoY Asset Value/Unit Growth'].iloc[-1]
        unit_or_sqft_singular   = 'unit'

    else:
        rent_var                = 'Market Rent/SF'
        rent_growth_var         = 'YoY Rent Growth'
        qoq_rent_growth_var     = 'QoQ Rent Growth'
        net_absorption_var_name = 'Net Absorption SF'
        inventory_var_name      = 'Inventory SF'
        unit_or_sqft            = 'SF'
        asset_value             = submarket_data_frame['Asset Value/Sqft'].iloc[-1]
        asset_value_change      = submarket_data_frame['YoY Asset Value/Sqft Growth'].iloc[-1]
        unit_or_sqft_singular   = 'SF'

    #Determine if market of submarket
    if submarket_data_frame.equals(market_data_frame):
        market_or_submarket = 'Market'
    else:
        market_or_submarket = 'Submarket'
    
    leasing_change                         = submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-1] -  submarket_data_frame[(net_absorption_var_name + ' 12 Mo')].iloc[-5]
    inventory_change                       = submarket_data_frame[inventory_var_name].iloc[-1] -  submarket_data_frame[inventory_var_name].iloc[-5]
    under_construction_share               = round(submarket_data_frame['Under Construction %'].iloc[-1],2)
    vacancy_change                         = submarket_data_frame['YoY Vacancy Growth'].iloc[-1]
    year_ago_cap_rate                      = submarket_data_frame['Market Cap Rate'].iloc[-5] 
    cap_rate                               = submarket_data_frame['Market Cap Rate'].iloc[-1] 
    avg_cap_rate                           = submarket_data_frame['Market Cap Rate'].mean() 
    cap_rate_change                        = submarket_data_frame['YoY Market Cap Rate Growth'].iloc[-1]
    current_period                         = str(submarket_data_frame['Period'].iloc[-1]) #Get most recent quarter
    submarket_yoy_growth                   =  submarket_data_frame[rent_growth_var].iloc[-1]
    submarket_qoq_growth                   =  submarket_data_frame[qoq_rent_growth_var].iloc[-1]
    submarket_previous_quarter_yoy_growth  =  submarket_data_frame[rent_growth_var].iloc[-2]
    
    #Section 2: Begin making varaiables that are conditional on the variables we have created in section 1
    #Describe YoY change in asset values
    if asset_value_change > 0:
        asset_value_change_description = 'expanded'
    elif asset_value_change < 0:
        asset_value_change_description = 'compressed'
    else:
        asset_value_change_description = 'remained steady'
    
    #Describe relationship between quarterly growth and annual rent growth
    if submarket_previous_quarter_yoy_growth > submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'contracting annual growth to'

    elif submarket_previous_quarter_yoy_growth < submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'pushing annual growth to'
    
    elif submarket_previous_quarter_yoy_growth == submarket_yoy_growth:
        qoq_pushing_or_contracting_annual_growth = 'keeping annual growth at'
    else:
        qoq_pushing_or_contracting_annual_growth = '[contracting/pushing] annual growth to'

    #Determine change in cap rate
    if cap_rate_change > 0:
        cap_rate_change_description  = 'expanded ' + "{:,.0f} bps".format(abs(cap_rate_change)) + ' to a rate of '            
    elif cap_rate_change < 0:
        cap_rate_change_description  = 'compressed ' + "{:,.0f} bps".format(abs(cap_rate_change)) + ' to a rate of '       
    else:
        cap_rate_change_description  = 'remained stable at '
    
    
    #Relationship betweeen current cap rate and the historical average
    #Cap Rate Below historical average
    if cap_rate < avg_cap_rate:
        
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'falling below'

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'remaining below'
        
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'falling below'

    #Cap Rate above historical average
    elif cap_rate > avg_cap_rate:
                
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'remaining above'
            

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'moving above'
            
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'moving above'
            
    #Cap equals  historical average
    elif  cap_rate == avg_cap_rate:
        #Cap Rate a year ago was above the historical average
        if year_ago_cap_rate > avg_cap_rate:
            cap_rate_above_below_average = 'falling to'

        #Cap Rate a year ago was below the historical average
        elif year_ago_cap_rate < avg_cap_rate:
            cap_rate_above_below_average = 'moving to'
        
        #Cap Rate a year ago was equal to the historical average
        elif year_ago_cap_rate == avg_cap_rate:
            cap_rate_above_below_average = 'remaining at'

   
    #Describe out change in fundamentals
    if submarket_yoy_growth >= 0     and vacancy_change <= 0: #if rent is growing (or flat) and vacancy is falling (or flat) we call fundamentals improving
        fundamentals_change = 'improving'
        values_likely_change = 'expand'
    elif submarket_yoy_growth < 0 and vacancy_change > 0 : #if rent is falling and vacancy is rising we call fundamentals softening
        fundamentals_change = 'softening'
        values_likely_change = 'compress'
    elif (submarket_yoy_growth > 0   and vacancy_change  > 0) or (submarket_yoy_growth < 0 and vacancy_change < 0 ) : #if rents are falling but vacancy is also falling OR vice versa, then mixed
        fundamentals_change = 'mixed'
        values_likely_change = 'stabilize'

    elif (submarket_yoy_growth == 0 and vacancy_change == 0): #no change in rent or vacancy
        fundamentals_change = 'stable'
        values_likely_change = 'stabilize'
        
    else:
        fundamentals_change = '[improving/softening/mixed/stable]'
        values_likely_change = '[expand/compress/stabilize]'

    
    #Forcast rent growth in 4th quarter
    if submarket_yoy_growth > 0:
       rent_future_path= 'accelerating'
    elif submarket_yoy_growth < 0:
        rent_future_path= 'compressing'
    elif submarket_yoy_growth == 0:
        rent_future_path= 'stabilizing'
    else:
        rent_future_path = '[stabilizing/accelerating/compressing]'
    
    #Forcast demand growth in 4th quarter
    if vacancy_change > 0:
       demand_future_path= 'remain muted'
    elif vacancy_change < 0:
        demand_future_path= 'continue to pick up'
    elif vacancy_change == 0:
        demand_future_path= 'stabilize'
    else:
        demand_future_path = '[continue to pick up/stabilize/remain muted]' 
    
      

    #Use the change in inventory, vacancy, and abosprtion over the past year to create a clause on market funamentals
    #Inventory expanded over past year
    if inventory_change > 0:

        #Vacancy increased
        if vacancy_change > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that demand has increased, but fallen short of rising inventory levels. Together, vacancy rates expanded over the past year. With vacancy rates expanding,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'a decrease in demand along with rising inventory levels. Together, vacancy rates have expanded over the past year. With vacancy rates expanding,'

               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'stagnant demand along with rising inventory levels. Together, vacancy rates have expanded over the past year. With vacancy rates expanding,'
            
            else:
                fundamentals_clause = ''
                
        #Vacancy decreased
        elif vacancy_change < 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'growing demand despite an increase in inventory. With demand outpacing new inventory, vacancy rates have compresssed over the past year. With vacancy rates compressing,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that despite falling demand and rising inventory levels, demolitions have aided the sector and vacancy rates have compresssed over the past year. With vacancy rates compressing,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'stable demand despite rising inventory levels. Together, vacancy rates have compresssed over the past year. With vacancy rates compressing,'

            else:
                fundamentals_clause = ''

        #Vacancy flat
        elif vacancy_change == 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that despite new supply additions over the past year, demand has kept pace and vacancy rates have managed to stay in line with last years vacancy rate. With stable vacancy rates,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'stable demand despite rising inventory levels. Together, vacancy rates have remained stable over the past year. With vacancy rates stable,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'that despite rising inventory levels and with no change in demand, vacancy rates have managed to stay in line with last years vacancy rate. With stable vacancy rates,'

            else:
                fundamentals_clause = ''
        else:
            fundamentals_clause = ''

    #Inventory contracted over the past year
    elif inventory_change < 0:

        #Vacancy increased
        if vacancy_change > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that despite a decrease in inventory levels and a recent rise in demand, vacancy rates have expanded over the past year. With vacancy rates expanding,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that despite a decrease in inventory levels, demand has fallen too, expanding vacancy rates over the past year. With vacancy rates expanding,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'that despite a decrease in inventory levels, demand remained flat, expanding vacancy rates over the past year. With vacancy rates expanding,'

            else:
                fundamentals_clause = ''

        #Vacancy decreased
        elif vacancy_change < 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'positive trends. Inventory levels have contracted and leasing activity picked up, compressing vacancy rates over the past year. With vacancy rates compressing,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that despite a decrease in absorption levels, the market_or_submarket has been aided by a decrease in inventory, compressing vacancy rates over the past year. With vacancy rates compressing,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'positive trends. Despite flat absorption over the past year, inventory has decreased, allowing for vacancy rate compression. With vacancy rates compressing,'

            else:
                fundamentals_clause = ''
        #Vacancy flat
        elif vacancy_change == 0:
            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'Despite falling inventory and growing demand,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'Despite falling inventory levels, with demand falling,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'Despite falling inventory levels and no change in demand,'
            
            else:
                fundamentals_clause = ''
        else:
            fundamentals_clause = ''

    #Inventory flat over the past year
    elif inventory_change == 0:

        #Vacancy increased
        if vacancy_change > 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that despite a stable inventory count and rising demand, vacancy rates have increased. With vacancy rates expanding,'


            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that despite no new additions to the inventory, demand has fallen, expanding vacancy rates. With vacancy rates expanding,'

               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'that despite no change in inventory or demand, vacancy rates have expanded over the past year. With vacancy rates expanding,'

            else:
                fundamentals_clause = ''

        #Vacancy decreased
        elif vacancy_change < 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that demand picked up in the absence of inventory growth, compressing vacancy rates over the past year. With vacancy rates compressing,'

            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that although demand has declined, vacancy rates have been aided by an absence of inventory growth, allowing for vacancy rate compression. With vacancy rates compressing,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'that despite flat demand and no change in inventory, vacancy rates have compressed over the past year. With vacancy rates compressing,'

            else:
                fundamentals_clause = ''            
        #Vacancy flat
        elif vacancy_change == 0:

            #12m net absorption grew over past year
            if leasing_change > 0:
                fundamentals_clause = 'that demand picked up in the absence of inventory growth, but vacancy rates have been stable over the past year. With stable vacancy rates,'


            #12m net absorption declined over past year
            elif  leasing_change < 0:
                fundamentals_clause = 'that although demand has declined, vacancy rates have been aided by an absence of inventory growth, allowing for steady vacancy rates. With stable vacancy rates,'
               
            #12m net absorption flat over past year
            elif leasing_change == 0:
                fundamentals_clause = 'that despite flat demand and no change in inventory, vacancy rates have remained steady over the past year. With stable vacancy rates,'
            
            else:
                fundamentals_clause = ''
        else:
            fundamentals_clause = ''
    
    else:
        fundamentals_clause = ''

    #Sector Specific language
    if sector == "Multifamily":
        sector_specific_outlook_language=('Strong economic growth and a drastically improving public health situation helped boost multifamily fundamentals over the first three quarters of 2021. With demand and rent growth indicators surging, investors have regained confidence in the sector, and sales volume has returned to more normal levels over the past few quarters. Still, a few headwinds exist that could put upward pressure on vacancies over the next few quarters. The ' + market_or_submarket + ' still faces a robust near-term supply pipeline, and those units will deliver amid a potential slowdown in demand due to seasonality and the fading effects of fiscal stimulus that has helped thousands of people pay rent. Furthermore, single-family starts have ramped up, and the increase in new for-sale housing could draw higher-income renters away from luxury properties.')
    
    elif sector == "Office":
        sector_specific_outlook_language=('The first three quarters of 2021 remained in line with pandemic-era trends in terms of office market performance. Although leasing activity has picked up slightly, it remained rather subdued. Many tenants continue to downsize and adopt hybrid work models, limiting demand and rent growth. Investment volume remains subdued, but investors are looking at alternatives such as the medical office sector or single-tenant assets with sticky tenants and lengthy leases in place. Looking ahead over the next few quarters, supply additions will be met with muted demand, limiting improvement in rents and values.')

    elif sector == "Retail":
        sector_specific_outlook_language=('The new year has delivered encouraging news for the retail sector: Retail sales activity surged as the year commenced, vaccine rollouts are supporting strong consumer confidence metrics, and leasing activity among many tenant segments remains strong.  Such positive news does not, however, overshadow the complexity and nuance that the sector possesses. Indeed, a tale of two recoveries continues to unfold, and property performance continues to vary significantly by subtype, location, class, and tenant composition. Even with the vaccines, it is probable retailers will continue to face turbulence in the coming quarters. Those effects will likely linger for the foreseeable future, impacting demand, rent growth, and the capital markets in the process.')
    
    elif sector == "Industrial":
        sector_specific_outlook_language=("""The new year has brought much needed support to the nation's economy and to its consumers, who continue to buy record amounts of goods online. In response, industrial users continue to seek more warehouse space closer to the consumer as they evolve their supply chains to meet the demand for fast delivery times. Industrial's rent growth prospects continue to lead across sectors, as well, with both retail and office posting rent declines as multifamily gradually regains momentum after plateauing throughout much of 2020. Still, following the national theme, most markets are set to experience a deceleration in rent growth. With such strength prevailing throughout industrial's operating environment, and with other sectors and asset classes registering more volatility and relatively weaker performance, investors continue to aggressively pursue industrial acquisitions. Looking ahead over the next few quarters, demand from consumers, tenants, and investors will continue driving growth in fundamentals.""")





        
    #Section 3: Begin Formatting variables


    #Section 4: Begin putting sentences together with our variables
    large_supply_pipeline_threshold = 5
    if vacancy_change > 0 and under_construction_share == 0:
        pipeline_sentence = ('However, an empty supply pipeline could allow for vacancy to stabilize. ' )
    elif vacancy_change < 0 and under_construction_share >= large_supply_pipeline_threshold:
        pipeline_sentence = ('However, a large supply pipeline could put upward pressure on vacancy rates. ' )
    elif vacancy_change > 0 and under_construction_share >= large_supply_pipeline_threshold:
        pipeline_sentence = ('Furthermore, a large supply pipeline could put futher upward pressure on vacancy rates. ' )
    elif vacancy_change < 0 and under_construction_share == 0:
        pipeline_sentence = ('Furthermore, an empty supply pipeline should allow for further vacancy rate compression. ' )
    else:
        pipeline_sentence = ('')



    capital_markets_summary = (
                'With fundamentals '              +
                fundamentals_change                +
                 ' for '                           +
                 sector.lower()                    +
                 ' properties in the '             +
                 market_or_submarket               +
                ', values have '                   +
                asset_value_change_description     +
                ' over the past year to '          +
                "${:,.0f}/". format(asset_value)   + 
                unit_or_sqft_singular              +
                ' and cap rates have '             +
                cap_rate_change_description        +
                "{:,.1f}%".format(cap_rate)        +
                ', '                               +
                cap_rate_above_below_average       +
                ' the long-term average.'
                                    )

    
    general_outlook_language = ('Current fundamentals for ' +
                            sector                          +    
                            ' properties in the '           +
                            market_or_submarket             +
                            ' indicate '                     + 
                            fundamentals_clause             +
                            ' quarterly growth in ' + current_period + ' reached ' +  "{:,.1f}%".format(submarket_qoq_growth) + ', ' + qoq_pushing_or_contracting_annual_growth + ' ' +  "{:,.1f}%".format(submarket_yoy_growth) + '. ' +
                            capital_markets_summary)
                            
                            
    outlook_conclusion_language =  (
                            'Looking ahead to the ' +
                            'near-term' + 
                            ', it is likely that demand will ' +
                            demand_future_path +
                            ' with rents ' +
                            rent_future_path +
                            ' further. ' +
                            pipeline_sentence +
                            'With fundamentals ' +
                            fundamentals_change +
                            ', values will likely ' +
                            values_likely_change+
                            '.'
                            )


    #Section 5: Combine sentences and return the conclusion langage
    return(sector_specific_outlook_language  + '\n' + '\n' + general_outlook_language + '\n' + '\n' + outlook_conclusion_language)







        #Old Code Below:
                            # 'general [stability/instability]' +
                            # ' in demand while the count of new deliverables have been ' + 
                            # '[expanding/steady/limited/absent]' +
                            # '. Together, vacancy rates have ' +
                            # '[managed to remain stable/expanded considerably/compressed]'  +
                            # ' over the course of the pandemic. ' +
                            # 'Rents responded by ' + 
                            # '[remaining stable/expanding/softening]' +
                            # '. The general ' +
                            # '[stability/instability/acceleration/deceleration]' +
                            # ' in fundamentals have helped improve the capital market, resulting in ' +
                            # '[stable/accelerating/decelerating]' +
                            # ' growth in property values across the sector. '
                            # )







