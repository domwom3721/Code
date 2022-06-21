# Research-Report-Automation
3 scripts (1 for each of our 3 main report products) automate the production of our report documents. 

    Market Code - This code proceesses exported data from CoStar.com on real estate markets, loops through all the markets and submarkets, creates a directory for each, creates multiple graphs to dipslay the data (png images), and finally makes a word documents for the market/submarket where we insert langauge and the png images into the report document.
    
    Area Code   - This code takes 5 digit FIPS codes (for US Counties) as inputs from the user, loops through the given list of FIPS codes and prepares reports on the County economy. Using the FIPS code the script does the following:
                                                            1.) Identify the name of the county
                                                            2.) Identify if the county is in a Metro area (and the name/CBSA code of it if so)
                                                            3.) Idenfity the state of the county
                                                            4.) Goes to multiple pulbic data APIs and pulls in county/msa/state/national economic data
                                                            5.) Goes to Google Maps and takes a screen shot of the map of the County
                                                            6.) Uses this data to create png imgage graphs and produce natural lanuguage
                                                            7.) Creates a word document report with the map, graphs and language inserted
                                                            8.) Saves the report in the correct directory for the County within a state directory

Neighbrohood Code - This script takes Census Place (Cities/CDPs) FIPS codes as inputs and uses Census API to pull in public data, create graphs, and a word document report for the city. It will also pull data for a larger comparison area (eg: County) and display that in the graphs.

General Code - Scripts I have written to help with some basic tasks while cleaning up folders, and getting project started. None of these scripts are currently in use or part of the project on a regular basis.
