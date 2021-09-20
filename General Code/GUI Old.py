

###Select one of the Markets
##OPTIONS = unique_markets_list 
##
##
##master = Tk()
##
##variable = StringVar(master)
##variable.set(OPTIONS[0]) # default value
##
##w = OptionMenu(master, variable, *OPTIONS)
##w.pack()
##
###Keep if in desired market
##def SelectMarket():
##    #print ("value is:" + variable.get())
##    desired_market =  variable.get()
##    df = df1[df1['Geography Name'] == desired_market]
## 
##    price_sum_stat = df['Average Sale Price'].describe()
##    print(price_sum_stat)
##    mean_sale_price = price_sum_stat['mean']
##    print(mean_sale_price)
##    #Line Graph
##    
##
##    #Create Scatter  Plot of Price in each year
##    scatter_plot = df.plot.scatter(x='Period',y='Average Sale Price',title=desired_market)
##    figure = scatter_plot.get_figure()
##    figure.savefig(os.path.join(output_location,'output.png'))
##    document.add_picture(os.path.join(output_location,'output.png'))
##    
##    print('The mean valuation in ' + desired_market + " is " + str(mean_sale_price))
##    document.save(report_path)
##
##button = Button(master, text="OK", command=SelectMarket)
##button.pack()
##mainloop()
##
##
##
###Save Report


