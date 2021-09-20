# stock-analysis
This repository will analyze several "green" companies to invest in.
###Analysis Stock Outcomes of Several Green Companies

##Overview of Project

This project analyzed the stock daily volumes and rate of return for several different green companies over the span of two years. The project was designed to help Steve's parents determine which would be the most profitable business to invest in. Steve's parents had decided to invest all their money in DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. The project will help Steve analyze the stock performance for several different companies, including DAQO, to determine which company is the most successful.

##Results
#2017 Results

The table below shows the analysis of the group of stocks for 2017.



#2018 Results

The table below shows the analysis of the group of stocks for 2018.



##Summary

1. The initial project began by compiling several different pieces of code to come to the final results. The code included:

	a. Setting up an array for the tickers
	b. Calculating the number of rows
	c. Looping through the rows to determine the starting and ending prices, the total daily volume and the yearly return for each sticker
	d. Outputting the results for each ticker, then looping through the data again for the next ticker
	e. Formatting the results
	f. Configuring a timer for the run time of the code
	g. Adding code for year input
	h. Adding a button to calculate the data and output it to the sheet as well as a button to clear the sheet

2. The refactored project altered the code in the initial project by using an array rather than variables to collect data for the total daily volume and the yearly return. The VBA used the array to loop through the data so that the code did not run all the way through the code multiple times as in the original project. Refactoring code saved over one second for each year:

	a. 2017 original code, length of run time: 		1.433594 seconds
	b. 2017 refactored code, length of run time:	0.03203125 seconds
	c. 2018 original code, length of run time: 		1.46875 seconds
	d. 2018 refactored code, length of run time: 	0.3046875 seconds

Although it may not seem like refactoring the code saved much time, the dataset was actually fairly small as it only included data for 12 companies. If the analysis were done on a large dataset, the initial VBA code could cause performance issues. The refactored code performed much more quickly.