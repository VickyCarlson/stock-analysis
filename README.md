# VBA of Wall Street

## Overview of Project

This project analyzes the stock performance of twelve different "green" companies. By looking at the stock daily volumes and rate of return for several different green companies over the span of two years, it is possible to determine which company would be the best to invest in. The project was designed to help a friend (Steve) assist his parents in determining which would be the most profitable business for them to allocate their resources to. Steve's parents had decided to invest all their money in DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. Before his parents invested their money in DAQO, Steve wanted to analyze other "green" companies to ensure that their investment was worthwhile. The project will help Steve analyze the stock performance for several different companies, including DAQO, to determine which company is the most successful.

## Stock Analysis Results

### 2017 Results

Looking at the results for the 2017 stock performance, it may appear that Steve's parents have picked a good stock to invest in. Their choice for investment was DQ (DAQO New Energy Corporation), and it had the highest overall return of all companies in 2017, with a 199.4% return. That said, the total daily volume of DQ was the *lowest* of all companies, with only 35,796,200 shares traded, while the highest volume of 782,187,000 was for SPWR. The difference between these two companies is 746,390,800. *Investopedia* notes that investors use trading volume to signal whether a trend will continue. Since the trading volume is low in 2017, there is a possibility that this may not be a good stock to invest in. The image below shows the outcome for all 2017 stocks using the refactored code.

![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

There was a great difference in the processing time between the refactored analysis (see image above) and the original analysis, which did not include an array. The processing time for the refactored code for 2017 was 0.3554688 seconds, while the original code took over four times as long, taking 1.558594 seconds to run. See below for the original analysis.

![VBA_Challenge_2017_Original_Code.png](/Resources/VBA_Challenge_2017_Original_Code.png)

The original code looped through each ticker, then looped through each row to find the starting and ending prices as well as the volume. It would then output the results of that ticker to the spreadsheet before starting on the next ticker. See below for code for the original project.

	*'Loop through the tickers
	For i = 0 To 11

	    ticker = tickers(i)
	    totalVolume = 0

	    'Activate worksheet
	    Sheets(yearValue).Activate

	    'Loop through rows in the data
	    For j = 2 To RowCount

	        'Find the total volume for the current ticker
	        If Cells(j, 1).Value = ticker Then

	        'increase totalVolume by the value in the current row
	        totalVolume = totalVolume + Cells(j, 8).Value

	        End If

	        'Find starting price per ticker
	        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

	            startingPrice = Cells(j, 6).Value

	        End If

	        'Find ending price
	        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

	            endingPrice = Cells(j, 6).Value

	        End If

	    Next j


	'Output the data for the current ticker
	Worksheets("All Stocks Analysis").Activate
	Cells(4 + i, 1).Value = ticker
	Cells(4 + i, 2).Value = totalVolume
	Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

	Next i*

The refactored code used arrays to perform the loops rather than lists, which dramatically improved the performance. Below is the code used to run the arrays in the refactored VBA.

	*'Initialize array of all tickers
	    Dim tickers(12) As String
	    
	        tickers(0) = "AY"
	        tickers(1) = "CSIQ"
	        tickers(2) = "DQ"
	        tickers(3) = "ENPH"
	        tickers(4) = "FSLR"
	        tickers(5) = "HASI"
	        tickers(6) = "JKS"
	        tickers(7) = "RUN"
	        tickers(8) = "SEDG"
	        tickers(9) = "SPWR"
	        tickers(10) = "TERP"
	        tickers(11) = "VSLR"
	    
	    'Activate data worksheet
	    Worksheets(yearValue).Activate
	    
	    'Get the number of rows to loop over
	    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
	    
	    '1a) Create a ticker Index
	    Dim tickerIndex As Single
	    tickerIndex = 0

	    '1b) Create three output arrays
	    Dim tickerVolumes(12) As Long
	    Dim tickerStartingPrices(12) As Single
	    Dim tickerEndingPrices(12) As Single
	    
	    ''2a) Create a for loop to initialize the tickerVolumes to zero.
	    For tickerIndex = 0 To 11
	        tickerVolumes(tickerIndex) = 0
	        
	    ''2b) Loop over all the rows in the spreadsheet.
	    For i = 2 To RowCount
	    
	        '3a) Increase volume for current ticker
	    If Cells(i, 1).Value = tickers(tickerIndex) Then
	        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
	    End If
	        
	        '3b) Check if the current row is the first row with the selected tickerIndex.
	        'If  Then
	        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
	    
	            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
	       
	        'End If
	         End If
	     
	        
	        '3c) check if the current row is the last row with the selected ticker
	         'If the next row’s ticker doesn’t match, increase the tickerIndex.
	        'If  Then
	         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
	            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
	            
	        '3d Increase the tickerIndex.
	            tickerIndex = tickerIndex + 1
	            
	        'End If
	        End If
	    Next i
	    
	    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
	        For i = 0 To 11
	            
	            Worksheets("All Stocks Analysis").Activate
	            Cells(4 + i, 1).Value = tickers(i)
	            Cells(4 + i, 2).Value = tickerVolumes(i)
	            Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
	        
	        Next i

	Next tickerIndex*

### 2018 Results

When analyzing stock performance, it is recommended that trends be analyzed over time. While the DQ stock did well in 2017, it did not perform well in 2018. While the total daily volume was up from 2017 (107,873,900 in 2018 compared to 35,796,200 in 2017), the return was down substantially. The rate of return went from 199.4% in 2017 to -62.6% in 2018. The total volume includes both purchases and sales of stocks, so there is a possibility that the volume was high because investors were selling their shares. The image below shows the outcome for all 2018 stocks.

![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

Again, there was a large difference in the processing time between the refactored code (see image above) and the original analysis, which did not include an array. The processing time for the refactored code for 2018 was 0.3789063 seconds, while the original code took 1.53125 seconds to run. See below for the time for the original analysis.

![VBA_Challenge_2018_Original_Code.png](/Resources/VBA_Challenge_2018_Original_Code.png)


## Summary

By analyzing two years of stock, Steve's parents should probably find a different stock to invest in than DQ. ENPH seemed to have sustainable returns and volumes. Its volumes were 221,772,100 in 2017 and increased to 607,473,500 in 2018. The return for ENPH was 129.5% in 2017, the third highest of all companies, and achieved the second highest results of all companies in 2018 at 81.9%. Only two companies achieved positive returns in 2018.

The refactored project altered the code from the initial project by using an array rather than a list. The code using the array performed over four times faster. 

	a. 2017 original code, length of run time: 		1.558594 seconds
	b. 2017 refactored code, length of run time:	0.3554688 seconds
	c. 2018 original code, length of run time: 		1.53125 seconds
	d. 2018 refactored code, length of run time: 	0.3789063 seconds

Although it may not seem like refactoring the code saved much time, the dataset was actually fairly small as it only included data for 12 companies. If the analysis were done on a large dataset, the initial VBA code could save significant time. The refactored code using the array structure seems like one solution for processing issues.

There are, however, disadvantages to using arrays. Because developers use lists so frequently, they are easy to use in code and to debug. Using the array format was more difficult initially to create. The array must be declared, while a list can be created simply by adding items into brackets. Given the increase in performance, however, using an array may be the best solution with large datasets.