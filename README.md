# Green Stock Analysis 

## Overview of Project

Analysis of "Green Stocks" using Visual Basic on Excel.  

### Purpose

Run scripts to check total volume and return values of *"Green Stocks"* for a specified year. Compared runtime of code before and after refactoring.

## Analysis and Challenges
###Original code
Original subscripts for analysis of stocks ran nested loops. The outer loop runs through each ticker of the array, established as "tickers(i)", and the inner loop runs through the rows of data of the reference sheet. 

>...
	For i = 0 To 11 
	    ticker = tickers(i)
            totalVolume = 0
            '5) loop through rows in the data
                    
            Sheets(yearValue).Activate
            
            For j = 2 To RowCount
                
                '5a) Get totalVolume for current ticker
                
                If Cells(j, 1).Value = ticker Then
                    
                    totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
                '5b) get starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    
                    startingPrice = Cells(j, 6).Value
                
                End If
                    
                '5c) get ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    
                    endingPrice = Cells(j, 6).Value
                    
                End If
                    
            Next j
  ...

These subscripts run around 1.79 seconds for All Stocks in 2017 and around 1.98 seconds for All Stocks in 2018

![2017_All_Stocks_Analysis](Resources/2017_All_Stocks_Analysis.png)

![2018_All_Stocks_Analysis](Resources/2018_All_Stocks_Analysis.png)

### Refactored code

The refactored code contains 3 arrays that loop once through the code.  A **tickerIndex** was added into the newly created arrays to simplify the loop:
>
 	For i = 2 To RowCount
        
            tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
        
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
         
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          
            tickerIndex = tickerIndex + 1
            
            End If
    	
	Next i

The refactored code ran in about .21 and .01 seconds for All Stocks in 2017 and 2018, respectively.

![2017_Refactored_All_Stocks_Analysis](Resources/2017_Refactored_All_Stocks_Analysis.png)

![2018_Refactored_All_Stocks_Analysis](Resources/2018_Refactored_All_Stocks_Analysis.png)

### Challenges and Difficulties Encountered

## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?

- What can you conclude about the Outcomes based on Goals?

- What are some limitations of this dataset?

- What are some other possible tables and/or graphs that we could create?
