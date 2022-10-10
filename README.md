# VBA Challenge

## Overview of Project

This project analyses green energy stocks to aid understand their annual return rate. Steve(client) needs to develop a tool that shows by year how well green energy stocks performed in the market.


### Purpose
This project analyses green energy stocks to aid understand their daily volume and annual return rate, so that Steve's parents can diversify th investement risk based on historical information. Furthermore, insights in refactoring code will be discussed throughout the document.


## Analysis 

Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

In the analysis two main concerns will be addressed: 
- Comparison of stock performance in 2017 vs 2018. 
- Execution time and results in 2018 all stocks analysis with original and refactored code. 

### Stock Performance

In the table below, daily volume and return rates are shown for the year 2017.

NOTE: Positive rates are highlighted in green and negative rates in red. 

![Captura de Pantalla 2022-10-05 a la(s) 23 04 17](https://user-images.githubusercontent.com/114015620/194211499-38821e26-e309-4239-b6c0-341345dfe063.png)

  - This year DQ, SEDG, ENPH had the highest return rates of all stocks. 
  - TERP was the only stock with a negative return rate. 

For the year 2018, the stock performance was significantly different. 

![Captura de Pantalla 2022-10-05 a la(s) 23 11 49](https://user-images.githubusercontent.com/114015620/194212241-067f3dfb-152c-4a2a-9216-3c4e8ae1b242.png)

  - The stocks with higher return rate were RUN and ENPH. Actually, these two are the only stocks with positive return rate.
  - ENPH remains in the top three in return rate for 2018.
  - DQ had the worst performance of all. Contrary to its performance in 2017, where DQ was top 1. Falling more than 200%.

### Execution Time

The screenshot below, shows the results with the original code and refactored. The results are the same, however, the code performance is different.

![Captura de Pantalla 2022-10-05 a la(s) 23 30 54](https://user-images.githubusercontent.com/114015620/194214419-6c1676f2-362c-404e-a98b-a4ba88db4064.png)

In the original code, the time to perform the entire operations was of 1.1796 seconds. 

![Captura de Pantalla 2022-10-05 a la(s) 20 29 02](https://user-images.githubusercontent.com/114015620/194213221-f5ceea3b-2ead-42dd-a9e9-f64001f83714.png)

The time for the refactored code was 1.0625

![Captura de Pantalla 2022-10-05 a la(s) 20 28 48](https://user-images.githubusercontent.com/114015620/194213270-92f63a66-4159-493f-830f-2a06f1f0bcb6.png)

This means the code was optimized by 10% in operation time. In both codes the result is the same but using less computations.

Let's show some examples of refractored script:

The first main difference was that in the refactored script, arrays were used to store data of daily volume and prices. 
The code is as follows:

    Dim tickerVolumes(12) As Long

    Dim tickerStartingprices(12) As Single

    Dim tickerEndingprices(12) As Single
    
 In order to access the arrays, tickerIndex variable needed to be introduced. 
 This way we could get ticker names and use them as index to fill the three arrays created (tickerVolumes, tickerStartingprices, tickerEndingprices.
 
 A for loop was created to initialize the tickerVolumes to zero:

    
    For tickerIndex = 0 To 11
    
    'Get names of tickers
    ticker = tickers(tickerIndex)
    tickerVolumes(tickerIndex) = 0


    
After this, another loop was created just like in the original code. The difference lays in the arguments within the if-conditions. Ifs were used to identify tickerStartingprices and tickersEndingprices.

    
    For tickerIndex = 0 To 11
    
    'Get names of tickers
    ticker = tickers(tickerIndex)
    tickerVolumes(tickerIndex) = 0

        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
            If Cells(i, 1).Value = ticker Then
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
    
        tickerStartingprices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            
            '3d Increase the tickerIndex.
            
            
        'End If

        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

        tickerEndingprices(tickerIndex) = Cells(i, 6).Value
    
        End If
    
    Next i
    
    Next tickerIndex

Through this code the arrays were fed, therefore, it will loop through the arrays to output the Ticker, Total Daily Volume, and Return.
 
 
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysisRefactored").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingprices(i) / tickerStartingprices(i) - 1
      
        
    Next i

Basically, the main change from original code and refactored code was the use of arrays to manage information. T

## Summary

### Advantages and disadvantages of refactoring code

The project a refactored code that reached to the same results with a different way to deliver them.  Through arrays computation time was reduced by 10%.
This shows that improvement is always possible. Although, 


### Advantages and disadvantages of refactoring code of the original and refactored VBA script

