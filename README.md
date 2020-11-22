# Green Stock Analysis

## Project Overview

My buddy Steve asked me to make an analysis on stocks from 2017 and 2018 to see what would be worth investing in. I decided to help build this analysis through using VBA in Excel to find what were the various stock's total daily volume and annual returns. By analyzing what the total volume from the beginning of the year compared to the end, I will be able to determine if a stop had performed well or not that year and will be able to assist Steve's parents in what would be some safe options.

### Purpose

The purpose of this project was to make an efficient way to view and analyze multiple stocks to help assist Steve’s parents in regards to stocks by utilizing VBA tools. With VBA we are going to run through 12 different stocks over the span of 2 years to identify how each stock did per year. This project will look to test how efficient my refactoring analysis code  is.

## Results

### Refactoring the Code

In my code I will be testing nested loops to run through each of the various stock tickers. I created 4 arrays:tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The main 3 arrays will be used by the variable called tickerIndex to flow through the stock ids.

#### Refactored Code

'3) Initialize array of all tickers
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '4b) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '5a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '5b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    '6a) Initialize ticker volumes to zero
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    '6b) loop over all the rows
    
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
            
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i

#### Original Code

'Initialize array of all tickers
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
    'what do they mean by ticker index? counts more like a row number
    'you know which row you are
    
    tickerIndex = 0
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.

    For stocks_i = 0 To 11
    tickerVolumes(stocks_i) = 0
    
    Next stocks_i
            ''2b) Loop over all the rows in the spreadsheet.
            For Rows_j = 2 To RowCount
    
           
                    '3a) Increase volume for current ticker
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(Rows_j, 8).Value
                                        
                    '3b) Check if the current row is the first row with the selected tickerIndex.
                     If Cells(Rows_j - 1, 1).Value <> tickers(tickerIndex) Then
                        tickerStartingPrices(tickerIndex) = Cells(Rows_j, 6).Value
                    End If
                    
                    '3c) check if the current row is the last row with the selected ticker
                     'If the next row’s ticker doesn’t match, increase the tickerIndex.
                    If Cells(Rows_j + 1, 1).Value <> tickers(tickerIndex) Then
                        tickerEndingPrices(tickerIndex) = Cells(Rows_j, 6).Value
                                    
                        '3d) Increase the tickerIndex.
                        tickerIndex = tickerIndex + 1
                        
                    End If
                    
            Next Rows_j
        
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For stocks_i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + stocks_i, 1).Value = tickers(stocks_i)
        Cells(4 + stocks_i, 2).Value = tickerVolumes(stocks_i)
        Cells(4 + stocks_i, 3).Value = tickerEndingPrices(stocks_i) / tickerStartingPrices(stocks_i) - 1
        
        
    Next stocks_i


