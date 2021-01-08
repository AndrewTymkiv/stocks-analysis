# VBA Challenge
---
## Project Overview
### In this project, we help Steve develop an efficient analysis of the stock market based on total volume and return in Excel VBA. The purpose of the VBA script was to create an efficient method of analyzing multiple stocks in the stock market that loops through all the data cleanly and quickly. We initially developed the code to run for the user's choice of either 2017 or 2018, and we refactored this code to run more efficiently.
---
## Results
### 1. Create a ticker index and 3 output arrays
#### First I had to create a tickerIndex variable and set it equal to zero so it could be used across the 4 different arrays in the code. Next, I had to add 3 arrays in addition to the tickers array; tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
   
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

### 2. Create for loops for the tickerVolumes and to loop through all the rows in the spreadsheet
#### A For loop labelled with i was created to set the tickerVolumes to 0, then another For loop was created to loop through all the rows in the spreadsheet, labelled by j. 

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount

### 3. Read through tickerVolumes and returns for each ticker
#### Within the j For loop, the code pulls each stock's volume by increasing the tickerVolume using the tickerIndex for the current ticker. Then, it finds where the first and last rows are for each specific ticker once again using the tickerIndex.

        '3a) Increase volume for current ticker
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
        'End If
        End If
        
    
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
         End If
         
### 4. Loop through arrays to put ticker, tickerVolume, and return onto the worksheet
#### Another For loop, k, was created and the All Stocks Analysis worksheet was activated so I could output the data correctly on the worksheet.

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
     For k = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumes(k)
        Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
        
        
    Next k
    
### Times
#### The MsgBox part of the code tells me how quickly the analysis ran for each year and the results show that the refactored code runs more efficiently because the time is quick.

![VBA_Challenge_2017](https://github.com/AndrewTymkiv/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


