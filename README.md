# StockAnalysis

## Overview of Project

Steve has asked me to write a code to analyze stock data he provided me. Using VBA, I wrote a set of macros that allowed Steve to compare the performance of various stocks based on the data provided. Steve wants to be able to perform a similar analysis on all stock market data. To accommodate so much data I am refactoring my original code to see if I can make the macros run more quickly.

## Results

### Original Code

My Original code utilized 1 array to collect data for the 12 stocks provided in Steve's data set. The code sets an array of stock types, with each stock being an index in the array. It then loops through all rows of data, asks if the stock name matches the index in the current loop. The volume for each stock is added together, it finds the start price of the stock and the end price to find the total return for each stock. This code was successful in collecting comparative data, but consistently has a run time over 1 second, see code below:

````
```

Sub AllStocksAnalysis()

    'Tracking code run time
    Dim startTime As Single
    Dim endTime As Single

    Worksheets("All Stocks Analysis").Activate
    
    'Gives option to run code on different worksheets
    yearValue = InputBox("What year would you like to run the 	analysis on?")

    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Stock"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2.  Initialize an array of all tickers.
    Dim tickers(12) As String
        
    '2.  Initialize an array of all tickers.
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
    
    'a.  Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    'b.  Activate the data worksheet.
    Sheets(yearValue).Activate
    
    startTime = Timer
    
    'c.  Find the number of rows to loop over.
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4.  Loop through the tickers.
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
    
        '5.  Loop through rows in the data.
        Sheets(yearValue).Activate
        
        For j = 2 To rowEnd
        
            'a.  Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'b.  Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'c.  Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If

            
        Next j
        
    
        '6.  Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
    'Formatting Headers
    Worksheets("All Stocks Analysis").Activate
    With Range("A3:C3").Font
    .Bold = True
    .Size = 14
    .Underline = xlUnderlineStyleDouble
    End With
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Formatting results
    Range("B4:B15").NumberFormat = "#,###,###"
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit
    Columns("C").AutoFit
       
    'Format returns analysis to have color Red for loss, green for gain
    rowStart = 4
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = rowStart To rowEnd
        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            
        Else: Cells(i, 3).Interior.Color = xlNone
        
        End If
        
    Next i
    
    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year " & (yearValue)
        
End Sub

```
````

The run time for the above code to analyze the 2017 data was 1.132813 seconds.

![Original_Code_2017_RunTime](https://github.com/lmobashe/StockAnalysis/blob/main/Resources/Original_Code_2017_RunTime.PNG)

The run time for the above code to analyze the 2018 data was 1.117188 seconds.

![Original_Code_2018_RunTime](https://github.com/lmobashe/StockAnalysis/blob/main/Resources/Original_Code_2018_RunTime.PNG)


### Refactored Code

The refactored code utilizes 4 arrays to collect data for the 12 stocks provided in Steve's data set. The code sets one array of stock types with each stock being an index in the array, one array of total volumes per stock, one array of stock starting prices, and one array of stock ending prices. It sets a variable (tickerIndex) at zero, the index increases by one every time the loop cycles through. It then loops through all rows of data tracking the tickerIndex count which matches each index in the tickers array. The volume for each index is added together, it finds the start price of the index and the end price to find the total return for each index. This code was successful in collecting comparative data and increased the code processing speed by almost 600%, see code below:


````
```

Sub AllStocksAnalysisRefactored()
    
    'Counts code run time
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    rowStart = 2
    
    '1a) Create a ticker Index and initialize tickerIndex to zero
    Dim tickerIndex As Integer
        tickerIndex = 0

    '1b) Create three output arrays
            Dim tickerVolumes(12) As Long
            
            Dim tickerStartingPrices(12) As Single
            
            Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0
         
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = rowStart To RowCount
        
                    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
                
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'Then set current ticker starting price.
        'If  Then
                
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3c) check if the current row is the last row with the selected ticker. Then set current ticker ending price.
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                        
        End If
    Next i
        
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'Output code run time as message
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```
````

After refactoring the code, the run time for the 2017 data was 0.1953125 seconds. This is an improved run time of 580%.

![Refactored_Code_2017_RunTime](https://github.com/lmobashe/StockAnalysis/blob/main/Resources/Refactored_Code_2017_RunTime.PNG)

The refactored run time for the 2018 data was 0.1875. This is an improved run time of 595.83%. 

![Refactored_Code_2018_RunTime](https://github.com/lmobashe/StockAnalysis/blob/main/Resources/Refactored_Code_2018_with_RunTIme.PNG)

## Summary

- What are the advantages or disadvantages of refactoring code?

### Advantages of Refactoring Code:

Refactoring code allows you to improve the processing times of your code. Additionally, refactoring gives you an opportunity to simplify the code logic. A VBA macro with simpler logic and lower processing times will allow it to be applied to larger data sets. A simpler logic also makes changes and corrections easier to make in the future. Lastly, refactoring your code can give a programmer an opportunity to find and correct bugs and errors in logic in the original code.

### Disadvantages of Refactoring Code:

The disadvantage of refactoring code is that it can take a lot of time and effort to rework the logic and an improved code is not guaranteed. If the code is long and complicated it can be risky to try and rework it, you could invest a lot of time into refactoring a code and break it along the way. 

- How do these pros and cons apply to refactoring the original VBA script?

Refactoring our original code was worth the effort because our new macro runs so much more quickly. Steve wants to use this code to analyze data for all the stock market which will be a large data set. Our refactored code will be more capable of managing such a large data set in a reasonable amount of time without crashing.