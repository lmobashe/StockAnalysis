# StockAnalysis

##Overview of Project

Steve has asked me to write a code to analyze stock data he provided me. Using VBA I wrote a set of macros that allowed Steve to compare the performance of various stocks based on the data provided. Steve wants to be able to perform a similar analysis on all stock market data. To accomadate so much data I am refactoring my original code to see if I can make the macros run more quickly.

##Results

My Original code utilized 1 array to collect data for the 12 stocks provided. This code was successful in collecting comparative data, but consistently has a run time over 1 second, see code below:

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
    Dim tickers(11) As String
        
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
    
    'a.  Initialize variables for the starting price and ending 	price.
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
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 				1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'c.  Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 				1).Value = ticker Then
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
       
    'Format returns analysis to have color Red for loss, green for 	gain
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
    MsgBox "This code ran in" & (endTime - startTime) & "seconds 	for the year " & (yearValue)
        
End Sub

```
````



##Summary

- What are the advantages or disadvantages of refactoring code?

- How do these pros and cons apply to refactoring the original VBA scrpt?