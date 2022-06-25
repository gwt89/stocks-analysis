# stocks-analysis
Purpose

The purpose of this analysis was to refactor the code from sheets 2017 and 2018 in Microsoft Excel VBA to see if any of the stocks are worth investing in and to see if refactoring the code would allow it to be processed faster than originally.

Results

Sub AllStocksAnalysisRefactored()
    
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
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
         End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            End If

            
        'End If
    
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Both sheets 2017 and 2018 contain the same type of information; a ticker name for the stock, the date the stock was issued, the opening price, the highest price, the lowest price, the closing price, the adjusted closing price, and the volume. The sheet I created, “All Stocks Analysis” contains the ticker, the total daily volume, and the return for each stock.

![Screenshot 2022-06-24 2017 excel](https://user-images.githubusercontent.com/105949411/175752644-a7904094-aad9-4aa7-b873-471549c26a83.png)

![Screenshot 2022-06-24 2018 excel](https://user-images.githubusercontent.com/105949411/175752650-29ff7203-eb82-40c2-8925-d4a50ca76d72.png)

Pros And Cons of Refactoring Code

The main purpose for refactoring code is to make code more precise and readable. It can also allow you to narrow things down to get the most useful data out of your original source code. Refactoring code can help to eliminate unnecessary or confusing code. If the original source data is unclear refactoring the code could lead to creating more problems. Sometimes refactoring also may not yield results that improve your analysis of the data. If the original data is more straight forward than the refactored data, then refactoring the data is likely a waste of time.
The main benefit of refactoring the code for this analysis is that the refactored code was able to run much faster than the original source code. The original source for 2017 took 0.58 seconds to run and the 2018 code also took .58 seconds to run. 

![VBA_Challenge_2017 Original Data](https://user-images.githubusercontent.com/105949411/175752662-e7bb81c9-f22e-4368-b8b3-e9ea399bfdf3.png)
![VBA_Challenge_2018 Original Data](https://user-images.githubusercontent.com/105949411/175752664-0759dcb0-119e-4046-b1a5-6b271fc064f9.png)

The new 2017 code took 0.10 seconds to run and the 2018 code took 0.09 seconds to run. The refactored code was clearly much faster to run than the original source code.

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/105949411/175752674-561cd582-cf21-4a71-b049-b794b7b6bc3c.png)
![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/105949411/175752675-3d68da98-11b9-4c35-983b-ec73b095e39e.png)
