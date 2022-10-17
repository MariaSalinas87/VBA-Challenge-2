# VBA-Challenge-2
Stocks Research

Stece asked me to help him research a few years of stocks, to help his parents pick the best stocks to invest in. To do so I had to refactor the worksheet I had already submitted to Steve.

There were some edits that needed to be done to the program in order to make it work more efficiently. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/112505962/196081400-64413429-b19b-48fa-8c6c-4111faeb4c67.png)
As you can see in this screenshot, you can see that it took the program .070 seconds to run for the year 2017, as opposed to .496 in the previous code.


In this chart below, you will find that chart for the year 2018 and it took .070 as opposed to .484 in the previous written code.
![VBA_Challenge_2018](https://user-images.githubusercontent.com/112505962/196081553-a6790bef-29b8-44e8-ac4a-8c4c589c2a2f.png)


According to my research, refactoring the code makes the program run faster and more efficient. In less amount of time, I was able to get the results for Steve. 


Below is the code used to run the program:

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

    tickerIndex = 0

    '1b) Create 3 arrays named tickerVolumes, tickerStartingPrices, tickerEndingPrices
    Dim tickerVolumes() As Long
    ReDim tickerVolumes(0 To RowCount)
    Dim tickerStartingPrices() As Single
    ReDim tickerStartingPrices(0 To RowCount - 1)
    Dim tickerEndingPrices() As Single
    ReDim tickerEndingPrices(0 To RowCount - 1)
    
    '2a) Create a for loop to initialize the tickerVolumes to zero
    Worksheets(yearValue).Activate
    For i = 0 To RowCount - 1
        tickerVolumes(i) = 0
    Next i
    
    '2b) Create a for loop that will loop over all the rows in the spreadsheet
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If
        
        '3b)"   Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable
        If Cells(j - 1, 1).Value = tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(i) = Cells(j, 6).Value
        End If
        '3c) "   Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(i) = Cells(j, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i - 1, 1).Value And tickerIndex < 12 Then
            tickerIndex = tickerIndex + 1
        End If

    Next j
    
    
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

Advantages and Disadvantages of Refactoring

Generally speaking, refactoring codes makes the code cleaner and faster. it is a very helful tool to explain the process of the code. A disadvantage is that it may take too long and not advisable for larger code files. 

