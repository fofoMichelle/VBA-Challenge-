# stock-analysis
# Overview of the project

Steve wants to analyze a group of green stocks for her to determine if they are worth investing, analysis he is doing for his parents's future investment decisions. To do so, we had to determine the annual volume and rate of return (ROI) on investment, and what better way than using the excel integrated Visual Basic Application (VBA) since the data were given in an excel.

By using VBA, we will be able to assist Steve with going through our analysis just by clicking on a button and he will be able to analyse 12 stocks 

He loved being able to analyze each stock at the click of a button and now wants to expand his research beyond the 12 green stocks.
Since Steve wants to analyse a high number of stocks we had to improve our code by refactoring it for the code to run faster and efficiently which is what I will be demonstrating below.
# Result
As mentioned above I have refactored the code to make it run more efficiently and easy to read and modify for future users.

#### Original Code
```
Sub AllStocksAnalysis()
 Dim startTime As Single
    Dim endTime  As Single
yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer
    
'1) Format the output sheet on All Stocks Analysis worksheets
Worksheets("All Stocks Analysis").Activate
Range("A1").Value = "All Stocks (" + yearValue + ")"
'Create a header row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
'2) Initialilize array of all tickers
Dim tickers(11) As String
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

'3a) Initialize variables for starting price and ending price
Dim startingPrice As Single
Dim endingPrice As Single

'3b) Activate data worksheet

Worksheets(yearValue).Activate

'3c)Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
'4) Loop through tickers
For i = 0 To 11
ticker = tickers(i)
totalVolume = 0
'5) loop through rows in the data

Worksheets(yearValue).Activate
For j = 2 To RowCount

'5a) Get total volume for current ticker
If Cells(j, 1).Value = ticker Then
totalVolume = totalVolume + Cells(j, 8).Value

End If
'5b) get stating price for current ticker
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

startingPrice = Cells(j, 6).Value

End If
'5c) get ending price for current ticker
If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
 endingPrice = Cells(j, 6).Value

End If

Next j
'6) Output data for current ticker

Worksheets("All Stocks Analysis").Activate

Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


Next i
  endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

#### Refactored Code
```
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
    
    '1a)Creating tickerIndex variable and setting it to 0 to access the correct index across the four different arrays.
    
    tickerIndex = 0

    '1b)Creating three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
   
    Next i

    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker and add ticker volume for the current stock ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex and assign current Starting Price.
            
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker and assign current ending price
        
            
           If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex if the next row's ticker does not match the previous row's ticker
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Analysis
Based on the screenshot below we can clear see that these stocks performed well in 2017 whereas there was a huge decline in returns in 2018 where only 2 stocks had a gain which will allow Steve to show to his parents that it might not be a good idea to invest in these stocks when looking at their return in 2018.

> 2018 performance
![VBA_Challenge_2018 refactored png](https://user-images.githubusercontent.com/99924850/158087729-748e56cb-6e4f-4f16-a373-9dabc0753318.png)
> 2017 performance
![VBA challenge 2017 refactored png(refactored)](https://user-images.githubusercontent.com/99924850/158088187-80fa7c83-3e99-4762-919b-551c3b45b555.png)

# Summary
## Advantages of refactoring a code
.* The code will run faster which can be useful when you have a huge amount of data, below I have included the execution time for 2018 to show the improvement in execution time 
![VBA_Challenge_2018 Original png](https://user-images.githubusercontent.com/99924850/158088564-2ad58292-65e5-4c1f-a0fc-7f7c871a6999.png)
![VBA_Challenge_2018 refactored png](https://user-images.githubusercontent.com/99924850/158088584-f01815c7-7bd3-40ca-a917-15f5bdf2388f.png)

.* Clean code are easier to understand and improve
## Disadavatanges
.* Refactoring a code can be time consuming
.* There is a room for error especially when working with applications that are too large
Regarding how this apply to our vba script, we have clearly improved our code's execution time but was this worth the time if our original code was already running efficiently.

