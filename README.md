# Green Stock Analysis using VBA
 
## Overview

#### The parents of our client, Steve, want to invest in green energy stock, especially in DAQO. Because of this, Steve asked us to do a dataset analysis to identify the market performance of green stocks. We used excel and VBA, with which we created a macro that allowed us to facilitate the analysis and data visualization so that Steve can make the best recommendations to his parents.
#### However, once we programmed our VBA script, the newest challenge was to improve the macro.

## Purpose
#### Our purpose is to refactor and improve the VBA code so it can run faster. That is, our goal is to refactor the macro to make it more efficient.

## Results
#### Previously, our code for green energy stocks analysis used nested loops to report data from 2017 and 2018 sheets. Below are the images of the results and the time it took to run the macro:

###### _Screenshots of the elapsed time of the previous version for 2017 and 2018_

![Alt text](/module2_2017.png "imagen0")
![Alt text](/module2_2018.png "imagen1")

#### Through the revision and refactoring of the VBA script, it was possible to improve the code by reducing the number of iterations to find and collect the data in less time. The order and logic of the code was also improved, making it more functional. As we can see in the next images, editing the code also reduced the time:

###### _Screenshots of the elapsed time of the newest version for 2017 and 2018_

![Alt text](/challenge_2017.png "imagen3")
![Alt text](/challenge_2018.png "imagen4")

#### To conclude with the results section, we share  our refactored code, which allowed us to improve the performance of the macro. However, we attach the files worked with VBA and Excel in Module 2 and the Challenge.

```
Sub AllStocksAnalysisRefactored()
   
   Dim startTime As Single
   Dim endTime As Single
   
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

   '1b) Create three outputs arrays
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single
   
    '2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

    '2b) Loop over all the rows in thespreadsheet
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        '3c) Check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3d) Increase the tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
        End If
            
    Next i
    
    '4) Loop through your arrays to output the Ticker,Total Daily Volume, and Return
    For i = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i
   
    'Formatting
    Sheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        Else
            'Color the cell red
             Cells(i, 3).Interior.Color = vbRed

        End If

    Next i
    
    endTime = Timer
    MsgBox "This code can ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Summary

## Advantages a Disadvantages of refactoring code in general
### Advantages:
#### -Refactoring code teaches us that in programming there are many ways to solve problems. It's always important to find the most efficient solution.
#### -Improving the code helps to reduce time.
#### -Refactoring is fundamental to making the code more efficient because it looks cleaner and more structured.
### Disadvantages:
#### -Code refactoring requires a lot of time, effort, and understanding.

## Advantages and disadvantages of the original and refactored VBA script
#### -Advantages: Although the code worked in its older version, after refactoring, the macro could run faster. So, we can conclude that improving the code helps to reduce the time in which the macros run.
#### -Disadvantages: It's hard to find ways to refactor VBA scripts because it requires experience and patience.

## Recommendation for our client:
#### The analysis of the dataset shows the performance of the green stocks during 2017 and 2018. In the first year of study (2017), we overall observed good performance on green energy stocks; while in the next year (2018) the performance worsened for most companies.
#### We conclude that investing in DAQO is not recommended, so we suggest Steve update the dataset (with 2019, 2020 and 2021 data) and replicate the analysis before deciding where to invest.
