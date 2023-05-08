Attribute VB_Name = "Module1"
Option Explicit

Sub StockMarket():
    'Create a worksheet variable to be able to loop through all the years
    Dim ws As Worksheet
    'Create Ticker variable and Total Volume variable
    Dim Ticker As String
    Dim Volume As Double
    'Create Open Price, Close Price, Percent Change, and Yearly Change
    Dim OpenPrice, ClosePrice, PercentChange, YearlyChange As Double
    'Create a counter for the total number of rows in the worksheet
    Dim nr As Long
    Dim i As Long
    'Create a counter to track the stock number
    Dim StockCount As Long
    'Create counters for a for loop to loop through columns I through L
    Dim j As Long
    Dim nr2 As Long
    'Create largest percent increase and decrease variables, and largest total volume variable
    Dim LargestPerIncr, LargestPerDecr, LargestTotalVol As Double
    
    'Use a for each statement to run through the years (worksheets) of the workbook
    For Each ws In Worksheets
              
        'Place headers on all worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change ($)"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Set opening price for the very first stock
        OpenPrice = Range("C2").Value
        
        'Set stock counter to 1
        StockCount = 1
        
        'Count all rows in column A
        nr = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Create a for loop to run through all data points in columns A through G
        For i = 2 To nr
            'If the ticker changes
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            
                'Increment StockCount by 1
                StockCount = StockCount + 1
                
                'Populate the Ticker column
                'Get ticker symbol
                Ticker = ws.Cells(i, 1).Value
                'Place ticker symbol under header
                ws.Cells(StockCount, 9).Value = Ticker
                
                'Deal with the yearly change
                'Closing price at the end of the year
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ws.Cells(StockCount, 10).Value = YearlyChange
                
                
                'Calculate the percent change, but do not divide by 0
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                    PercentChange = FormatPercent(PercentChange, 2)
                Else
                    PercentChange = 0
                End If
                
                'Display the Percent Change
                ws.Cells(StockCount, 11) = PercentChange
                
                'Set color to the YearlyChange column
                'If change is positive, set to green
                If ws.Cells(StockCount, 10).Value > 0 Then
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 4
                'If change is negative, set to red
                ElseIf ws.Cells(StockCount, 10).Value < 0 Then
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 3
                'Do nothing if no change
                End If
                
                'Apply red to a negative percent change
                If ws.Cells(StockCount, 11).Value < 0 Then
                    ws.Cells(StockCount, 11).Interior.ColorIndex = 3
                'Apply green to a positive percent change
                ElseIf ws.Cells(StockCount, 11) > 0 Then
                    ws.Cells(StockCount, 11).Interior.ColorIndex = 4
                'Do nothing if no change in percent change
                End If
             
                 'Find next opening price
                 OpenPrice = ws.Cells(i + 1, 3)
                 
                 'Add Volume
                 Volume = Volume + ws.Cells(i, 7).Value
                 ws.Cells(StockCount, 12).Value = FormatNumber(Volume, 0)
                                
                 'Reset Volume to 0
                  Volume = 0
                   
             Else
                  'Add Volume when the ticker symbol is the same
                  Volume = Volume + ws.Cells(i, 7).Value
                 
            End If
        
        Next i
    
    'Reset variables after running through a worksheet
    LargestPerIncr = 0
    LargestPerDecr = 0
    LargestTotalVol = 0
    
    'Get the number of rows in column I
    nr2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Create a for loop to run through columns I through L
    For j = 2 To nr2
        'Grab Percent Change if higher than that in Largest Percent Increase placeholder
         If ws.Cells(j, 11).Value > LargestPerIncr Then
                'Grab Stock ticker and percentage increase
                LargestPerIncr = ws.Cells(j, 11).Value
                ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 16).Value = FormatPercent(LargestPerIncr, 2)
        'Don't do anything if Percent Change is smaller than placeholder
        End If
        
        'Grab Percent Change if lower than that in Largest Percent Decrease placeholder
         If ws.Cells(j, 11).Value < LargestPerDecr Then
                'Grab Stock ticker and percentage decrease
                LargestPerDecr = ws.Cells(j, 11).Value
                ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 16).Value = FormatPercent(LargestPerDecr, 2)
        'Don't do anything if Percent Change is larger than placeholder
        End If
        'Grab largest total volume traded
        If ws.Cells(j, 12) > LargestTotalVol Then
                'Grab Stock ticker and total volume
                LargestTotalVol = ws.Cells(j, 12).Value
                ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
                ws.Cells(4, 16).Value = FormatNumber(LargestTotalVol, 0)
        'Don't do anything if volume change is smaller than placeholder
        End If
        
    Next j
       
   'AutoFit the width of the columns
    ws.Columns("A:Z").AutoFit

    'Call next worksheet in workbook
    Next ws
End Sub
