Sub StockExercise():

    'Go to the first worksheet
Worksheets(1).Activate

    'Define variables for worksheet loop
Dim WS As Integer
Dim SheetCount As Integer

    'Sheet count
SheetCount = Sheets.Count

    'Sheet loop
For WS = 1 To SheetCount
Worksheets(WS).Activate
Cells(1, 1).Activate

    'Insert headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

    'Total row count
Dim RowCount As Double
RowCount = Cells(Rows.Count, 1).End(xlUp).Row

    'Define variables for tickers loop
Dim TickerCounter As Integer
Dim EarliestDate As Double
Dim LatestDate As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim VolumeSum As Double

    'Define starting values for tickers loop
TickerCounter = 2
EarliestDate = 99999999
LatestDate = 0

    'Loop throught the data
For i = 2 To RowCount
            
        'If statement to identify new ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
            'Insert name of ticker
        Cells(TickerCounter, 9).Value = Cells(i, 1)
            
            'If statement to stablish EarliestDate and OpenPrice
        If Cells(i, 2).Value < EarliestDate Then
            EarliestDate = Cells(i, 2).Value
            OpenPrice = Cells(i, 3).Value
        End If
            
            'If statement to stablish LatestDate and ClosePrice
        If Cells(i, 2).Value > LatestDate Then
            LatestDate = Cells(i, 2).Value
            ClosePrice = Cells(i, 6).Value
        End If
            
            'Calculate YearlyChange and PercentChange
        YearlyChange = ClosePrice - OpenPrice
        PercentChange = YearlyChange / OpenPrice
            
            'Increase VolumeSum
        VolumeSum = VolumeSum + Cells(i, 7).Value
             
             'Insert YearlyChange, PercentChange and VolumeSum per Ticker
        Cells(TickerCounter, 10).Value = YearlyChange
        Cells(TickerCounter, 11).Value = PercentChange
        Cells(TickerCounter, 12).Value = VolumeSum
            
            'Reset EarliestDate, LatestDate and VolumeSum for next ticker
        EarliestDate = 99999999
        LatestDate = 0
        VolumeSum = 0
            
            'Add extra row to TickerCounter for next ticker
        TickerCounter = TickerCounter + 1
                                 
    Else
            
            'If statement to stablish EarliestDate and OpenPrice
        If Cells(i, 2).Value < EarliestDate Then
            EarliestDate = Cells(i, 2).Value
            OpenPrice = Cells(i, 3).Value
        End If
            
            'If statement to stablish LatestDate and ClosePrice
        If Cells(i, 2).Value > LatestDate Then
            LatestDate = Cells(i, 2).Value
            ClosePrice = Cells(i, 6).Value
        End If
            
             'Increase VolumeSum
        VolumeSum = VolumeSum + Cells(i, 7).Value
            
    End If
                    
Next i

    'Total ticker count
Dim TickerCount As Double
TickerCount = Cells(Rows.Count, 9).End(xlUp).Row
        
    'Loop to set conditional formatting to YearlyChange
For i = 2 To TickerCount

        'Set YearlyChange format
    Cells(i, 10).NumberFormat = "0.00"
       
        'If statement to define fill color
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    Else
        Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    End If

Next i

    'Set PercentChange format
Range("K2:K" & TickerCount).NumberFormat = "0.00%"

    'Define variables for summaries loop
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim TickerIncrease As String
Dim TickerDecrease As String
Dim TickerVolume As String

    'Define starting values for summary Loop
GreatestIncrease = Cells(2, 11).Value
GreatestDecrease = Cells(2, 11).Value
GreatestVolume = Cells(2, 12).Value
TickerIncrease = Cells(2, 9).Value
TickerDecrease = Cells(2, 9).Value
TickerVolume = Cells(2, 9).Value

For i = 2 To TickerCount
        'If statement to define GreatestIncrease and its ticker
    If Cells(i, 11).Value > GreatestIncrease Then
        GreatestIncrease = Cells(i, 11).Value
        TickerIncrease = Cells(i, 9).Value
     End If

        'If statement to define GreatestDecrease and its ticker
    If Cells(i, 11).Value < GreatestDecrease Then
        GreatestDecrease = Cells(i, 11).Value
        TickerDecrease = Cells(i, 9).Value
    End If
    
        'If statement to define GreatestVolume and its ticker
    If Cells(i, 12).Value > GreatestVolume Then
        GreatestVolume = Cells(i, 12).Value
        TickerVolume = Cells(i, 9).Value
     End If
     
Next i

    'Insert GreatestIncrease value and ticker, and define percent format
Cells(2, 16).Value = TickerIncrease
Cells(2, 17).Value = GreatestIncrease
Cells(2, 17).NumberFormat = "0.00%"

    'Insert GreatestDecrease value and ticker, and define percent format
Cells(3, 16).Value = TickerDecrease
Cells(3, 17).Value = GreatestDecrease
Cells(3, 17).NumberFormat = "0.00%"

    'Insert GreatestVolume value and ticker
Cells(4, 16).Value = TickerVolume
Cells(4, 17).Value = GreatestVolume

    'Final column size adjustments for readability
Columns("I:Q").EntireColumn.AutoFit

Next WS
        
    'Go back to Worksheet 1
Worksheets(1).Activate

End Sub
