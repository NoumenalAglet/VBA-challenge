Attribute VB_Name = "Module1"
Sub main():

'Cycle through Worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    '''Initial Row Conditions'''
    'Set second row as first row of data. This is set up before Main Loop
    'RowIndexResults will also be used in the Main Loop to direct results output
    RowIndexResults = 2
    'Set Ticker value
    Tick = ws.Cells(RowIndexResults, 1).Value
    'Output Tick
    ws.Cells(RowIndexResults, 9).Value = Tick
    'Set FirstOpen value
    FirstOpen = ws.Cells(RowIndexResults, 3).Value
    'Set TotalStockVolume value
    TotalStockVolume = ws.Cells(RowIndexResults, 7).Value
    'Find and set last row with data
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    '''Output Labels'''
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"

    '''Define Variables so they are reset for each new sheet
    MaxRatioChange = 0
    MinRatioChange = 0
    MaxTick = ""
    MinTick = ""

    '''Main Loop'''
    ' For each row check if Same or Different than the cell below, then output results
    For i = 3 To EndRow
        'If Same
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            'add to TotalStockVolume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            'In the rare case of a stock opening at 0 then starting up later
            If FirstOpen = 0 Then
                'Record new FirstOpen if it is different from the cell below
                If FirstOpen <> ws.Cells(i + 1, 3).Value Then
                    FirstOpen = ws.Cells(i + 1, 3).Value
                End If
                
            End If

        'If Different
        Else
            'Set LastClose, calc and output YearlyChange
            LastClose = ws.Cells(i, 6).Value
            YearlyChange = LastClose - FirstOpen
            ws.Cells(RowIndexResults, 10) = YearlyChange

            'For rare cases when the FirstOpen is zero we print "N/A", as a numerical value would be inapplicable
            If FirstOpen = 0 Then
                ws.Cells(RowIndexResults, 11) = "N/A"
            Else
                'We carry the Ratio change for a given Ticker, Percent change is only created at output
                RatioChangeTicker = YearlyChange / FirstOpen
                PercentChange = FormatPercent(RatioChangeTicker, 2)
                ws.Cells(RowIndexResults, 11) = PercentChange
                
                'We search for the highest or lowest Ratio change based on if the values are =>0 or not
                If RatioChangeTicker >= 0 Then
                    'If the ratio change is greater than current MaxRatioChange, save both it and the associated ticker
                    If RatioChangeTicker > MaxRatioChange Then
                        MaxRatioChange = RatioChangeTicker
                        MaxTick = Tick
                    End If
                 'If the ratio change is less than current MinRatioChange, save both it and the associated ticker
                Else
                    If RatioChangeTicker < MinRatioChange Then
                        MinRatioChange = RatioChangeTicker
                        MinTick = Tick
                    End If
                End If
                
            End If
            'Record new FirstOpen
            FirstOpen = ws.Cells(i + 1, 3).Value
            'Set green or red based on value range
            If RatioChangeTicker >= 0 Then
                ws.Cells(RowIndexResults, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(RowIndexResults, 10).Interior.ColorIndex = 3
            End If
            'add to TotalStockVolume, output and reset
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Cells(RowIndexResults, 12) = TotalStockVolume
            TotalStockVolume = 0
            '+1 to RowIndexResults to set up for next iteration
            RowIndexResults = RowIndexResults + 1
            'reset Tick and output for next iteration
            Tick = ws.Cells(i + 1, 1).Value
            ws.Cells(RowIndexResults, 9) = Tick
        End If

    Next i

    '''Done after the Main Loop is completed'''
    'Greatest % increase
    'Output results
    ws.Cells(2, 17).Value = FormatPercent(MaxRatioChange, 2)
    ws.Cells(2, 16).Value = MaxTick


    'Greatest % decrease
    'Output results
    ws.Cells(3, 17).Value = FormatPercent(MinRatioChange, 2)
    ws.Cells(3, 16).Value = MinTick

    'Greatest total volume
    GTV = 0
    For j = 2 To EndRow
        'If cell value is greater than current GTV, save GTV and ticker
        If ws.Cells(j, 12).Value > GTV Then
            GTV = ws.Cells(j, 12).Value
            Tick = ws.Cells(j, 9).Value
        End If
    Next j
    'Output results
    ws.Cells(4, 17).Value = GTV
    ws.Cells(4, 16).Value = Tick

Next ws
End Sub



