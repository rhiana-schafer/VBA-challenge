Sub WallStreet()
    'define variables
    Dim StartRow As Long
    Dim RowCounter As Integer
    Dim OpenPrice As Double
    Dim TotalVol As LongLong
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVol As LongLong
    For Each ws in Worksheets
        'set/reset variables
        StartRow = 2
        RowCounter = 2
        TotalVol = 0
        'format summary table
        ws.Columns("I:L").ColumnWidth = 18
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
       
        'define how many rows we need to loop over
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'set starting price
        OpenPrice = ws.Cells(2, 3).Value
        'generate table
        For i = 2 To lastrow
            TotalVol = TotalVol + ws.Cells(i,7)
            'checks to see if last row of ticker has been reached
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'outputs ticker
                ws.Cells(RowCounter, 9).Value = ws.Cells(i, 1).Value
                'outputs delta price for year
                ws.Cells(RowCounter, 10).Value = ws.Cells(i, 6).Value - OpenPrice
                'outputs % change in price for year
                ws.Cells(RowCounter, 11).Value = ws.Cells(RowCounter, 10) / OpenPrice
                ws.Cells(RowCounter, 11).NumberFormat = "0.00%"
                'outputs total stock volume
                'ws.Cells(RowCounter, 12).Value = WorksheetFunction.Sum(ws.Range("G" & startrow & ":G" & i))
                ws.Cells(RowCounter, 12).Value = TotalVol
                'color formatting
                If ws.Cells(RowCounter, 10) < 0 Then
                'turns negative change red
                    ws.Cells(RowCounter, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(RowCounter, 10) > 0 Then
                'turns positive change green
                    ws.Cells(RowCounter, 10).Interior.ColorIndex = 4
                End If
                'reset conditions for next loop
                RowCounter = RowCounter + 1
                OpenPrice = ws.Cells(i + 1, 3).Value
                StartRow = i + 1
                TotalVol = 0
            End If
        Next 
        'bonus table
        'set/reset variables
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVol = 0
        'format bonus table
        ws.Columns("N:P").ColumnWidth = 18
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        'define how many rows we need to loop over
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'generate bonus table
        For j = 2 to lastrow2
            'Check if current % increase is bigger than previous
            If MaxIncrease < ws.Cells(j, 11).Value Then 
            'If so, replace previous with current'
                MaxIncrease = ws.Cells(j, 11).Value
                ws.Cells(2,15).Value = ws.Cells(j, 9).Value
                ws.Cells(2,16).Value = ws.Cells(j, 11).Value
                ws.Cells(2,16).NumberFormat = "0.00%"
            End If
            'Check if current % decrease is bigger than previous
            If MaxDecrease > ws.Cells(j, 11).Value Then
            'If so, replace previous with current
                MaxDecrease = ws.Cells(j,11).Value
                ws.Cells(3,15).Value = ws.Cells(j, 9).Value
                ws.Cells(3,16).Value = ws.Cells(j,11).Value
                ws.Cells(3,16).NumberFormat = "0.00%"
            End If
            'Check if current stock volume is bigger than previous
            If MaxVol < ws.Cells(j, 12).Value Then
            'If so, replace previoous with current
                MaxVol = ws.Cells(j,12).Value
                ws.Cells(4,15).Value = ws.Cells(j, 9).Value
                ws.Cells(4,16).Value = ws.Cells(j, 12).Value
            End If

        Next j
    Next ws
End Sub