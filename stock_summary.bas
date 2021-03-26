Attribute VB_Name = "Module1"
Sub stock_summary()
    For Each ws In Worksheets
        Dim stockOpen As Double
        Dim stockClose As Double
        Dim stockVolume As Double
        Dim summaryTable As Integer
        Dim greatIncrease As Double
        Dim greatDecrease As Double
        Dim greatVolume As Double
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryTable = 2
        stockVolume = 0
        stockOpen = ws.Cells(2, 3).Value
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                stockClose = ws.Cells(i, 6).Value
                ws.Cells(summaryTable, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summaryTable, 10).Value = stockOpen - stockClose
                If stockOpen = 0 Then
                    ws.Cells(summaryTable, 11).Value = 0
                Else
                    ws.Cells(summaryTable, 11).Value = (stockOpen - stockClose) / stockOpen
                End If
                ws.Cells(summaryTable, 12).Value = stockVolume
                ws.Cells(summaryTable, 11).NumberFormat = "0.00%"
          
                If ws.Cells(summaryTable, 10).Value < 0 Then
                    ws.Cells(summaryTable, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryTable, 10).Interior.ColorIndex = 4
                End If
            
                stockOpen = ws.Cells(i + 1, 3).Value
                stockVolume = 0
                summaryTable = summaryTable + 1
            Else
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
        Next i
    
        summarylastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        greatIncrease = ws.Cells(2, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(2, 9).Value
        greatDecrease = ws.Cells(2, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(2, 9).Value
        greatVolume = ws.Cells(2, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(2, 9).Value
        ws.Cells(2, 17).Value = greatIncrease
        ws.Cells(3, 17).Value = greatDecrease
        ws.Cells(4, 17).Value = greatVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        For j = 2 To summarylastrow
            If ws.Cells(j, 11).Value > greatIncrease Then
                greatIncrease = ws.Cells(j, 11).Value
                ws.Cells(2, 17).Value = greatIncrease
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ElseIf ws.Cells(j, 11).Value < greatDecrease Then
                greatDecrease = ws.Cells(j, 11).Value
                ws.Cells(3, 17).Value = greatDecrease
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            End If
        
            If ws.Cells(j, 12).Value > greatVolume Then
                greatVolume = ws.Cells(j, 12).Value
                ws.Cells(4, 17).Value = greatVolume
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            End If
        Next j
    Next ws
End Sub
