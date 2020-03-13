Sub stockdata()

For Each ws In Worksheets

    Dim tickersymbol As String
    Dim lastrow As Double
    Dim totalstockvolume As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    SummaryRow = 1
    totalstockvolume = 0
    
    'Naming summary table headers
    ws.Cells(1, 9).Value = "ticker_symbol"
    ws.Cells(1, 10).Value = "yearly_change"
    ws.Cells(1, 11).Value = "percent_change"
    ws.Cells(1, 12).Value = "total_stock_volume"
    ws.Cells(3, 14).Value = "max_%_change"
    ws.Cells(4, 14).Value = "min_%_change"
    ws.Cells(5, 14).Value = "largest_total_volume"
    ws.Cells(2, 15).Value = "Ticker"
    ws.Cells(2, 16).Value = "Value"
    

    'Formatting summary table columns
    ws.Columns(12).NumberFormat = "#,##0"
    ws.Columns(11).NumberFormat = "0.00%"
    ws.Range("P3:P4").NumberFormat = "0.00%"
    ws.Range("P5").NumberFormat = "#,##0"
    
    On Error Resume Next
    
    'Getting unique ticker symbols and total stock volume
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'Ticker symbol
            tickersymbol = ws.Cells(i, 1).Value
            SummaryRow = SummaryRow + 1
            ws.Cells(SummaryRow, 9).Value = tickersymbol
            'Total stock Volume
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            ws.Cells(SummaryRow, 12).Value = totalstockvolume
            totalstockvolume = 0
            'Change in price
            Closeprice = ws.Cells(i, 6).Value
            yearlychange = Closeprice - openprice
            ws.Cells(SummaryRow, 10).Value = yearlychange
            'percent change
            percentchange = yearlychange / openprice
            ws.Cells(SummaryRow, 11).Value = percentchange
            'Conditional formatting for yearly change
            If yearlychange >= 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            End If
            openprice = 0
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            openprice = ws.Cells(i, 3).Value
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        Else
            'Total stock Volume
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Max = 0
    Min = 0
    
    'Summary table for min and max percent change
    For i = 2 To lastrow
        If ws.Cells(i, 11).Value > Max Then
            Max = ws.Cells(i, 11)
            ws.Cells(3, 16).Value = Max
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value < Min Then
            Min = ws.Cells(i, 11)
            ws.Cells(4, 16).Value = Min
            ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        End If
    Next i
     
    'Summary table for max volume
    For i = 2 To lastrow
        If ws.Cells(i, 12) > Max Then
            Max = ws.Cells(i, 12).Value
            ws.Cells(5, 16).Value = Max
            ws.Cells(5, 15).Value = ws.Cells(i, 9).Value
        End If
    Next i
        
    'Autofitting the column widths
    ws.Columns.AutoFit
    
Next ws
    
End Sub

