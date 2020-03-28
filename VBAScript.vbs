Sub stock()

Dim count_ticker As Integer
Dim summary_row As Integer
Dim volume As Double
Dim yearly_change As Double
Dim percent_change As Double


For Each ws In Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    count_ticker = 0
    
    summary_row = 2
    
    volume = 0
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
            'Fill ticker
            ticker = ws.Cells(i, 1).Value
            ws.Cells(summary_row, 9).Value = ticker
            
            
            'Yearly change
            count_ticker = count_ticker + 1
            yearly_change = ws.Cells(i, 6).Value - ws.Cells(i - count_ticker + 1, 3).Value
            ws.Cells(summary_row, 10).Value = yearly_change
            
            'Percent change
            percent_change = yearly_change / ws.Cells(i - count_ticker + 1, 3).Value
            ws.Cells(summary_row, 11).Value = percent_change
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            
            'Volume
            volume = volume + ws.Cells(i, 7).Value
            ws.Cells(summary_row, 12).Value = volume
            
            count_ticker = 0
            
            summary_row = summary_row + 1
            
            volume = 0
            
        Else
            count_ticker = count_ticker + 1
            volume = volume + ws.Cells(i, 7).Value
            
            
        End If
    Next i

Next ws

End Sub


