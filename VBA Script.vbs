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
            
                If yearly_change < 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                End If
            
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

Sub maxmin()
    For Each ws In Worksheets

    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"

    greatest_percent_increase = WorksheetFunction.Max(ws.Range("K:K"))
    Greatest_percent_decrease = WorksheetFunction.Min(ws.Range("K:K"))
    greatest_volume = WorksheetFunction.Max(ws.Range("L:L"))
    
    ws.Cells(2, 17).Value = greatest_percent_increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = Greatest_percent_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = greatest_volume
    
    Ticker_row_greatest_percent_increase = WorksheetFunction.Match(greatest_percent_increase, ws.Range("K:K"), 0)
    Ticker_row_greatest_percent_decrease = WorksheetFunction.Match(Greatest_percent_decrease, ws.Range("K:K"), 0)
    ticker_row_greatest_volume = WorksheetFunction.Match(greatest_volume, ws.Range("L:L"), 0)
    
    Get_ticker = ws.Cells(Ticker_row_greatest_percent_increase, 9).Value
    ws.Cells(2, 16).Value = Get_ticker
    
    get_ticker2 = ws.Cells(Ticker_row_greatest_percent_decrease, 9).Value
    ws.Cells(3, 16).Value = get_ticker2
    
    get_ticker3 = ws.Cells(ticker_row_greatest_volume, 9).Value
    ws.Cells(4, 16).Value = get_ticker3

    Next ws

End Sub

