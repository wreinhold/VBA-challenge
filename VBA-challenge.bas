Attribute VB_Name = "Module1"
Sub ticker_analysis()
    ' Creating the variables
    Dim symbol As String
    Dim open_price As Double
    Dim close_price As Double
    Dim year_change As Double
    Dim volume As Double
    Dim num_rows As Double
    Dim count_stocks As Integer
    
    ' Looping through all of the worksheets
    For Each ws In Worksheets
        ' Finding the # of rows to use in future for loops
        num_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        count_stocks = 1
        ' Creating the headers for the new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To num_rows
            If ws.Cells(i, 1).Value <> "" Then
                ' Checks to see if the stock is new or not
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' sets the starting values for a stock
                    symbol = ws.Cells(i, 1).Value
                    open_price = ws.Cells(i, 3).Value
                    count_stocks = count_stocks + 1
                    volume = ws.Cells(i, 7).Value
                Else
                    ' if stock is not new, adds to volume and sets a new close price
                    close_price = ws.Cells(i, 6).Value
                    volume = volume + ws.Cells(i, 7).Value
                End If
                
                ' Creates the new columns
                ws.Cells(count_stocks, 9).Value = symbol
                ws.Cells(count_stocks, 10).Value = close_price - open_price
                ' if the year change was negative, it will be highlighted red. Otherwise, it will be green
                If ws.Cells(count_stocks, 10).Value < 0 Then
                    ws.Cells(count_stocks, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(count_stocks, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(count_stocks, 11).Value = FormatPercent((ws.Cells(count_stocks, 10).Value / open_price), 2)
                ws.Cells(count_stocks, 12).Value = volume
            End If
        Next i
        
        ' Looks for the Greatest % Increase, % Decrease, and Total Volume
        For j = 2 To count_stocks
            If ws.Cells(j, 11).Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 17).Value = FormatPercent(ws.Cells(j, 11).Value, 2)
            ElseIf ws.Cells(j, 11).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 17).Value = FormatPercent(ws.Cells(j, 11).Value, 2)
            ElseIf ws.Cells(j, 12).Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
            End If
        Next j
    Next ws
End Sub
