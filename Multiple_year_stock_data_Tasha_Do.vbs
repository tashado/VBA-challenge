Sub ticker()

For Each ws In Worksheets

    ' Insert Summary Table column names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ' Set variables for the Summary Table columns
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    total_stock = 0
    
    ' track location for each ticker in summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ' find last row of data set
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' loop through all rows of data
    For i = 2 To last_row
    
        ' Check if it's the first row of data
        If i = 2 Then
            
            ' Set the opening price of the first ticker AAB
            open_price = ws.Cells(i, 3).Value
            
            ' Add to the total stock volume amount
            total_stock = total_stock + ws.Cells(i, 7).Value
        
        ' Check if we are within the same ticker, if not
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Set the ticker
            ticker = ws.Cells(i, 1).Value
            
            ' Add to the total volume
            total_stock = total_stock + ws.Cells(i, 7).Value
            
            ' Set the closing price of current ticker
            close_price = ws.Cells(i, 6).Value
            
            ' Calculate yearly change
            yearly_change = close_price - open_price
            
            ' Calculate percent change
            percent_change = (close_price - open_price) / open_price
            
            ' Print ticker name to Summary Table
            ws.Range("I" & summary_table_row).Value = ticker
            
            ' Print yearly change to Summary Table
            ws.Range("J" & summary_table_row).Value = yearly_change
            
            ' Print percent change to Summary table and format to percent
            ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change)
            
            ' Conditional formatting of yearly and percent change
            If yearly_change >= 0 Then
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
                Else
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
            End If
            
            ' Print total volume to Summary Table
            ws.Range("L" & summary_table_row).Value = total_stock
            
            ' Move to next row of Summary Table
            summary_table_row = summary_table_row + 1
            
            ' Reset total stock volume amount
            total_stock = 0
            
            ' Set the opening price of next ticker
            open_price = ws.Cells(i + 1, 3).Value
            
        Else
        
            ' Add to the total stock volume amount
            total_stock = total_stock + ws.Cells(i, 7).Value
            
        End If
            
    Next i
    
    ' ----------
    ' BONUS
    ' ----------
    
    ' set values of the bonus summary table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    ' set variables for the bonus summary table
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim stock_ticker As String
    Dim increase As Double
    increase = 0
    Dim decrease As Double
    decrease = 0
    Dim stock As Double
    stock = 0
    
    ' loop through all rows of the Summary Table
    For i = 2 To summary_table_row
    
        ' if the current ticker has a bigger percent change than the previous ticker saved in the "increase" variable, then the new greatest % increase value is the current ticker's
        If ws.Cells(i, 11).Value > increase Then
            increase = ws.Cells(i, 11).Value
            increase_ticker = ws.Cells(i, 9).Value
        
        ' else if it's less than the previous ticker's decrease percent change, change thethen the new greatest % decrease value is the current ticker's
        ElseIf ws.Cells(i, 11).Value < decrease Then
            decrease = ws.Cells(i, 11).Value
            decrease_ticker = ws.Cells(i, 9).Value
            
        End If
        
        If ws.Cells(i, 12).Value > stock Then
            stock_ticker = ws.Cells(i, 9).Value
            stock = ws.Cells(i, 12).Value
        
        End If
        
    Next i
    
    ' Printing bonus summary table findings
    ws.Range("O2").Value = increase_ticker
    ws.Range("P2").Value = FormatPercent(increase)
    ws.Range("O3").Value = decrease_ticker
    ws.Range("P3").Value = FormatPercent(decrease)
    ws.Range("O4").Value = stock_ticker
    ws.Range("P4").Value = stock

Next ws

End Sub
