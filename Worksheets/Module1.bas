Attribute VB_Name = "Module1"
Sub stock_analysis()

'define variables

Dim row_count, iterator, row_number As Integer
Dim ticker As String
Dim open_price, close_price, yearly_change As Double
Dim percent_change, total_volume As Double
Dim max_inc, max_dec, max_total As Double
Dim max_inc_ticker, max_dec_ticker, max_total_ticker As String


'loop through each worksheet
For Each ws In Worksheets
    ws.Cells(1, "I") = "Ticker"
    ws.Cells(1, "J") = "Yearly Change"
    ws.Cells(1, "K") = "Percentage Change"
    ws.Cells(1, "L") = "Total Volume"
    ws.Cells(2, "N") = "Greatest Percent Increase"
    ws.Cells(3, "N") = "Greatest Percent Decrease"
    ws.Cells(4, "N") = "Greatest Total Volume"
    ws.Cells(1, "O") = "Ticker"
    ws.Cells(1, "P") = "Value"

'get total row count as you go through each worksheet
row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row

'assign starting values
row_number = 2
open_price = ws.Cells(2, "C")
yearly_change = 0
percent_change = 0
max_inc = 0
max_dec = 0
max_total = 0

'iterate through rows and populate variables
For i = 2 To row_count

    'conditional statement to get unique ticker values
    If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
    
        'get ticker name and add it to a separate column
        ticker = ws.Cells(i, "A")
        ws.Cells(row_number, "I") = ticker
        
        'get close price, calculate yearly change, and add it to a separate column
        close_price = ws.Cells(i, "F")
        yearly_change = close_price - open_price
        ws.Cells(row_number, "J") = yearly_change
        
        'conditional statement (if then) for red and green
        If yearly_change > 0 Then
            ws.Cells(row_number, "J").Interior.ColorIndex = 4
            Else
            ws.Cells(row_number, "J").Interior.ColorIndex = 3
        End If
        
        'calculate percent change and add it to a separate column
        percent_change = yearly_change / open_price
        ws.Cells(row_number, "K") = FormatPercent(percent_change)
        
        'grab tickers with max increase, decrease, and total volume
        If ws.Cells(row_number, "K") > max_inc Then
            max_inc = ws.Cells(row_number, "K")
            max_inc_ticker = ticker
        End If
        
        If ws.Cells(row_number, "K") < max_dec Then
            max_dec = ws.Cells(row_number, "K")
            max_dec_ticker = ticker
        End If
        
        If ws.Cells(row_number, "L") > max_total Then
            max_total = ws.Cells(row_number, "L")
            max_total_ticker = ticker
        End If
        
        'assign a different value to starting variables
        open_price = ws.Cells(i + 1, "C")
        row_number = row_number + 1
        
    End If
    
    'calculate total volume and add it to a separate column
    If ws.Cells(i, "A") = ws.Cells(i + 1, "A") Then
       total_volume = total_volume + ws.Cells(i, "G")
       ws.Cells(row_number, "L") = total_volume + ws.Cells(i + 1, "G")
    Else
        total_volume = 0
    End If

Next i
    'print max increase, decrease, and total volume in ws
    ws.Cells(2, "O") = max_inc_ticker
    ws.Cells(2, "P") = FormatPercent(max_inc)
    ws.Cells(3, "O") = max_dec_ticker
    ws.Cells(3, "P") = FormatPercent(max_dec)
    ws.Cells(4, "O") = max_total_ticker
    ws.Cells(4, "P") = max_total
Next ws

End Sub

