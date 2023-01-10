Attribute VB_Name = "Module1"
    Sub stocks()
    
    'Declare all your variables and what format they will go in as.
    Dim Ticker As String
    Dim Yearly_change As Double
    Dim Total_stock_volume As Double
    Dim Percent_change As Double
    Dim Summary_ticker_row As Integer
    Dim Greatest_percent_increase As Double
    Dim Greatest_percent_decrease As Double
    Dim Greatest_total_volume As Double
    Dim Opening_price_value As Double
    
    'Also you can declare a worksheet
    
    Dim ws As Worksheet
    
    
    ' First loop or Outer Loop, this is for each worksheet in the file.
    
    For Each ws In Worksheets
    Total_stock_volume = 0
    Summary_ticker_row = 2
    Opening_price_value = ws.Cells(2, 3).Value
    
    
    'The loop continues and you can use it to make the headers or titles for the column
    
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Yearly Change"
    
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"
    
    
    
    'This first determines the value of the last row by finding the last cell with an entry
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Then this loops through the rows, checks for continutity in ticker name per row and then adds values to the total stock volume
    
    For i = 2 To LastRow
    
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
    
    Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
    
    'This prints the ticker name into the column id'd
    
    ws.Range("I" & Summary_ticker_row).Value = Ticker
    
    ws.Range("L" & Summary_ticker_row).Value = Total_stock_volume
    
    '
    'Next we can change both yearly and percentage wise based on intial change of 0
    
    Yearly_change = ws.Cells(i, 6) - Opening_price_value
    ws.Range("J" & Summary_ticker_row).Value = Yearly_change
    
    'Change theformat of Column "Yearly Change" to accounting with "$"
    ws.Range("J" & Summary_ticker_row).NumberFormat = "$0.00"
    
    If Opening_price_value = 0 Then
    Percent_change = 0
    Else
    
    Percent_change = Yearly_change / Opening_price_value
    End If
    
    ws.Range("K" & Summary_ticker_row).Value = Percent_change
    
    'As with Yearly Change you change the format to "%"
    
    ws.Range("K" & Summary_ticker_row).NumberFormat = "0.00%"
    
    'Next reset the variables to move down the row and put total volume of tickers and opening day price back to 0
    Summary_ticker_row = Summary_ticker_row + 1
    
    Total_stock_volume = 0
    
    Opening_price_value = ws.Cells(1 + i, 3)
    
    Else
    
    Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
    
    
    End If
    
    
    Next i
    
    
    'Conditional formatting part for positive and negative values and color change
    
    Yearly_change = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    If ws.Range("J" & Summary_ticker_row).Value < 0 Then
    ws.Range("J" & Summary_ticker_row).Interior.ColorIndex = 3
    
    Else: ws.Range("J" & Summary_ticker_row).Interior.ColorIndex = 4
    
    End If
    
    
    Next ws
    
    
    End Sub
    




