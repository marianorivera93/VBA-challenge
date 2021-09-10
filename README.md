# VBA-challenge
Sub MultipleYearStockData():

'Basic String Variables

Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String

' loop over each worksheet
For Each ws In Worksheets

    ' Activate Worksheet
    ws.Activate

    ' Find final row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Worksheets Variables
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Loop through list of tickers.
    For i = 2 To lastRowState

        ' Ticker Value
        ticker = Cells(i, 1).Value
        
        ' Ticker Opening Price
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Ticker Total Stock Volume Values
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' If Different Ticker Value Occurs
        If Cells(i + 1, 1).Value <> ticker Then
            ' Different Ticker Number
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' Ticker End of Year Closing Price
            closing_price = Cells(i, 6)
            
            ' Yearly Change Value
            yearly_change = closing_price - opening_price
            
            ' Yearly Change Value For Every Worksheet
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' If Greater than 0, Green.
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' If Less than 0, Red.
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' If 0, Yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Percent Change Value
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Change Opening Price
            opening_price = 0
            ' Change to Percent
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' Total Stock Volume Value
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Total Stock Volume Back to 0
            total_stock_volume = 0
        End If
        
    Next i
    
    ' Bonus
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Last Row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Variables
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' Loop
    For i = 2 To lastRowState
    
        ' Greatest Percent Increase
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Greatest Percent Decrease
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Greatest Stock Volume
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Values For Each Worksheet
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub
