VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAHomework():
        'define variables,then prep for overall worksheet movements
        'dim as double for decimals
        'dim as integer for whole numbers only
        'dim as long for long numbers
        'dim as string for words or phrases
Dim ticker As String
Dim number_tickers As Integer
Dim lastRow As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

' set up dynamic end row and to go through all worksheets
For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add headers and start points
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0

    'start loop
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        
        'get the opening price for the ticker
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        'Add up the total stock volume values
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        'move to next tickers
        If ws.Cells(i + 1, 1).Value <> ticker Then
            number_tickers = number_tickers + 1
            ws.Cells(number_tickers + 1, 9) = ticker
            
            'get the closing price for ticker
            closing_price = Cells(i, 6)
            
            'get yearly change value
            yearly_change = closing_price - opening_price
            
            'print the yearly change calue
            ws.Cells(number_tickers + 1, 10).Value = yearly_change
            
            'change colors with index green is 4
            If yearly_change > 0 Then
                ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
                'red is 3
                ElseIf yearly_change < 0 Then
                ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            End If
            
            ' Calculate percent change value for tickers
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            ' re-format the percent_change value as a percent
            ws.Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            'reset opening price for new tickers
            opening_price = 0
            'print total stock volume
            ws.Cells(number_tickers + 1, 12).Value = total_stock_volume
            'reset stock volumes to start new total with new ticker
            total_stock_volume = 0
        End If
        
    Next i
    
        
Next ws


End Sub

