Attribute VB_Name = "Module1"
Sub VBAHomework()

    ' Loop throuhgh each worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets

        'Label the summary table
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Format Summary Table Headers as Bold
        ws.Range("I1:L1").Font.FontStyle = "Bold"
        
        ' Define Ticker Symbol
        Dim TickerSymbol As String
        
        ' Define total stock volume & set to zero
        Dim TotalStock_Volume As Long
        TotalStockVolume = 0
        
        'Define the summary table & start data in the second row
        Dim SummaryTable As Long
        SummaryTable = 2
        
        ' Define open & close prices & percent change
        Dim StockOpen As Double
        Dim StockClose As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        ' Set value of stock open price
        StockOpen = ws.Cells(2, 3)
        
        ' Find the last row with data
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all ticker symbols
        For i = 2 To LastRow
        
            ' Set value of stock open price
            StockOpen = ws.Cells(2, 3)
        
                ' Check if we are still within the same ticker symbol, if not then ...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                    ' Set ticker symbol
                    TickerSymbol = ws.Cells(i, 1).Value
        
                    ' Add to the stock volume
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                    ' Grab closing prices
                    StockClose = ws.Cells(i, 6).Value
                
                    ' Calculate the yearly change
                    YearlyChange = StockClose - StockOpen
            
                    ' Print ticker symbol in summary table
                    ws.Range("I" & SummaryTable).Value = TickerSymbol
        
                    ' Print stock volume in summary table
                    ws.Range("L" & SummaryTable).Value = TotalStockVolume
                
                    ' Print yearly change in summary table
                    ws.Range("J" & SummaryTable).Value = YearlyChange
                
                    ' Calculate % change
                    PercentChange = (YearlyChange / StockOpen)
                
                    'Print percent change in summary table
                    ws.Range("K" & SummaryTable).Value = PercentChange
                
                    'Change to a %
                    ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
                
                        ' Check if % change was positive or negative
                        If ws.Range("J" & SummaryTable).Value >= 0 Then
                    
                            ' Fill green
                            ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                        
                        Else
                        
                        ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                        
                        End If
            
                    ' Add one to the summary row
                    SummaryTable = SummaryTable + 1
            
                    ' Reset stock volume
                    TotalStockVolume = 0
            
            ' If the cell following a row is the same brand
            Else
        
                ' Add to the stock volume
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
    Next ws

End Sub



