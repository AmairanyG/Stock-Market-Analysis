Sub StockMarket()

    'Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        'Display headers with "Ticker" and "Total Stock Volume"
        Range("I1") = "Ticker"
        Range("J1") = "Total Stock Volume"
        Range("I1:J1").Font.Size = 12
        Range("I1:J1").Font.Bold = True
                
        'Define Variables
        Dim TotalStockVolume As Double
        Dim Volume As Double
        Dim Ticker1 As String
        Dim Ticker2 As String
        Dim TickerCounter As Long
        Dim i As Long
        Dim last_row As Long
        
        'Finding the last row in the ticker column
        last_row_1 = Cells(ws.Rows.Count, 1).End(xlUp).Row
        'Initializing ticker counter at 2 (the second row)
        TickerCounter = 2
        
            'Loop through first ticker through last ticker
            For i = 2 To last_row_1
            
                'Assign Values for ticker, next and ticker volume
                Ticker1 = ws.Cells(i, 1).Value
                Ticker2 = ws.Cells(i + 1, 1).Value
                Volume = ws.Cells(i, 7).Value
                
                'Take total Stock volume and add the next total volume of the same ticker variable
                TotalStockVolume = TotalStockVolume + Volume
                   
                'If ticker 1 is not equal to ticker 2, then combine the tickers alike and add the volume to the total Stock volume
                If Ticker1 <> Ticker2 Then
                    ws.Cells(TickerCounter, 9).Value = Ticker
                    ws.Cells(TickerCounter, 10).Value = TotalStockVolume
                    
                    'Reset the Total Stock Volume to 0 before moving to the next ticker value
                    TotalStockVolume = 0
                    
                    'Adding 1 to TickerCounter
                    TickerCounter = TickerCounter + 1
                    
                End If
                
            Next i
    
    Next ws

End Sub



