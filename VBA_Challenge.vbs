Sub Module2Challenge()
Dim Ticker As String
Dim StockOpen As Double
Dim StockClose As Double
Dim YearlyChange As Double
Dim TotalStockVolume As Double
Dim PercentChange As Double
Dim TickerCount As Integer
Dim CustomIndex As Long

Dim ws As Worksheet


    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        TickerCount = 2
        CustomIndex = 1
        TotalStockVolume = 0
        
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To RowCount
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                CustomIndex = CustomIndex + 1
                
                StockOpen = ws.Cells(CustomIndex, 3).Value
                StockClose = ws.Cells(i, 6).Value
                
                
                For j = CustomIndex To i
                
                TotalStockVolume = TotalStockVolume + ws.Cells(j, 7).Value
                
                Next j
                
                    If StockOpen = 0 Then
                    PercentChange = StockClose
                
                    Else
                    YearlyChange = StockClose - StockOpen
                    PercentChange = YearlyChange / StockOpen
                
                    End If
                
                    ws.Cells(TickerCount, 9).Value = Ticker
                    ws.Cells(TickerCount, 10).Value = YearlyChange
                    ws.Cells(TickerCount, 11).Value = PercentChange
                    ws.Cells(TickerCount, 12).Value = TotalStockVolume
                    ws.Cells(TickerCount, 11).NumberFormat = "0.00%"
                    CondColor = ws.Cells(TickerCount, 10).Value
                        Select Case CondColor
                        Case Is > 0
                        ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                        Case Is < 0
                        ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                        Case Else
                        ws.Cells(TickerCount, 10).Interior.ColorIndex = 0
                        End Select

                
                    TickerCount = TickerCount + 1
            
                    TotalStockVolume = 0
                    YearlyChange = 0
                    PercentChange = 0
                    
                    CustomIndex = i
                
                End If
            
                
            Next i
            
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Range("P2") = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("P3") = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L:L"))
    MaxIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0)
    ws.Range("O2") = ws.Cells(MaxIndex, 9)
    MinIndex = WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0)
    ws.Range("O3") = ws.Cells(MinIndex, 9)
    VolIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)
    ws.Range("O4") = ws.Cells(VolIndex, 9)
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"

            
        
Next ws

End Sub
