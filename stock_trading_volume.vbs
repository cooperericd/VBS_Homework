Sub stock_trading_volume()

    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest Percent Increase"
        ws.Range("N3").Value = "Greatest Percent Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        

        Dim WorksheetName As String
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        WorksheetName = ws.Name
    
        Dim Ticker_Symbol As String
    
        Dim Volume As Double

        Dim Ticker_Table_Row As Integer
        
        Dim Open_Price As Double
        
        Dim Close_Price As Double
    
        Open_Price = ws.Cells(2, 3).Value
        Volume = 0
        Ticker_Table_Row = 2
        

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Ticker_Symbol = ws.Cells(i, 1).Value
        
                Volume = Volume + ws.Cells(i, 7).Value
                
                Close_Price = ws.Cells(i, 6).Value
        
                ws.Range("I" & Ticker_Table_Row).Value = Ticker_Symbol
                
                ws.Range("J" & Ticker_Table_Row).Value = Close_Price - Open_Price
                
                ws.Range("K" & Ticker_Table_Row).Value = Format((Close_Price - Open_Price) / Open_Price, "Percent")
        
                ws.Range("L" & Ticker_Table_Row).Value = Volume
                
                If ws.Range("J" & Ticker_Table_Row).Value < 0 Then
                    ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
                Else: ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
                End If
        
                Ticker_Table_Row = Ticker_Table_Row + 1
        
                Volume = 0
                
                Open_Price = ws.Cells(i + 1, 3).Value
        
            Else
        
                Volume = Volume + ws.Cells(i, 7).Value
    
            End If

        Next i
        
        Max_return = Format(WorksheetFunction.Max(ws.Range("K:K")), "Percent")
        Min_return = Format(WorksheetFunction.Min(ws.Range("K:K")), "Percent")
        Max_volume = WorksheetFunction.Max(ws.Range("L:L"))
        
        ws.Range("P2").Value = Max_return
        ws.Range("P3").Value = Min_return
        ws.Range("P4").Value = Max_volume
        
        For j = 2 To WorksheetFunction.CountA(Range("I:I"))
            
            If ws.Cells(j, 11).Value = ws.Range("P2").Value Then
            ws.Range("O2").Value = ws.Cells(j, 9).Value
            
            ElseIf ws.Cells(j, 11).Value = ws.Range("P3").Value Then
            ws.Range("O3").Value = ws.Cells(j, 9).Value
            
            ElseIf ws.Cells(j, 12).Value = ws.Range("P4").Value Then
            ws.Range("O4").Value = ws.Cells(j, 9).Value
            
            End If
        
        Next j
       
    Next ws
    
    
End Sub