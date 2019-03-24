Attribute VB_Name = "Module2"
Sub StockData()
    For Each ws In Worksheets
        ws.Activate
            Dim Ticker As String
            Dim Total_Stock_Volume As Double
            Total_Stock_Volume = 0
 
            Dim Table_Row As Integer
            Table_Row = 2
            For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Ticker = Cells(i, 1).Value
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                    Range("I" & Table_Row).Value = Ticker
                    Range("J" & Table_Row).Value = Total_Stock_Volume
                    Table_Row = Table_Row + 1
                    Total_Stock_Volume = 0
                Else
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
        

                End If
    
            Next i
    Next ws
    

End Sub

