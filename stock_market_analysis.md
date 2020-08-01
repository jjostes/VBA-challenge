Sub MultiYearJJ():

    For Each ws In Worksheets

        'Creating core variables before For loops
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        Dim ticker_name As String
        
        Dim open_value As Double
        Dim close_value As Double
        
        Dim yearly_change As Double
        Dim percent_change As Double
        
        Dim stock_volume As Double
        stock_volume = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                '#####TICKER##############################################
                ticker_name = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                
                '#####STOCKVOLUME########################################
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = stock_volume
      
                stock_volume = 0
                
                Summary_Table_Row = Summary_Table_Row + 1

            Else

                stock_volume = stock_volume + ws.Cells(i, 7).Value

            End If
        Next i
        
        Summary_Table_Row = 2
        
        '#####YEARLY AND PERCENT CHANGE###############################
        For i = 2 To LastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_value = ws.Cells(i, 3).Value
                
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                close_value = ws.Cells(i, 6).Value
                
            yearly_change = close_value - open_value
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
            If open_value = 0 Then
                ws.Range("K" & Summary_Table_Row).Value = "N/A"
            Else
            percent_change = (close_value - open_value) / open_value
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.0%"
            
            End If
            
            yearly_change = 0
            percent_change = 0
            
            Summary_Table_Row = Summary_Table_Row + 1
                
            End If
        Next i
        
        '#####CONDITIONAL FORMAT######################################
        For i = 2 To LastRow
        
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        Next i
        
    Next ws
        

End Sub




