
Sub tickerloop()

'Looping through all the sheets.
    For Each ws In Worksheets
    
        Dim ticker_name As String
    
        Dim ticker_volume As Double
        ticker_volume = 0

        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
    
        Dim opening_price As Double
        
        'Set initial opening_price.
        opening_price = ws.Cells(2, 3).Value
        
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percentage_change As Double

        'Labelling the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Counting the number of rows in the first column.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Looping through the rows by the ticker names
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              ticker_name = ws.Cells(i, 1).Value
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value
              ws.Range("I" & summary_ticker_row).Value = ticker_name
              ws.Range("L" & summary_ticker_row).Value = ticker_volume
              closing_price = ws.Cells(i, 6).Value

              'Calculating the yearly change
               yearly_change = (closing_price - opening_price)
        
              ws.Range("J" & summary_ticker_row).Value = yearly_change
              
                If opening_price = 0 Then
                    percentage_change = 0
                
                Else
                    percentage_change = yearly_change / opening_price
                
                End If

              ws.Range("K" & summary_ticker_row).Value = percentage_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter and Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              ticker_volume = 0

              'Reset the opening price
              opening_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To lastrow_summary_table
        
            'Find the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub