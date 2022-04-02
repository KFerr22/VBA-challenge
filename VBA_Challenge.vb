Sub StockLoop()

'Loop through sheets
    For Each ws In Worksheets

       'Set variable for ticker name
        Dim tickername As String
    
        'Set varable for ticker volume & ticker name in sumarry table
        Dim tickervolume As Double
        tickervolume = 0

        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Set variable for Open price, Close price, yearly change & percent change
        Dim openprice As Double
        openprice = ws.Cells(2, 3).Value
        
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double

        'Label the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Find last row in first column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through rows

        For i = 2 To lastrow

            'Where next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set ticker name
              tickername = ws.Cells(i, 1).Value

              'Add volume of ticker
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'Print ticker name in summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'Print volume of ticker for each ticker in summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'Obtain close price
              closeprice = ws.Cells(i, 6).Value

              'Calculate yearly change
               yearlychange = (closeprice - openprice)
              
              'Print yearly change for each ticker in summary table
              ws.Range("J" & summary_ticker_row).Value = yearlychange

              'Calculate Percent change
                If openprice = 0 Then
                    percentchange = 0
                Else
                    percentchange = yearlychange / openprice
                End If

              'Print yearly change for each ticker in summary table
              ws.Range("K" & summary_ticker_row).Value = percentchange
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset row counter & add one to summary ticker row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of ticker
              tickervolume = 0

              'Reset open price
              openprice = ws.Cells(i + 1, 3)
            
            Else
              
               'Add volume of ticker
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Formatting yearly change (<0 = red & >0 = green)

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To lastrow_summary_table
            
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
            
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
        Next i

    'Greatest % & Total Stock Volume

        For i = 2 To lastrow_summary_table
        
        'Find max percent change
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"

        'Find min percent change
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
            
        'Find max volume
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
        End If
        
        Next i
    
    Next ws
    
End Sub
