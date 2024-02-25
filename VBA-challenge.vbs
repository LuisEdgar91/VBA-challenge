Sub Stocks():

    'Loop through all sheets
    For Each ws In Worksheets

        'Set an initial variable for holding the ticker symbol
        Dim ticker As String
              
        'Set a variable for counting the ticker symbol
        Dim ticker_count As Integer
        ticker_count = 0
              
        'Keep track of the location for each ticker symbol in the summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        'Use last_row as variable instead of a specific row number
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
             
        'Add headers in summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
             
        'Loop through all ticker symbols
        For i = 2 To last_row
                
            'Check if is still the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker name
                ticker = ws.Cells(i, 1).Value
                
                'Set the ticker opening price
                Dim opening_price As Double
                opening_price = ws.Cells(i - ticker_count, 3).Value
                
                'Set the ticker closing price
                Dim closing_price As Double
                closing_price = ws.Cells(i, 6).Value
                                
                'Add to the ticker volume
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                
                'Print the ticker symbol in the summary table
                ws.Range("I" & summary_table_row).Value = ticker
                
                'Define yearly change
                Dim yearly_change As Double
                yearly_change = closing_price - opening_price
                
                'Print the ticker yearly change in the summary table
                ws.Range("J" & summary_table_row).Value = yearly_change
                
                'Add color
                    If yearly_change < 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                                                   
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                                                   
                    End If
                                               
                'Define ticker percentage
                Dim ticker_percentage As Double
                ticker_percentage = closing_price / opening_price - 1
                                               
                'Print the ticker percentage change in the summary table
                ws.Range("K" & summary_table_row).Value = ticker_percentage
                                                               
                'Add percentage property
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                                                               
                'Print the ticker total amount in the summary table
                ws.Range("L" & summary_table_row).Value = ticker_volume
                
                'Jump one row in the summary table
                summary_table_row = summary_table_row + 1
                
                'Reset the ticker total volume and the ticker count
                ticker_volume = 0
                ticker_count = 0
                                            
            'If the cell immediately following a row is the same ticker...
            Else
                
                'Add to the ticker volume
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                        
                'Add to the ticker symbol
                ticker_count = ticker_count + 1
                        
            End If
        
        Next i
        
'----------------------------------------------- FIND THE HIGHEST, LOWEST AND MAX VOLUME -----------------------------------------------
        
        Dim last_row_summary As Long
        
        Dim tickerColumn As Range
        Dim percentageColumn As Range
        Dim volumeColumn As Range
        
        Dim highest_percentage As Double
        Dim lowest_percentage As Double
        Dim greatest_volume As Double
        
        Dim highest_percentage_ticker As String
        Dim lowest_percentage_ticker As String
        Dim greatest_volume_ticker As String
        
        Dim x As Long
             
        'Find the last row with data in the summary table
        last_row_summary = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        'Set the ranges for ticker, yearly change and volume columns
        Set tickerColumn = ws.Range("I2:I" & last_row_summary)
        Set percentageColumn = ws.Range("K2:K" & last_row_summary)
        Set volumeColumn = ws.Range("L2:L" & last_row_summary)
        
        'Initialize highest percentage, lowest percentage and greatest volume as the first value
        highest_percentage = percentageColumn.Cells(1).Value
        lowest_percentage = percentageColumn.Cells(1).Value
        greatest_volume = volumeColumn.Cells(1).Value
        
        highest_percentage_ticker = tickerColumn.Cells(1).Value
        lowest_percentage_ticker = tickerColumn.Cells(1).Value
        greatest_volume_ticker = tickerColumn.Cells(1).Value
        
            'Loop through each cell in the summary table starting from the second row
            For x = 2 To tickerColumn.Rows.Count
                
                'Check if the current value is greater than the current highest percentage
                If percentageColumn.Cells(x).Value > highest_percentage Then
                    
                    'Update highest percentage and highest percentage ticker if the current value is greater
                    highest_percentage = percentageColumn.Cells(x).Value
                    highest_percentage_ticker = tickerColumn.Cells(x).Value
                
                End If
                
                'Check if the current value is less than the current lowest percentage
                If percentageColumn.Cells(x).Value < lowest_percentage Then
                    
                    'Update lowest percentage and lowest percentage ticker if the current value is less
                    lowest_percentage = percentageColumn.Cells(x).Value
                    lowest_percentage_ticker = tickerColumn.Cells(x).Value
                
                End If
            
                'Check if the current value is greater than the current greatest volume
                If volumeColumn.Cells(x).Value > greatest_volume Then
                    
                    'Update greatest volume and greatest volume ticker if the current value is greater
                    greatest_volume = volumeColumn.Cells(x).Value
                    greatest_volume_ticker = tickerColumn.Cells(x).Value
            
                End If
            
            Next x
        
        'Print the highest value and corresponding ticker in column P and Q respectively
        ws.Range("P2").Value = highest_percentage_ticker
        ws.Range("Q2").Value = highest_percentage
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Print the lowest value and corresponding ticker in column P and Q respectively
        ws.Range("P3").Value = lowest_percentage_ticker
        ws.Range("Q3").Value = lowest_percentage
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Print the greatest volume and corresponding ticker in column P and Q respectively
        ws.Range("P4").Value = greatest_volume_ticker
        ws.Range("Q4").Value = greatest_volume
        
    Next ws

End Sub
