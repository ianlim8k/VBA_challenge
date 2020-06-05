Attribute VB_Name = "Module1"
Sub stock_market_analysis()

     
    Dim current_sheet As Worksheet
    Dim number_of_sheets As Integer
    Dim Rowcount As Double
    
    Dim ticker_row As Integer
    Dim ticker_symbol As String
    
    Dim ticker_symbol_start As Boolean
    Dim ticker_symbol_end As Boolean
    
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
               
    Dim s_lastrow As Double
    Dim s_row As Long
        
    Dim s_percentchange As Double
    Dim s_volume As Double
   
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_vol As String
        
    Dim g_inc As Double
    Dim g_dec As Double
    Dim g_vol As Double
         
              
        'Count the number of worksheets
        number_of_sheets = ActiveWorkbook.Worksheets.Count

        For w = 1 To number_of_sheets
             
        'Set worksheet to current worksheet in the workbook
        Set current_sheet = ActiveWorkbook.Worksheets(w)
                             
        'Get the number of rows of data
        Rowcount = current_sheet.UsedRange.Rows.Count
                    
        'Create summary column
    
        current_sheet.Range("J1").Value = "Ticker"
        current_sheet.Range("K1").Value = "Yearly Change"
        current_sheet.Range("L1").Value = "Percent Change"
        current_sheet.Range("M1").Value = "Total Stock Volume"
        current_sheet.Range("Q1").Value = "Ticker"
        current_sheet.Range("R1").Value = "Value"
        current_sheet.Range("P2").Value = "Greatest % increase"
        current_sheet.Range("P3").Value = "Greatest % decrease"
        current_sheet.Range("P4").Value = "Greatest Total increase"
               
        
        'Output summary from row 2
        ticker_row = 2
        
         'Start of ticker symbol
        ticker_symbol_start = True
        
        
        For i = 2 To Rowcount
                   
            ticker_symbol = current_sheet.Cells(i, 1).Value
                          
            'Get the opening price for start of year
                   
            If ticker_symbol_start = True Then
                year_open_price = current_sheet.Cells(i, 3).Value
                ticker_symbol_start = False
            End If
            
            'accumulate the total volume
            total_volume = total_volume + current_sheet.Cells(i, 7).Value
            
            'Print to summary row when ticker symbol changes
            
            If ticker_symbol <> current_sheet.Cells(i + 1, 1).Value Then
            
                'output ticket symbol to ticker_symbol column
                
                current_sheet.Cells(ticker_row, 10) = ticker_symbol
              
                'Get the year end closing price
                year_close_price = current_sheet.Cells(i, 6)
                
                'Calculate price change from opening price to closing price for given year
                yearly_change = year_close_price - year_open_price
                
                'Calculate percent change from opening price to closing price for given year
                If year_open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / year_open_price
                End If
                
                                                                     
                'output to the summary row
                current_sheet.Cells(ticker_row, 11).Value = yearly_change
                current_sheet.Cells(ticker_row, 12).Value = percent_change
                current_sheet.Cells(ticker_row, 13).Value = total_volume
                
                'change format to percentage
                current_sheet.Range("L:L").NumberFormat = "0.00%"
                
                'add conditionals formatting
                If yearly_change > 0 Then
                    current_sheet.Cells(ticker_row, 11).Interior.ColorIndex = 4
                ElseIf yearly_change < 0 Then
                    current_sheet.Cells(ticker_row, 11).Interior.ColorIndex = 3
                Else
                    current_sheet.Cells(ticker_row, 11).Interior.ColorIndex = 2
                End If
                
                'increment row to output to next summary row for next ticker symbol
                ticker_row = ticker_row + 1
                
                'reset ticker symbol start
                ticker_symbol_start = True
                
                'reset total volume
                total_volume = 0
                        
            End If
            
        Next i
        
        
        'Find the greatest % increase from summary table
   
        'Get the last row of data in summary table
         s_lastrow = current_sheet.Range("J" & Rows.Count).End(xlUp).Row
               
            For j = 2 To s_lastrow
         
                s_percentchange = current_sheet.Cells(j, 12).Value
                           
                'Set initial value
                
                If j = 2 Then
                    g_inc = current_sheet.Cells(j, 12).Value
                End If
                         
                'Get the maximum value and set row number to retrieve ticker symbol
                If s_percentchange > g_inc Then
                    g_inc = s_percentchange
                    s_row = j
                End If
                        
            Next j
            
            'Get ticker symbol
            max_ticker = current_sheet.Cells(s_row, 10).Value
            
            'Ouput the greatest % increase to summary row
            current_sheet.Cells(2, 17).Value = max_ticker
            current_sheet.Cells(2, 18).Value = g_inc
        
        
        'Find the greatest % decrease from summary table
        
            For h = 2 To s_lastrow
             
                s_percentchange = current_sheet.Cells(h, 12).Value
                
                'Set initial value
                If h = 2 Then
                    g_dec = current_sheet.Cells(h, 12).Value
                End If
                        
                'Get the minimum value and set row number to retrieve ticker symbol
                If s_percentchange < g_dec Then
                    g_dec = s_percentchange
                    s_row = h
                End If
                        
            Next h

            'Get ticker symbol
            min_ticker = current_sheet.Cells(s_row, 10).Value
            
            'Ouput the greatest % decrease to summary row
            current_sheet.Cells(3, 17).Value = min_ticker
            current_sheet.Cells(3, 18).Value = g_dec
        
            'change format to percentage
            current_sheet.Range("R2:R3").NumberFormat = "0.00%"
        
        'Find the greatest Total Volume from summary table
        
            For m = 2 To s_lastrow
            
                s_volume = current_sheet.Cells(m, 13).Value
             
                'Set initial value
                If m = 2 Then
                    g_vol = current_sheet.Cells(m, 13).Value
                End If
                
                'Get the greatest volume and set row number to retrieve ticker symbol
                If s_volume > g_vol Then
                   g_vol = s_volume
                   s_row = m
                End If
                
            Next m
                
            'Get ticker symbol
            max_vol = current_sheet.Cells(s_row, 10).Value
            
            'Ouput the greatest volume to summary row
            current_sheet.Cells(4, 17).Value = max_vol
            current_sheet.Cells(4, 18).Value = g_vol
            
        
         'Adjust the column width to fit data for the summary columns
         current_sheet.Range("J:J").ColumnWidth = 10
         current_sheet.Range("K:M").Columns.AutoFit
         current_sheet.Range("P:P").ColumnWidth = 23
         current_sheet.Range("Q:Q").ColumnWidth = 10
         current_sheet.Range("R:R").Columns.AutoFit
        
      Next w
    
       

End Sub


 

