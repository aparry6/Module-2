Sub StockAnalysis()

    Dim ws As Worksheet
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim last_row As Long
    Dim summary_table_row As Integer
    
    For Each ws In Worksheets
        
        'initialize variables
        ticker = ""
        year_open = 0
        year_close = 0
        yearly_change = 0
        percent_change = 0
        total_volume = 0
        summary_table_row = 2
        
        'set up summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'get last row in sheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'loop through all rows in sheet
        For i = 2 To last_row
        
            'check if current row is first row for current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                'set year open value for new ticker
                year_open = ws.Cells(i, 3).Value
                
                'set ticker value for new ticker
                ticker = ws.Cells(i, 1).Value
                
            End If
            
            'add volume to total_volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'check if current row is last row for current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'set year close value for current ticker
                year_close = ws.Cells(i, 6).Value
                
                'calculate yearly_change and percent_change
                yearly_change = year_close - year_open
                If year_open <> 0 Then
                    percent_change = yearly_change / year_open
                End If
                
                'populate summary table with ticker info
                ws.Range("I" & summary_table_row).Value = ticker
                ws.Range("J" & summary_table_row).Value = yearly_change
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("L" & summary_table_row).Value = total_volume
                
                'format percent_change as percentage
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                
                'reset variables for next ticker
                year_open = 0
                year_close = 0
                yearly_change = 0
                percent_change = 0
                total_volume = 0
                summary_table_row = summary_table_row + 1
                
            End If
            
        Next i
        
        'get last row in summary table
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'loop through all rows in summary table
        For i = 2 To last_row
        
            'highlight positive yearly change in green and negative yearly change in red
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        Next i
    
    Next ws
    
End Sub

Sub Summary()
    Dim last_row As Long
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    
    'Set initial values for max_increase and max_decrease
    max_increase = Cells(2, 11).Value
    max_decrease = Cells(2, 11).Value
    max_volume = Cells(2, 12).Value
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through each row and compare values to determine max_increase, max_decrease, and max_volume
    For i = 2 To last_row
        If Cells(i, 11).Value > max_increase Then
            max_increase = Cells(i, 11).Value
            max_increase_ticker = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value < max_decrease Then
            max_decrease = Cells(i, 11).Value
            max_decrease_ticker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > max_volume Then
            max_volume = Cells(i, 12).Value
            max_volume_ticker = Cells(i, 9).Value
        End If
    Next i
    
    'Output results to corresponding cells
    Range("P2").Value = max_increase_ticker
    Range("Q2").Value = max_increase
    Range("P3").Value = max_decrease_ticker
    Range("Q3").Value = max_decrease
    Range("P4").Value = max_volume_ticker
    Range("Q4").Value = max_volume
End Sub

