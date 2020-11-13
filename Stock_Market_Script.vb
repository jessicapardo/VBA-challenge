Sub Stock_Market_Analysis()

'Set variables
'---------------------------------------

Dim ticker_symbol As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim total_vol As Double
Dim ws As Worksheet


'Create a loop through all worksheets
'---------------------------------------
For Each ws In Worksheets

    'Placing the headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Autofit table columns
    ws.Columns("I:L").EntireColumn.AutoFit
    
'Initial Values
'---------------------------------------
    
    'Set variable for total rows
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim rowcount As Long
    rowcount = 2

    'Loop for search
    For i = 2 To lastrow
        
        'Conditional to determine year open price
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            year_open = ws.Cells(i, 3).Value
        End If
        
        'Total stock volume for the year
        total_vol = total_vol + ws.Cells(i, 7)

        'Conditional to determine if the ticker symbol is changing
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'Ticker symbol
            ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

            'Total stock
             ws.Cells(rowcount, 12).Value = total_vol

            'Define year end price
            year_close = ws.Cells(i, 6).Value

            'Calculate the price change
             year_change = year_close - year_open
             ws.Cells(rowcount, 10).Value = year_change
             
            'Calculate the percent change for the year
            If year_open = 0 And year_close = 0 Then
                percent_change = 0
                ws.Cells(rowcount, 11).Value = percent_change
                'Formating
                ws.Cells(rowcount, 11).NumberFormat = "0.00%"
            ElseIf year_open = 0 Then
                Dim percent_change_NA As String
            percent_change_NA = "New Stock"
            ws.Cells(rowcount, 11).Value = percent_change
 
        Else
        percent_change = year_change / year_open
        ws.Cells(rowcount, 11).Value = percent_change
        ws.Cells(rowcount, 11).NumberFormat = "0.00%"
        End If
        
        'Highlight positive change in green and negative change in red
        If year_change >= 0 Then
        ws.Cells(rowcount, 10).Interior.ColorIndex = 4
        
        Else
        
        ws.Cells(rowcount, 10).Interior.ColorIndex = 3
        
        End If
 
       'Move it to the next empty row
       rowcount = rowcount + 1
 
       'Reset values
        total_vol = 0
        year_open = 0
        year_close = 0
        year_change = 0
        percent_change = 0

        End If
        
                
    Next i
'------------------------------------------------------------------------------
'BONUS:"Greatest % increase", "Greatest % decrease" and "Greatest total volume"
'------------------------------------------------------------------------------
'Set variables
'---------------------------------------
    Dim max_value As Double
    Dim min_value As Double
    Dim max_vol As Double
    Dim lastrow_b As Double


'Placing the headers
'---------------------------------------
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
   
      
    lastrow_b = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
    'Determining greatest % increase
    max_value = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17).Value = max_value
    
     For j = 2 To lastrow_b
        If ws.Cells(j, 11).Value = max_value Then
        ws.Cells(2, 16) = ws.Cells(j, 9).Value
        End If
     Next j
    
    'Determining greatest % decrease
    min_value = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17).Value = min_value
    
     For k = 2 To lastrow_b
        If ws.Cells(k, 11).Value = min_value Then
        ws.Cells(3, 16) = ws.Cells(k, 9).Value
        End If
     Next k
    
    'Determining greatest Total Volume
    max_vol = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17).Value = max_vol
        
     For l = 2 To lastrow_b
        If ws.Cells(l, 12).Value = max_vol Then
        ws.Cells(4, 16) = ws.Cells(l, 9).Value
        End If
     Next l
  
'Formating
    ws.Columns("O:Q").EntireColumn.AutoFit
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
Next ws

End Sub