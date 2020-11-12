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
    '------------------------------------
    
    'Set variable for total rows
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim rowcount As Long
    rowcount = 2

    'Loop for search
    For i = 2 To lastrow
        
        'Conditional to determine year open price
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
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

'Display Woersheet Name
MsgBox ws.Name
    
Next ws


End Sub