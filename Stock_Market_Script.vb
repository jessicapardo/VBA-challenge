Sub Stock_Market_Analysis()

'Set initial variables
'---------------------------------------

Dim ticker_symbol As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim total_vol As Double
Dim rowcount As Long
Dim lastrow As Long
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
    
    
    
    MsgBox ws.Name
    

Next ws


End Sub