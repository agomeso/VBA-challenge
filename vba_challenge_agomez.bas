Attribute VB_Name = "Module1"
Sub Investment()
'Loop through all sheets
For Each ws In Worksheets

'Create headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Year Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'Cells(1, 13).Value = "Close"
    'Cells(1, 14).Value = "Open"

'Set up a variable to hold the ticker simbol
    Dim Ticker As String
    
'Set up a variable to hold the opening price
    Dim OP As Double
    
'Set up a variable to hold the closing price
    Dim CP As Double
    
'Set up a variable to hold the total volume
    Dim Total_Vol As LongLong
    Total_Vol = 0
    
'Set up a variable to hold the yearly change
    Dim Yrly_Chg As Double
    
'Set up a variable to hold the percentage change
    Dim Pct_Chg As Double
    
'Keep track of each ticket simbol in the Summary...maybe
    Dim Sum_Row As Integer
    Sum_Row = 2
    
'Find the last row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Row counter to find open
    Dim counter As Integer
    
    counter = 0

'Loop through all ticker simbols
    For i = 2 To Lastrow
        'Grab the open price
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            counter = counter + 1
                If counter = 1 Then
        'Grab the open price
                OP = ws.Cells(i, 3).Value
        'Print opening price
                'Range("N" & Sum_Row).Value = OP
                End If
        End If
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        counter = 0
        'Set ticker name in Summary Table
       Ticker = ws.Cells(i, 1).Value
        
        'Add to the total volume
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
        'Print ticket symbol
        ws.Range("I" & Sum_Row).Value = Ticker
        
        'Print total volume
        ws.Range("L" & Sum_Row).Value = Total_Vol
        
        'Reset total
        Total_Vol = 0
        
        'Add one to the Summary Table row
        Sum_Row = Sum_Row + 1
        
        'If the cell immediately following is the same
        Else
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        'Find closing price
        CP = ws.Cells(i + 1, 6).Value
        
        'Print closing price
        'Range("M" & Sum_Row).Value = CP
        
        'Calculate the yearly change
        Yrly_Chg = CP - OP
        
        'Format Yrly_Chg
            If Yrly_Chg >= 0 Then
                ws.Range("J" & Sum_Row).Interior.ColorIndex = 4
                Else
                ws.Range("J" & Sum_Row).Interior.ColorIndex = 3
            End If
    
        'Print yearly change
        ws.Range("J" & Sum_Row).Value = Yrly_Chg
        
        'Calculate Percentage Change
        'Fix damn overflow error
        If OP = 0 Then
        Pct_Chg = 0
        Else
        Pct_Chg = (CP - OP) / OP
        End If
        
        'Print Percentage Change
        ws.Range("K" & Sum_Row).Value = Pct_Chg
        ws.Range("K" & Sum_Row).NumberFormat = "0.00%"
        
        End If
        
    Next i
    
'Table Parameters
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Declare variables for Summary Table
Dim Grt_Inc As Double
Dim Grt_Dec As Double
Dim Grt_Vol As LongLong

'Find the last row
    Dim LastSummRow As Integer
    
    LastSummRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Find greatest increase, decrease and total volume

    'Initialize Grt_Inc with 1st value on table
    Grt_Inc = ws.Cells(2, 11).Value
    Grt_Dec = ws.Cells(2, 11).Value
    Grt_Vol = ws.Cells(2, 12).Value

    For j = 2 To LastSummRow
    
        'Find greatest increase
        If ws.Cells(j, 11).Value > Grt_Inc Then
        Grt_Inc = ws.Cells(j, 11).Value
        End If
        
        'Find greatest decrease
        If ws.Cells(j, 11).Value < Grt_Dec Then
        Grt_Dec = ws.Cells(j, 11).Value
        End If
        
        'Find greatest volume
        If ws.Cells(j, 12).Value > Grt_Vol Then
        Grt_Vol = ws.Cells(j, 12).Value
        End If
    Next j
    
    'Print greatest increase, decrease and largest volume
    ws.Cells(2, 17).Value = Grt_Inc
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = Grt_Dec
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = Grt_Vol
    
    For k = 1 To LastSummRow
        'Find corresponding Ticker Symbols
        If ws.Cells(k, 11).Value = ws.Cells(2, 17).Value Then
        ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 11).Value = ws.Cells(3, 17).Value Then
        ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
        End If
        
        If ws.Cells(k, 12).Value = ws.Cells(4, 17).Value Then
        ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
        End If
        
    Next k
    
Next ws
    
End Sub
