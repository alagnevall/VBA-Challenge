Attribute VB_Name = "Module1"
Sub StockmarketChanges()

Dim Ticker As String
Dim YearlyOpen As Double
Dim Percent As Double
Dim Volume As Double
Dim LineCount As Integer
Dim LastRow As Long
Dim rng As Range
Dim rng2 As Range

LineCount = 2
Volume = 0



'Create headers for each worksheet starting in column I; the following Ticker, Yearly change, Percent change, Total stock volume


' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    'Find last row for each worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'run loop for each row
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' If the next row is not the same
            ws.Cells(LineCount, 9).Value = ws.Cells(i, 1).Value
            'sum up the total stock volume for each ticker
            ws.Cells(LineCount, 12).Value = ws.Cells(i, 7).Value + Volume
            'To find the yearly change; find closing cost (minus) opening cost
            ws.Cells(LineCount, 10).Value = ws.Cells(i, 6).Value - YearlyOpen
            
            If YearlyOpen = 0 Then
                ws.Cells(LineCount, 11).Value = 0
            Else
                'To find the percent change; yearly change divided opening cost
                ws.Cells(LineCount, 11).Value = FormatPercent((ws.Cells(i, 6).Value - YearlyOpen) / YearlyOpen, 2)
            End If
       
            ' format positive yearly change to green
            If ws.Cells(LineCount, 10).Value > 0 Then
                ws.Cells(LineCount, 10).Interior.ColorIndex = 4
            Else
                'format negative yearly change to red
                ws.Cells(LineCount, 10).Interior.ColorIndex = 3
            End If
               
            ' Make sure data is populating on a new line and not overwriting
            LineCount = LineCount + 1
       
            ' Reset volume to 0 once it finds the end value
            Volume = 0
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                YearlyOpen = ws.Cells(i, 3).Value
                Volume = Volume + ws.Cells(i, 7).Value
           
            Else
                ' Do this if it is NOT the last row of data(summing values)
                Volume = Volume + ws.Cells(i, 7).Value
        End If
    Next i
   
    'Reset line count before moving to next sheet
    LineCount = 2

    ' --------------------------------------------
    'BONUS CHALLENGE
    ' --------------------------------------------

    'Add headers and rows

    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("o1") = "Ticker"
    ws.Range("p1") = "Value"
    
    'define another last row since I am looking at a different column

    SumRow = ws.Cells(Rows.Count, 11).End(xlUp).Row


    'set ranges to look at for increase, decrease and volume
    Set rng = ws.Range("K2:K" & SumRow)
    Set rng2 = ws.Range("L2:L" & SumRow)
    'find the max increase using worksheet functions
    ws.Range("p2").Value = FormatPercent(Application.WorksheetFunction.Max(rng), 2)
    ws.Range("o2").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(rng), rng, 0) + 1, 9).Value
    'find the max decrease using worksheet functions
    ws.Range("p3").Value = FormatPercent(Application.WorksheetFunction.Min(rng), 2)
    ws.Range("o3").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(rng), rng, 0) + 1, 9).Value
    'find the max volume using worksheet functions
    ws.Range("p4").Value = Application.WorksheetFunction.Max(rng2)
    ws.Range("o4").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(rng2), rng2, 0) + 1, 9).Value

    
'Stop looping through worksheets
Next ws

End Sub
