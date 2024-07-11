Sub Ticker()

For Each ws In Worksheets

'Title the Columns and rows

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increaase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

'Copy Ticker titles to Column I, no repeats

Dim y As Double

    y = 2

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

    ws.Cells(y, 9).Value = ws.Cells(i, 1).Value

    y = y + 1

End If

Next i

Next ws

'Calculate total stock volume

For Each ws In Worksheets

lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow2

    ws.Cells(i, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(i, 9), ws.Range("G:G"))

Next i

Next ws

'Calculate quarterly change and percent change

For Each ws In Worksheets

    Dim o As Double
    Dim c As Double

    y = 2
    o = 0
    c = 0

For i = 2 To lastrow

    If ws.Cells(i, 1) = ws.Cells(i + 1, 1) And (Month(ws.Cells(i, 2).Value) = 1 And Day(ws.Cells(i, 2).Value) = 2) Or (Month(ws.Cells(i, 2).Value) = 4 And Day(ws.Cells(i, 2).Value) = 1) Or (Month(ws.Cells(i, 2).Value) = 7 And Day(ws.Cells(i, 2).Value) = 1) Or (Month(ws.Cells(i, 2).Value) = 10 And Day(ws.Cells(i, 2).Value) = 1) Then

        o = ws.Cells(i, 3).Value

        y = y + 1
    Else

        c = ws.Cells(i, 6).Value

        ws.Cells(y - 1, 10).Value = c - o

            If ws.Cells(y - 1, 10).Value > 0 Then
    
            ws.Cells(y - 1, 10).Interior.ColorIndex = 4
    
        Else
    
            ws.Cells(y - 1, 10).Interior.ColorIndex = 3
        
    End If

        ws.Cells(y - 1, 11).Value = ((c - o) / o)


    End If

Next i

ws.Range("K2:K" & lastrow2).NumberFormat = "0.00%"

Next ws

'Find greatest increase and decrease

For Each ws In Worksheets

    Dim inchange As Double
    Dim dechange As Double

    inchange = 0
    dechange = 0

For i = 2 To lastrow2

    If ws.Cells(i, 11).Value > inchange Then

    inchange = ws.Cells(i, 11).Value

    ws.Range("p2") = ws.Cells(i, 9).Value
    ws.Range("q2") = inchange

ElseIf ws.Cells(i, 11).Value < dechange Then

    dechange = ws.Cells(i, 11).Value

    ws.Range("p3") = ws.Cells(i, 9).Value
    ws.Range("q3") = dechange


End If

Next i

ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

'Find greatest total volume

For Each ws In Worksheets

    Dim volume As Double

    volume = 0

For i = 2 To lastrow2

    If ws.Cells(i, 12).Value > volume Then

    volume = ws.Cells(i, 12).Value

    ws.Range("p4") = ws.Cells(i, 9).Value
    ws.Range("q4") = volume

End If

Next i

Next ws



End Sub



