Attribute VB_Name = "Module1"
Sub TickerLoop()

    ' Declare data types
    Dim ws As Worksheet
    Dim r, i, j As Integer
    
    ' Loop all functions over each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Create headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N1").Value = "Greatest % Increase"
        ws.Range("O1").Value = "Greatest % Decrease"
        ws.Range("P1").Value = "Greatest Total Volume"
        ws.Columns("I:P").AutoFit
        
        ' Initiate values and calcualte last row
        j = 0
        r = ws.Cells(Rows.Count, "A").End(xlUp).Row
        rs = 2
        
        ' Loop over all rows of ticker data
        For i = 2 To r
            ' Find end of ticker range
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Assign ticker name to a variable
                t = ws.Cells(i, 1).Value
                ' Assign total stock volume to a variable
                v = ws.Cells(i, 7).Value
                ' Insert ticker name
                ws.Range("I" & j + 2).Value = t
                ' Insert total stock volume
                ws.Range("L" & j + 2).Value = v
                ' Assign closing ticker value to a variable
                c = ws.Cells(i, 6).Value
                ' Assign opening ticker value to a variable
                o = ws.Cells(rs, 3).Value
                ' Calculate yearly change and assign to variable
                yc = c - o
                ' Insert yearly change
                ws.Range("J" & j + 2).Value = yc
                ' Calculate percent chage and assign to variable
                pc = yc / o
                ' Format percent change greater than or equal to 0
                If pc >= 0 Then
                    ws.Range("K" & j + 2).Value = pc
                    ws.Range("K" & j + 2).NumberFormat = "0.00%"
                    ws.Range("K" & j + 2).Interior.ColorIndex = 4
                ' Format percent change less than 0
                Else
                    ws.Range("K" & j + 2).Value = pc
                    ws.Range("K" & j + 2).NumberFormat = "0.00%"
                    ws.Range("K" & j + 2).Interior.ColorIndex = 3
                End If
                ' Increment row in ticker summary table
                rs = i + 1
                j = j + 1
            End If
        Next i
        
        ' Calcualte last row and assign to a variable
        rg = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        ' Loop through all rows in percent change
        For i = 2 To rg
            ' Find maximum percent change and format
            ws.Cells(2, 14).Value = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Cells(2, 14).NumberFormat = "0.00%"
            ' Find minimum percent change and format
            ws.Cells(2, 15).Value = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Cells(2, 15).NumberFormat = "0.00%"
            ' Find maximum total stock volume
            ws.Cells(2, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))
        Next i
    Next ws

End Sub
