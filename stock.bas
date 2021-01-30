Sub stock()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Line = 2
        Change = 0
        Total = 0
        Start = 0
        percent = 0

        For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
            Change = Change + Cells(i, "F") - Cells(i, "C")
            Total = Total + Cells(i, "G")
            If Start = 0 Then
                Start = i
            End If
            
            If Cells(i, "A") <> Cells(i + 1, "A") Then
                Cells(Line, "I") = Cells(i, "A")
                Cells(Line, "J") = Change
                If Change > 0 Then
                    Cells(Line, "J").Interior.ColorIndex = 10
                ElseIf Change < 0 Then
                    Cells(Line, "J").Interior.ColorIndex = 3
                End If
                Cells(Line, "L") = Total
                percent = Round((Cells(Line, "J").Value / Cells(Start, "C").Value) * 100, 2) 
                Cells(Line, "K") = percent & "%"
                Change = 0
                Total = 0
                Start = 0
                
                Line = Line + 1
            End If
        Next i
    Next ws
End Sub

    

