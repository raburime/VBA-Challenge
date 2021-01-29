Sub stock()
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Line = 2
    Change = 0
    Total = 0
    Start = 0
    
    For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
        Change = Change + Cells(i, "F") - Cells(i, "C")
        Total = Total + Cells(i, "G")
        If Start = 0 Then
            Start = i
        End If
        
        If Cells(i, "A") <> Cells(i + 1, "A") Then
            Cells(Line, "I") = Cells(i, "A")
            Cells(Line, "J") = Change
            Cells(Line, "L") = Total
            Cells(Line, "K") = Round((Cells(Line, "J").Value / Cells(Start, "C").Value) * 100, 2) & "%"
            Change = 0
            Total = 0
            Start = 0
            
            Line = Line + 1
        End If
    Next i
End Sub

    
