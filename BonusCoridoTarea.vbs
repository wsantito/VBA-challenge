'Macro para Bonus 

RowCount = Cells(Rows.Count, "J").End(xlUp).Row
MMin = Cells(2, 11).Value
MMax = Cells(2, 11).Value

For i = 2 To RowCount - 1
If Cells(i, 11) >= MMax Then
    MMax = Cells(i, 11).Value
    Ticker1 = Cells(i, 9)
        If Cells(i, 11) <= MMin Then
            MMin = Cells(i, 11).Value
            Ticker2 = Cells(i, 9).Value
        End If
Else
        
        If Cells(i, 11) <= MMin Then
            MMin = Cells(i, 11).Value
            Ticker2 = Cells(i, 9).Value
        End If
End If

Cells(2, 15) = Ticker1
Cells(3, 15) = Ticker2
Cells(2, 16) = MMax
Cells(3, 16) = MMin


Next i

'Bonus Sum


RowCount = Cells(Rows.Count, "J").End(xlUp).Row
SMax = Cells(2, 12).Value

For i = 2 To RowCount - 1
If Cells(i, 12) >= SMax Then
    SMax = Cells(i, 12).Value
    Ticker3 = Cells(i, 9)

End If

Cells(4, 15) = Ticker3
Cells(4, 16) = SMax


Next i

Cells(1, 1).Select

End Sub

