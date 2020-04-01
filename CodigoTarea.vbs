Sub Charges2()
Dim i, t As Integer
Dim RowCount, RowAdd, RowAdd2 As Integer
Dim Max, Min, OP, CL, Sum As Double
Dim Ticker1, Ticker2, Ticker3 As String


Max = 0
Min = 0
OP = 0
CL = 0
Sum = 0


RowCount = Cells(Rows.Count, "A").End(xlUp).Row



For i = 2 To RowCount -1

RowAdd = Cells(Rows.Count, "I").End(xlUp).Row
Cells(i, 1).Select
If Cells(i, 1) <> Cells(i - 1, 1) Then
            Max = Cells(i, 2).Value
            Min = Cells(i, 2).Value
            OP = Cells(i, 3).Value
            CL = Cells(i, 6).Value
            Sum = Cells(i, 7).Value
    Cells(RowAdd + 1, 9).Value = Cells(i, 1).Value
    Cells(RowAdd + 1, 10) = OP - CL
    Cells(RowAdd + 1, 11) = 1 - OP / CL
Else
        If Cells(i, 2) > Max Then
            Max = Cells(i, 2).Value
            CL = Cells(i, 6).Value
                If Cells(i, 2) < Min Then
                    Min = Cells(i, 2).Value
                    OP = Cells(i, 3).Value
                End If
        Else
                If Cells(i, 2) < Min Then
                    Min = Cells(i, 2).Value
                    OP = Cells(i, 3).Value
                End If
        End If
End If

                RowAdd2 = Cells(Rows.Count, "J").End(xlUp).Row
                 Cells(RowAdd2, 10) = OP - CL
                 Cells(RowAdd2, 11) = OP / CL
                 Cells(RowAdd2, 12) = Sum + Cells(i, 7).Value


Next i

Range("I1").Select


End Sub

'Macro para colores si quiere juntarse Borrar Sub, Dim y End Sub

Sub Colores()
Dim RowCount As Integer
 
 


RowCount = Cells(Rows.Count, "J").End(xlUp).Row

For i = 2 To RowCount - 1
If Cells(i, 10) = 0 Then
    Cells(i, 10).Interior.ColorIndex = 6
Else
    If Cells(i, 10) > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    Else
    Cells(i, 10).Interior.ColorIndex = 3
    End If
End If

Next i

End Sub

Sub Bonus()
 Dim MMax, MMin, SMax As Double
 Dim RowCount As Integer

