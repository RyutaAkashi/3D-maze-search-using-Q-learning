Sub InputQInitialValue()

Dim i, j, k, m, n, x As Integer
i = 0
m = Cells(249, 20).Value

Do
    j = 0
    i = i + 1

    Do
        j = j + 1
        k = Cells(j + 5, i + 25).Value

        If k = 0 Then
            Cells(j + 5, i + 33).Value = -100
        Else
            Cells(j + 5, i + 33).Value = Fix(Rnd * 9) + 1
        End If

    Loop While j < m
Loop While i < 6

End Sub