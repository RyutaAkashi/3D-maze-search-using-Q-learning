Sub InputTransition()

Dim i, j, k, m, n, x As Integer
n = 10

Do
    n = n + 22
    i = 0

    Do
        j = 0
        i = i + 1

        Do
            j = j + 1
            k = Cells(j + n, i + 1).Value

            If k = 0 Then
            Else
                Cells(k + 5, 26).Value = Cells(j + n - 1, i + 1).Value
                Cells(k + 5, 27).Value = Cells(j + n + 1, i + 1).Value
                Cells(k + 5, 28).Value = Cells(j + n, i).Value
                Cells(k + 5, 29).Value = Cells(j + n, i + 2).Value
                Cells(k + 5, 30).Value = Cells(j + n + 22, i + 1).Value
                Cells(k + 5, 31).Value = Cells(j + n - 22, i + 1).Value
            End If

        Loop While j < 19
    Loop While i < 19
Loop While n < 209

m = Cells(249, 20).Value
x = 1000
Cells(m + 5, 26) = 0
Cells(m + 5, 27) = 0
Cells(m + 5, 28) = 0
Cells(m + 5, 29) = 0
Cells(m + 5, 30) = 0
Cells(m + 5, 31) = 0

End Sub