Sub ShowRoute()

Dim i, j, k, m, n, x As Integer
m = Cells(5807, 26)
k = 0

Do
    n = 10
    k = k + 1

    Do
        n = n + 22
        i = 0

        Do
            j = 0
            i = i + 1

            Do
                j = j + 1
                If Cells(j + n, i + 1).Value = Cells(5807, k + 26).Value Then
                    Cells(j + n, i + 1).Value = -1
                Else
                End If

            Loop While j < 19
        Loop While i < 19
    Loop While n < 209
Loop While k < m - 1

End Sub
