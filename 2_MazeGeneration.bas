Sub MazeGeneration()

Dim i, j, s, t, u, n As Integer
n = 10

Do
    n = n + 22
    i = 0

    Do
        j = 0
        i = i + 2

        Do
            j = j + 2
            s = Fix(Rnd * 4) + 1
            'Pointing cell:((j + 29, i + 1)

            If s = 1 Then
                Cells(j + n, i).Value = "0"
            Else
                If s = 2 Then
                    Cells(j + n - 1, i + 1).Value = "0"
                Else
                    If s = 3 Then
                        Cells(j + n, i + 2).Value = "0"
                    Else
                        Cells(j + n + 1, i + 1).Value = "0"
                    End If
                End If
            End If

            t = Fix(Rnd * 4) + 1
            'Pointing cell:((j + 29, i + 1)

            If t = 1 Then
                Cells(j + n, i).Value = "0"
            Else
                If t = 2 Then
                    Cells(j + n - 1, i + 1).Value = "0"
                Else
                    If t = 3 Then
                        Cells(j + n, i + 2).Value = "0"
                    Else
                        Cells(j + n + 1, i + 1).Value = "0"
                    End If
                End If
            End If

            u = Fix(Rnd * 4) + 1
            'Pointing cell:((j + 29, i + 1)

            If u = 1 Then
                Cells(j + n, i).Value = "0"
            Else
                If u = 2 Then
                    Cells(j + n - 1, i + 1).Value = "0"
                Else
                    If u = 3 Then
                        Cells(j + n, i + 2).Value = "0"
                    Else
                        Cells(j + n + 1, i + 1).Value = "0"
                    End If
                End If
            End If

        Loop While j < 18
    Loop While i < 18
Loop While n < 209

End Sub