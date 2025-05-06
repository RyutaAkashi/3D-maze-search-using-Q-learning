Sub InputRoomNumber()

Dim i, j, k As Integer
h = 0
n = 10

Do
    n = n + 22
    i = 0

    Do
        i = i + 1
        j = 0

        Do
            j = j + 1
            k = Cells(j + n, i + 1)
            If k = 0 Then
            Else
                h = h + 1
                Cells(j + n, i + 1) = h
            End If
            
        Loop While j < 19
    Loop While i < 19
Loop While n < 209

End Sub
