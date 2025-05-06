Sub reset()

Dim i, j, s, t, n As Integer
n = 10

Do
    n = n + 22
    i = 0

    Do
        j = 0
        i = i + 1

        Do
            j = j + 1
            Cells(j + n, i + 1).Value = "1"
        Loop While j < 19
    Loop While i < 19

    s = 0
    Do
        t = 0
        s = s + 2

        Do
            t = t + 2
            Cells(t + n, s + 1).Value = "0"
        Loop While t < 18
    Loop While s < 18
Loop While n < 209

Range("z6:ae2805").ClearContents
Range("ah6:am2805").ClearContents
Range("ap6:au2805").ClearContents
Range("z2808:ln5807").ClearContents

End Sub
