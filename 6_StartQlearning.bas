Sub StartQlearning()

Dim gamm, alph, epsi, rand As Single
Dim i, j, k, m, episode, step, state, action As Integer
m = Cells(249, 20).Value

Dim trans(2800, 6) As Integer
Dim r(2800), q(2800, 6) As Single

gamm = 0.7  '割引率
alph = 0.5  '学習率

'状態iの即時報酬
For i = 1 To (m - 1)
    r(i) = -1
Next i
r(m) = 100

'状態iでアクションjを採用したときの状態推移先
For i = 1 To m
    For j = 1 To 6
        trans(i, j) = Cells(i + 5, j + 25).Value
    Next j
Next i

'状態iでアクションjを採用したときのQの初期値
For i = 1 To m
    For j = 1 To 6
        q(i, j) = Cells(i + 5, j + 33).Value
    Next j
Next i

For episode = 1 To 3000

    state = 1
    Cells(episode + 2807, 26).Value = state  '初期状態
    step = 0

    Do
        step = step + 1
        epsi = 1 - (episode - 1) / 3000  'Greedy比

        If Rnd < epsi Then  '乱数がGreedy比より小さいときはランダムにアクションを選ぶ
            Do
                action = Fix(Rnd * 6) + 1  '1/6の確率で前後上下左右を振り分け
            Loop While trans(state, action) <= 0  '選んだアクションが実行できないときは選び直す
        Else  '乱数がGreedy比以上のときはQ値が最大のアクションを選ぶ
            action = 1
            For i = 2 To 6
                If q(state, i) > q(state, action) Then  'これまでより大きいQ値が見つかったら
                    action = i  'アクッションを変更
                End If
            Next i
        End If

        i = trans(state, action)  '現在の状態とアクションから次の状態iを求める
        k = 1
        For j = 2 To 6  '次の状態iにおいてQ値が最大となるアクションを求める
            If q(i, j) > q(i, k) Then  'これまでより大きいQ値が見つかったら
                k = j  'アクッションを変更
            End If
        Next j

        q(state, action) = (1 - alph) * q(state, action) + alph * (r(i) + gamm * q(i, k))   'Q値の更新
        state = i
        Cells(episode + 2807, step + 26).Value = state  '次の状態を出力

    Loop While state < m And step < 300
    Cells(episode + 2807, 26).Value = step  '状態8に到達するまでのステップ数を出力

Next episode

'Q値(最終値）
For i = 1 To m
    For j = 1 To 6
        Cells(i + 5, j + 41).Value = q(i, j)
    Next j
Next i

End Sub
