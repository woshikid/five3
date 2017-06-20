Attribute VB_Name = "Module1"
Option Explicit
Public files() As String
Public mlength As Long
Public mheight As Long
Public mcolour As Byte
Public handcut As Boolean
Public mmap() As Byte
Public order() As Long
Public mdeep As Long
Private Const valueMax As Long = 2147483647
Public Const valueMin As Long = -2147483647
Private Const valueWin As Long = 1000000000
Private Const valueStep As Long = 10000
Public M As Long
Public nextPoint As Long
Public winning As Boolean
Private mtxy() As Long
Private mfxy() As Long
Public mstop As Boolean
Public smain As Long
Public sdetail() As Long
Public cache() As Long
Public cachePos As Long
Public cacheMax As Long
Public data(44) As Byte '15*15\5-1
Public records As Long
Public trecords As Long
Public rindex As Long
Public memFull As Boolean
Private Type dbu
    dbi() As String
End Type
Public db() As dbu
Public dbMax As Long
Public dbCount As Long
Public dbLarge As Boolean
Public MD5Engine As MD5
Public dbHit As Long
Public dbUse As Long
Public cacheHit As Long
Public cacheUse As Long

'找到最佳位置的点
Public Function getPoint(ByVal colour As Byte, ByRef omap() As Byte) As Long
    winning = False
    'check opp's win
    If selectDB(omap) Then
        getPoint = -1
        Exit Function
    End If
    smain = 0
    'pre scan
    Dim ij As Long
    ij = findWin(colour, omap)
    If ij >= 0 Then
        winning = True
        getPoint = ij
        Exit Function
    End If
    Dim value As Long
    Dim cmap() As Byte
    ij = findWin(IIf(colour = 1, 2, 1), omap)
    If ij >= 0 Then
        value = scanMap(colour, omap, ij \ mheight, ij Mod mheight, False)
        If value = valueMin Then
            updateDB omap 'write to file
            getPoint = -1
            Exit Function
        ElseIf value >= valueWin Then
            winning = True
            getPoint = ij
            Exit Function
        End If
        value = stopFour
        If scanWin(colour, omap, 0, ij) >= 0 Then
            winning = True
            getPoint = ij
            Exit Function
        End If
        smain = mlength * mheight
        cmap = omap
        cmap(ij \ mheight, ij Mod mheight) = colour
        If scanWin(IIf(colour = 1, 2, 1), cmap, 0, value) < 0 Then
            getPoint = ij
        Else
            updateDB omap 'write to file
            getPoint = -1
        End If
        Exit Function
    End If
    ij = scanWin(colour, omap, 0, -1)
    If ij >= 0 Then
        winning = True
        getPoint = ij
        Exit Function
    End If
    'first scan
    Dim values() As Long
    ReDim values(mlength - 1, mheight - 1)
    Dim ijs() As Long
    ReDim ijs(mlength - 1, mheight - 1)
    Dim o As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If Not checkEmpty(omap, i, j) Then
                value = scanMap(colour, omap, i, j, False)
                If value > valueMin Then
                    ijs(i, j) = stopFour
                    If value < 53 Then
                        values(i, j) = 1
                    ElseIf value > 100 Then
                        values(i, j) = 4
                    ElseIf ijs(i, j) >= 0 Then
                        values(i, j) = 3
                    Else
                        values(i, j) = 2
                    End If
                End If
            End If
        End If
    Next o
    'deep scan
    getPoint = -1
    For k = 4 To 1 Step -1
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If values(i, j) = k Then
            smain = o + 1
            cmap = omap
            cmap(i, j) = colour
            If scanWin(IIf(colour = 1, 2, 1), cmap, 0, ijs(i, j)) < 0 Then
                getPoint = i * mheight + j
                If Rnd < 0.5 Then Exit Function
            End If
        End If
    Next o
    Next k
    If getPoint >= 0 Then Exit Function
    'new game or end game or losing
    Dim newGame As Boolean
    Dim endGame As Boolean
    newGame = True
    endGame = True
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        If mmap(i, j) = 0 Then
            endGame = False
        Else
            newGame = False
        End If
    Next j
    Next i
    If newGame = True Then
        getPoint = order(0, 0) * mheight + order(0, 1)
    ElseIf endGame = True Then
        getPoint = -1
    Else 'losing
        updateDB omap 'write to file
        getPoint = -1
    End If
End Function

'检查是否是空白区域
Private Function checkEmpty(ByRef omap() As Byte, ByVal x As Long, ByVal y As Long) As Boolean
    checkEmpty = False
    Dim i As Long
    Dim j As Long
    For i = x - 2 To x + 2
        If i >= 0 And i < mlength Then
            For j = y - 2 To y + 2
                If j >= 0 And j < mheight Then
                    If omap(i, j) <> 0 Then Exit Function
                End If
            Next j
        End If
    Next i
    checkEmpty = True
End Function

'找到能下棋的位置
Public Function findEmpty(ByVal colour As Byte, ByRef omap() As Byte) As Long
    Dim o As Long
    Dim i As Long
    Dim j As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If scanMap(colour, omap, i, j, True) > valueMin Then
                findEmpty = i * mheight + j
                Exit Function
            End If
        End If
    Next o
    findEmpty = -1
End Function

'找到能赢棋的位置
Public Function findWin(ByVal colour As Byte, ByRef omap() As Byte) As Long
    Dim o As Long
    Dim i As Long
    Dim j As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If checkWin(colour, omap, i, j) Then
                findWin = i * mheight + j
                Exit Function
            End If
        End If
    Next o
    findWin = -1
End Function

'检查下子之后是否赢棋
Public Function checkWin(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long) As Boolean
    checkWin = True
    'line |
    Dim middle As Long
    Dim i As Long
    Dim j As Long
    middle = 1
    j = y - 1
    Do While j >= 0
        If omap(x, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        j = j - 1
    Loop
    
    j = y + 1
    Do While j < mheight
        If omap(x, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        j = j + 1
    Loop

    If middle = 5 Then Exit Function
    If middle > 5 Then
        If colour = 2 Then Exit Function
        If handcut = False Then Exit Function
    End If

    'line -
    middle = 1
    i = x - 1
    Do While i >= 0
        If omap(i, y) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i - 1
    Loop

    i = x + 1
    Do While i < mlength
        If omap(i, y) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i + 1
    Loop

    If middle = 5 Then Exit Function
    If middle > 5 Then
        If colour = 2 Then Exit Function
        If handcut = False Then Exit Function
    End If

    'line /
    middle = 1
    i = x - 1
    j = y + 1
    Do While i >= 0 And j < mheight
        If omap(i, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i - 1
        j = j + 1
    Loop

    i = x + 1
    j = y - 1
    Do While i < mlength And j >= 0
        If omap(i, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i + 1
        j = j - 1
    Loop

    If middle = 5 Then Exit Function
    If middle > 5 Then
        If colour = 2 Then Exit Function
        If handcut = False Then Exit Function
    End If

    'line \
    middle = 1
    i = x - 1
    j = y - 1
    Do While i >= 0 And j >= 0
        If omap(i, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i - 1
        j = j - 1
    Loop

    i = x + 1
    j = y + 1
    Do While i < mlength And j < mheight
        If omap(i, j) = colour Then
            middle = middle + 1
        Else
            Exit Do
        End If
        i = i + 1
        j = j + 1
    Loop

    If middle = 5 Then Exit Function
    If middle > 5 Then
        If colour = 2 Then Exit Function
        If handcut = False Then Exit Function
    End If

    checkWin = False
End Function

'该对方下，但是我有活三的情况
Private Function scanWinThree(ByVal colour As Byte, ByRef omap() As Byte, ByVal index As Long, ByVal ij As Long) As Boolean
    'scan all
    scanWinThree = False
    Dim value As Long
    Dim cmap() As Byte
    Dim o As Long
    Dim i As Long
    Dim j As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If Not checkEmpty(omap, i, j) Then
                If scanMap(IIf(colour = 1, 2, 1), omap, i, j, False) >= valueWin Then Exit Function
                cmap = omap
                cmap(i, j) = IIf(colour = 1, 2, 1)
                value = stopFour
                If ij < 0 Or (i * mheight + j) = ij Or value >= 0 Then
                    If scanWin(colour, cmap, index + 1, value) < 0 Then Exit Function
                ElseIf scanMap(colour, cmap, ij \ mheight, ij Mod mheight, False) < valueWin Then
                    If scanWin(colour, cmap, index + 1, value) < 0 Then Exit Function
                End If
            End If
        End If
    Next o
    scanWinThree = True
End Function

'检查下此处是否能赢棋
Private Function scanWinS(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal index As Long) As Boolean
    scanWinS = False
    'pre scan
    Dim value As Long
    value = scanMap(colour, omap, x, y, False)
    If value = valueMin Then Exit Function
    If value >= valueWin Then
        scanWinS = True
        Exit Function
    End If
    If value < 53 Then Exit Function
    'deep scan
    Dim cmap() As Byte
    cmap = omap
    cmap(x, y) = colour
    value = stopFour
    'scan four
    If value >= 0 Then
        If scanMap(IIf(colour = 1, 2, 1), cmap, value \ mheight, value Mod mheight, False) >= valueWin Then Exit Function
        cmap(value \ mheight, value Mod mheight) = IIf(colour = 1, 2, 1)
        If scanWin(colour, cmap, index + 1, stopFour) >= 0 Then scanWinS = True
        Exit Function
    End If
    'scan three
    value = findThreeFast
    If value >= 0 Then scanWinS = scanWinThree(colour, cmap, index, value)
End Function

'找到盘面活三的位置
Public Function findThree(ByVal colour As Byte, ByRef omap() As Byte) As Long
    Dim o As Long
    Dim i As Long
    Dim j As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If Not checkEmpty(omap, i, j) Then
                If scanMap(colour, omap, i, j, False) >= valueWin Then
                    findThree = i * mheight + j
                    Exit Function
                End If
            End If
        End If
    Next o
    findThree = -1
End Function

'根据scanMap的结果快速找到活三的位置
Private Function findThreeFast() As Long
    Dim i As Long
    For i = 0 To 3
        If mtxy(i, 0) = 1 Or mtxy(i, 0) = 3 Then
            findThreeFast = mtxy(i, 1) * mheight + mtxy(i, 2)
            Exit Function
        ElseIf mtxy(i, 0) = 2 Then
            findThreeFast = mtxy(i, 3) * mheight + mtxy(i, 4)
            Exit Function
        End If
    Next i
    findThreeFast = -1
End Function

'得到阻止冲四的点的位置
Private Function stopFour() As Long
    Dim i As Long
    For i = 0 To 3
        If mfxy(i, 0) = 1 Then
            stopFour = mfxy(i, 1) * mheight + mfxy(i, 2)
            Exit Function
        End If
    Next i
    stopFour = -1
End Function

'找到必胜的点
Private Function scanWin(ByVal colour As Byte, ByRef omap() As Byte, ByVal index As Long, ByVal ij As Long) As Long
    DoEvents 'important
    If mstop = True Or index > mdeep Then
        scanWin = -1
        Exit Function
    End If
    'scan cache
    Dim value As Long
    If index = 0 Then
        Erase cache
        memFull = False
        ReDim cache(M)
        cacheMax = 3
    Else
        value = selectCache(omap)
        If value > -2 Then
            scanWin = value
            Exit Function
        End If
    End If
    'check opp's win
    If selectDB(omap) Then
        scanWin = -1
        updateCache omap, scanWin
        Exit Function
    End If
    'clean the progress data
    Dim o As Long
    For o = index To 99
        sdetail(o) = 0
    Next o
    'pre scan
    Dim cmap() As Byte
    If ij >= 0 Then
        sdetail(index) = 9
        cmap = omap
        cmap(ij \ mheight, ij Mod mheight) = colour
        'scan db
        If selectDB(cmap) Then
            scanWin = ij
            updateCache omap, scanWin
            Exit Function
        End If
        'normal scan
        If scanWinS(colour, omap, ij \ mheight, ij Mod mheight, index) Then
            updateDB cmap 'write to file
            scanWin = ij
            updateCache omap, scanWin
            Exit Function
        End If
        'scan if I have three
        value = findThree(colour, omap)
        If value >= 0 Then
            If scanWinThree(colour, cmap, index, value) Then
                updateDB cmap 'write to file
                scanWin = ij
                updateCache omap, scanWin
                Exit Function
            End If
        End If
        scanWin = -1
        updateCache omap, scanWin
        Exit Function
    End If
    'scan db
    Dim i As Long
    Dim j As Long
    If index = 0 Then
        For o = 0 To mlength * mheight - 1
            i = order(o, 0)
            j = order(o, 1)
            If omap(i, j) = 0 Then
                If Not checkEmpty(omap, i, j) Then
                    cmap = omap
                    cmap(i, j) = colour
                    If selectDB(cmap) Then
                        scanWin = i * mheight + j
                        updateCache omap, scanWin
                        Exit Function
                    End If
                End If
            End If
        Next o
    End If
    'first scan
    Dim total As Long
    Dim scaned As Long
    Dim values() As Long
    ReDim values(mlength - 1, mheight - 1)
    Dim k As Long
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If omap(i, j) = 0 Then
            If Not checkEmpty(omap, i, j) Then
                value = scanMap(colour, omap, i, j, False)
                If value >= 53 Then
                    If value >= valueWin Then
                        cmap = omap
                        cmap(i, j) = colour
                        updateDB cmap 'write to file
                        scanWin = i * mheight + j
                        updateCache omap, scanWin
                        Exit Function
                    End If
                    If value > 100 Then
                        values(i, j) = 3
                    ElseIf stopFour >= 0 Then
                        values(i, j) = 2
                    Else
                        values(i, j) = 1
                    End If
                    total = total + 1
                End If
            End If
        End If
    Next o
    'deep scan
    For k = 3 To 1 Step -1
    For o = 0 To mlength * mheight - 1
        i = order(o, 0)
        j = order(o, 1)
        If values(i, j) = k Then
            scaned = scaned + 1
            sdetail(index) = (scaned * 10 - 1) \ total
            If scanWinS(colour, omap, i, j, index) Then
                cmap = omap
                cmap(i, j) = colour
                updateDB cmap 'write to file
                scanWin = i * mheight + j
                updateCache omap, scanWin
                Exit Function
            End If
        End If
    Next o
    Next k
    scanWin = -1
    updateCache omap, scanWin
End Function

'计算下子之后的棋力
Public Function scanMap(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal checkonly As Boolean) As Long
    If checkonly = True And (colour = 2 Or handcut = False) Then
        scanMap = 1
        Exit Function
    End If
    Dim cthree As Long
    Dim three As Long
    Dim four As Long
    Dim lfour As Boolean
    Dim cuted As Boolean
    Dim tvalue As Long
    
    Dim i As Long
    Dim value As Long
    ReDim mtxy(3, 4)
    ReDim mfxy(3, 2)
    For i = 0 To 3
        If i = 0 Then value = scan0(colour, omap, x, y, True)
        If i = 1 Then value = scan1(colour, omap, x, y, True)
        If i = 2 Then value = scan2(colour, omap, x, y, True)
        If i = 3 Then value = scan3(colour, omap, x, y, True)
        If value = valueMax Then
            scanMap = valueMax
            Exit Function
        ElseIf value = valueMin Then
            cuted = True
        Else
            tvalue = tvalue + value
            If value = 300 Then lfour = True
            If value = 300 Or value = 54 Then four = four + 1
            If value = 53 Then three = three + 1
        End If
    Next i
    Dim txy() As Long
    Dim fxy() As Long
    txy = mtxy
    fxy = mfxy

    If colour = 1 And handcut = True Then
        If cuted = True Or four > 1 Then
            scanMap = valueMin
            Exit Function
        End If
        If checkonly = True And three < 2 Then
            scanMap = 1
            Exit Function
        End If
        If three > 0 Then
            Dim cmap() As Byte
            cmap = omap
            cmap(x, y) = 1
            For i = 0 To 3
                If checkonly = True And (cthree + three) < 2 Then
                    scanMap = 1
                    Exit Function
                End If
                If txy(i, 0) <> 0 Then
                    cuted = False
                    three = three - 1
                    tvalue = tvalue - 53
                    If txy(i, 0) = 1 Then
                        If scanMap(1, cmap, txy(i, 1), txy(i, 2), True) = valueMin Then cuted = True
                    ElseIf txy(i, 0) = 2 Then
                        If scanMap(1, cmap, txy(i, 3), txy(i, 4), True) = valueMin Then cuted = True
                    Else
                        If scanMap(1, cmap, txy(i, 1), txy(i, 2), True) = valueMin Then
                            If scanMap(1, cmap, txy(i, 3), txy(i, 4), True) = valueMin Then cuted = True
                        End If
                    End If
                    If cuted = False Then
                        cthree = cthree + 1
                        If cthree > 1 Then
                            scanMap = valueMin
                            Exit Function
                        End If
                        tvalue = tvalue + 53
                    End If
                End If
            Next i
            If checkonly = True Then
                scanMap = 1
                Exit Function
            End If
            three = cthree
        End If
    End If

    mtxy = txy
    mfxy = fxy
    If lfour = True Or four > 1 Then tvalue = tvalue + valueWin
    scanMap = tvalue
End Function

'line |
Private Function scan0(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal record As Boolean) As Long
    Dim middle As Long
    Dim x1 As Long
    Dim s1 As Long
    Dim e1 As Long
    Dim x2 As Long
    Dim s2 As Long
    Dim e2 As Long
    Dim j As Long
    middle = 1
    j = y - 1
    Do While j >= 0
        If omap(x, j) = colour Then
            If s1 = 0 Then
                x1 = x1 + 1
            Else
                e1 = e1 + 1
            End If
        ElseIf omap(x, j) = 0 Then
            s1 = s1 + 1
            If s1 = 2 Then Exit Do
        Else
            Exit Do
        End If
        j = j - 1
    Loop

    j = y + 1
    Do While j < mheight
        If omap(x, j) = colour Then
            If s2 = 0 Then
                x2 = x2 + 1
            Else
                e2 = e2 + 1
            End If
        ElseIf omap(x, j) = 0 Then
            s2 = s2 + 1
            If s2 = 2 Then Exit Do
        Else
            Exit Do
        End If
        j = j + 1
    Loop

    middle = middle + x1 + x2
    scan0 = calValue(colour, omap, x, y, 0, middle, s1, e1, x, y - 1 - x1, s2, e2, x, y + 1 + x2, record)
End Function

'line -
Private Function scan1(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal record As Boolean) As Long
    Dim middle As Long
    Dim x1 As Long
    Dim s1 As Long
    Dim e1 As Long
    Dim x2 As Long
    Dim s2 As Long
    Dim e2 As Long
    Dim i As Long
    middle = 1
    i = x - 1
    Do While i >= 0
        If omap(i, y) = colour Then
            If s1 = 0 Then
                x1 = x1 + 1
            Else
                e1 = e1 + 1
            End If
        ElseIf omap(i, y) = 0 Then
            s1 = s1 + 1
            If s1 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i - 1
    Loop

    i = x + 1
    Do While i < mlength
        If omap(i, y) = colour Then
            If s2 = 0 Then
                x2 = x2 + 1
            Else
                e2 = e2 + 1
            End If
        ElseIf omap(i, y) = 0 Then
            s2 = s2 + 1
            If s2 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i + 1
    Loop

    middle = middle + x1 + x2
    scan1 = calValue(colour, omap, x, y, 1, middle, s1, e1, x - 1 - x1, y, s2, e2, x + 1 + x2, y, record)
End Function

'line /
Private Function scan2(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal record As Boolean) As Long
    Dim middle As Long
    Dim x1 As Long
    Dim s1 As Long
    Dim e1 As Long
    Dim x2 As Long
    Dim s2 As Long
    Dim e2 As Long
    Dim i As Long
    Dim j As Long
    middle = 1
    i = x - 1
    j = y + 1
    Do While i >= 0 And j < mheight
        If omap(i, j) = colour Then
            If s1 = 0 Then
                x1 = x1 + 1
            Else
                e1 = e1 + 1
            End If
        ElseIf omap(i, j) = 0 Then
            s1 = s1 + 1
            If s1 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i - 1
        j = j + 1
    Loop

    i = x + 1
    j = y - 1
    Do While i < mlength And j >= 0
         If omap(i, j) = colour Then
            If s2 = 0 Then
                x2 = x2 + 1
            Else
                e2 = e2 + 1
            End If
        ElseIf omap(i, j) = 0 Then
            s2 = s2 + 1
            If s2 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i + 1
        j = j - 1
    Loop

    middle = middle + x1 + x2
    scan2 = calValue(colour, omap, x, y, 2, middle, s1, e1, x - 1 - x1, y + 1 + x1, s2, e2, x + 1 + x2, y - 1 - x2, record)
End Function

'line \
Private Function scan3(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal record As Boolean) As Long
    Dim middle As Long
    Dim x1 As Long
    Dim s1 As Long
    Dim e1 As Long
    Dim x2 As Long
    Dim s2 As Long
    Dim e2 As Long
    Dim i As Long
    Dim j As Long
    middle = 1
    i = x - 1
    j = y - 1
    Do While i >= 0 And j >= 0
        If omap(i, j) = colour Then
            If s1 = 0 Then
                x1 = x1 + 1
            Else
                e1 = e1 + 1
            End If
        ElseIf omap(i, j) = 0 Then
            s1 = s1 + 1
            If s1 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i - 1
        j = j - 1
    Loop

    i = x + 1
    j = y + 1
    Do While i < mlength And j < mheight
        If omap(i, j) = colour Then
            If s2 = 0 Then
                x2 = x2 + 1
            Else
                e2 = e2 + 1
            End If
        ElseIf omap(i, j) = 0 Then
            s2 = s2 + 1
            If s2 = 2 Then Exit Do
        Else
            Exit Do
        End If
        i = i + 1
        j = j + 1
    Loop

    middle = middle + x1 + x2
    scan3 = calValue(colour, omap, x, y, 3, middle, s1, e1, x - 1 - x1, y - 1 - x1, s2, e2, x + 1 + x2, y + 1 + x2, record)
End Function

'计算下子之后的棋力数值
Private Function calValue(ByVal colour As Byte, ByRef omap() As Byte, ByVal x As Long, ByVal y As Long, ByVal ltype As Long, ByVal middle As Long, ByVal s1 As Long, ByVal e1 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal s2 As Long, ByVal e2 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal record As Boolean) As Long
    Dim four1 As Boolean
    Dim four2 As Boolean
    Dim three As Boolean
    Dim hthree As Boolean

    If middle > 5 Then '六连珠等情况
        If colour = 2 Or handcut = False Then
            calValue = valueMax
        Else '禁手了
            calValue = valueMin
        End If
        Exit Function
    End If
    If middle = 5 Then '赢了
        calValue = valueMax
        Exit Function
    End If
    '嵌六的情况
    If (middle + e1) > 4 And (colour = 2 Or handcut = False) Then four1 = True
    If (middle + e2) > 4 And (colour = 2 Or handcut = False) Then four2 = True
    '嵌五或连四
    If (middle + e1) = 4 Then
        If middle = 4 Then
            If s1 > 0 Then four1 = True
        Else
            four1 = True
        End If
    End If
    If (middle + e2) = 4 Then
        If middle = 4 Then
            If s2 > 0 Then four2 = True
        Else
            four2 = True
        End If
    End If
    If four1 = True And four2 = True Then '活四
        If colour = 2 Or handcut = False Or middle = 4 Then
            calValue = 300
        Else '成一条直线的四四禁手
            calValue = valueMin
        End If
        Exit Function
    End If
    Dim value As Long
    Dim cmap() As Byte
    cmap = omap
    cmap(x, y) = colour
    If four1 = True Or four2 = True Then '冲四
        Dim txy() As Long
        Dim fxy() As Long
        If four1 = True Then
            If colour = 2 And handcut = True Then '对方禁手则升级为活四
                txy = mtxy
                fxy = mfxy
                value = scanMap(1, cmap, x1, y1, True)
                mtxy = txy
                mfxy = fxy
                If value = valueMin Then
                    calValue = 300
                    Exit Function
                End If
            End If
            If record = True Then
                mfxy(ltype, 0) = 1
                mfxy(ltype, 1) = x1
                mfxy(ltype, 2) = y1
            End If
        Else
            If colour = 2 And handcut = True Then
                txy = mtxy
                fxy = mfxy
                value = scanMap(1, cmap, x2, y2, True)
                mtxy = txy
                mfxy = fxy
                If value = valueMin Then
                    calValue = 300
                    Exit Function
                End If
            End If
            If record = True Then
                mfxy(ltype, 0) = 1
                mfxy(ltype, 1) = x2
                mfxy(ltype, 2) = y2
            End If
        End If
        calValue = 54
        Exit Function
    End If
    '活三或嵌四
    If (middle + e1) = 3 And s1 > 0 Then
        If ltype = 0 Then value = scan0(colour, cmap, x1, y1, False)
        If ltype = 1 Then value = scan1(colour, cmap, x1, y1, False)
        If ltype = 2 Then value = scan2(colour, cmap, x1, y1, False)
        If ltype = 3 Then value = scan3(colour, cmap, x1, y1, False)
        If value = 54 Then '再下一子是冲四，则目前眠三
            hthree = True
        ElseIf value = 300 Then '再下一子是活四，则目前活三
            three = True
            If record = True Then
                mtxy(ltype, 0) = mtxy(ltype, 0) + 1
                mtxy(ltype, 1) = x1
                mtxy(ltype, 2) = y1
            End If
        End If
    End If
    If (middle + e2) = 3 And s2 > 0 Then
        If ltype = 0 Then value = scan0(colour, cmap, x2, y2, False)
        If ltype = 1 Then value = scan1(colour, cmap, x2, y2, False)
        If ltype = 2 Then value = scan2(colour, cmap, x2, y2, False)
        If ltype = 3 Then value = scan3(colour, cmap, x2, y2, False)
        If value = 54 Then
            hthree = True
        ElseIf value = 300 Then
            three = True
            If record = True Then
                mtxy(ltype, 0) = mtxy(ltype, 0) + 2
                mtxy(ltype, 3) = x2
                mtxy(ltype, 4) = y2
            End If
        End If
    End If
    If three = True Then
        calValue = 53 '活三
    ElseIf hthree = True Then
        calValue = 3 '眠三
    ElseIf middle = 1 And e1 = 1 And e2 = 1 Then
        calValue = 3 '眠三
    ElseIf (middle + e1) = 2 And s1 = 2 Then
        If e2 > 0 Then
            calValue = 2 '眠二
        ElseIf s2 = 2 Then
            calValue = 5 '活二
        ElseIf s2 > 0 Then
            calValue = 4 '半活二
        Else
            calValue = 2 '眠二
        End If
    ElseIf (middle + e2) = 2 And s2 = 2 Then
        If e1 > 0 Then
            calValue = 2 '眠二
        ElseIf s1 = 2 Then
            calValue = 5 '活二
        ElseIf s1 > 0 Then
            calValue = 4 '半活二
        Else
            calValue = 2 '眠二
        End If
    ElseIf (middle + e1 + e2) = 1 And (s1 + s2) = 4 Then
        calValue = 1 '活一
    Else
        calValue = 0
    End If
End Function

Public Function inDB(ByVal value As String, ByVal insert As Boolean) As Boolean
    On Error Resume Next
    inDB = False
    If LenB(value) <> 45 Then Exit Function
    Dim h As Long
    h = hash(value)
    If isEmpty(db(h).dbi) Then
        If insert = False Or memFull = True Then Exit Function
        ReDim db(h).dbi(0)
        If Err.Number <> 0 Then
            memFull = True
            Exit Function
        End If
        db(h).dbi(0) = value
        If Err.Number <> 0 Then
            memFull = True
            Exit Function
        End If
        dbCount = dbCount + 1
        If 1 > dbMax Then dbMax = 1
        Exit Function
    End If
    '二分搜索
    Dim low, middle, high As Long
    low = 0
    high = UBound(db(h).dbi)
    Do While low <= high
        middle = (low + high) \ 2
        If db(h).dbi(middle) = value Then
            inDB = True
            Exit Function
        End If
        If db(h).dbi(middle) > value Then
            high = middle - 1
        Else
            low = middle + 1
        End If
    Loop
    If insert = False Or memFull = True Then Exit Function
    ReDim Preserve db(h).dbi(UBound(db(h).dbi) + 1)
    If Err.Number <> 0 Then
        memFull = True
        Exit Function
    End If
    Dim i As Long
    For i = UBound(db(h).dbi) - 1 To low Step -1
        db(h).dbi(i + 1) = db(h).dbi(i)
    Next i
    db(h).dbi(low) = value
    If Err.Number <> 0 Then
        memFull = True
        Exit Function
    End If
    dbCount = dbCount + 1
    If UBound(db(h).dbi) + 1 > dbMax Then dbMax = UBound(db(h).dbi) + 1
End Function

Private Sub insertCache(ByVal value As Byte)
    On Error Resume Next
    If memFull = True Then Exit Sub
    If cache(cachePos + value) = 0 Then '以前没出现过
        If cacheMax + 2 > UBound(cache) Then '内存不够要重新申请内存
            ReDim Preserve cache(UBound(cache) * 1.1)
            If Err.Number <> 0 Then
                memFull = True
                Exit Sub
            End If
        End If
        cache(cachePos + value) = cacheMax '新的棋谱数据指向最近的空白内存
        cacheMax = cacheMax + 3
    End If
    cachePos = cache(cachePos + value) '指针移向下一个棋位的内存空间
End Sub

'查询内存数据库中是否有待查询的棋局
Private Function selectDB(ByRef omap() As Byte) As Boolean
    dbUse = dbUse + 1
    dbHit = dbHit + 1
    selectDB = False
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim c As Long
    Dim b As Byte
    '正常模式
    k = 0
    c = 0
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next j
    Next i
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '上下镜像
    k = 0
    c = 0
    For i = 0 To mlength - 1
    For j = mheight - 1 To 0 Step -1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next j
    Next i
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '左右镜像
    k = 0
    c = 0
    For i = mlength - 1 To 0 Step -1
    For j = 0 To mheight - 1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next j
    Next i
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '对角镜像
    k = 0
    c = 0
    For i = mlength - 1 To 0 Step -1
    For j = mheight - 1 To 0 Step -1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next j
    Next i
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '旋转镜像
    k = 0
    c = 0
    For j = 0 To mheight - 1
    For i = 0 To mlength - 1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next i
    Next j
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '旋转上下镜像
    k = 0
    c = 0
    For j = 0 To mheight - 1
    For i = mlength - 1 To 0 Step -1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next i
    Next j
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '旋转左右镜像
    k = 0
    c = 0
    For j = mheight - 1 To 0 Step -1
    For i = 0 To mlength - 1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next i
    Next j
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    '旋转对角镜像
    k = 0
    c = 0
    For j = mheight - 1 To 0 Step -1
    For i = mlength - 1 To 0 Step -1
        If omap(i, j) > 2 Then Exit Function
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next i
    Next j
    selectDB = inDB(data, False)
    If selectDB = True Then Exit Function
    dbHit = dbHit - 1
End Function

Private Function selectCache(ByRef omap() As Byte) As Long
    cacheUse = cacheUse + 1
    Dim i As Long
    Dim j As Long
    selectCache = -2
    cachePos = 0
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        If omap(i, j) > 2 Then Exit Function
        cachePos = cache(cachePos + omap(i, j))
        If cachePos = 0 Then Exit Function
    Next j
    Next i
    selectCache = cache(cachePos)
    cacheHit = cacheHit + 1
End Function

'将棋谱插入内存与文件中
Private Sub updateDB(ByRef omap() As Byte)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim c As Long
    Dim b As Byte
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        If omap(i, j) > 2 Then Exit Sub
        If c = 0 Then
            b = 81 * omap(i, j)
        ElseIf c = 1 Then
            b = b + 27 * omap(i, j)
        ElseIf c = 2 Then
            b = b + 9 * omap(i, j)
        ElseIf c = 3 Then
            b = b + 3 * omap(i, j)
        ElseIf c = 4 Then
            b = b + omap(i, j)
        End If
        c = c + 1
        If c = 5 Then
            c = 0
            data(k) = b
            k = k + 1
        End If
    Next j
    Next i
    If inDB(data, True) = True Then Exit Sub
    'last check before write to file
    If k = 45 Then Put #1, , data
End Sub

Private Sub updateCache(ByRef omap() As Byte, ByVal value As Long)
    cachePos = 0
    Dim i As Long
    Dim j As Long
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        If omap(i, j) > 2 Then Exit Sub
        insertCache omap(i, j)
    Next j
    Next i
    If memFull = False Then cache(cachePos) = value
End Sub

'将文件1中的数据读取到内存中
Public Sub loadDB()
    records = LOF(1) \ 45 '225个棋谱数据除以5，占用45个字节
    trecords = trecords + records
    For rindex = 0 To records - 1
        DoEvents
        Get #1, , data
        inDB data, True
    Next rindex
End Sub

Private Function hash(ByRef value As String) As Long
    Dim md5str As String
    md5str = MD5Engine.DigestStrToHexStr(value)
    Dim i As Byte
    For i = 1 To 5
        hash = hash * 16 + Val("&H" & Mid(md5str, i, 1))
    Next i
    If dbLarge = True Then hash = hash * 16 + Val("&H" & Mid(md5str, i, 1))
End Function

Private Function isEmpty(ByRef list() As String) As Boolean
    On Error Resume Next
    isEmpty = True
    Dim i As Long
    i = UBound(list)
    If Err.Number <> 0 Then
        Err.Clear
    Else
        isEmpty = False
    End If
End Function

Public Sub Main()
    If App.PrevInstance Then End
    Set MD5Engine = New MD5
    M = 1024
    M = M * 1024
    files = Split(Trim(Command))
    If UBound(files) < 0 Then
        Form1.Show
    ElseIf UBound(files) = 0 Then
        MsgBox "参数不正确", vbCritical, "五子棋棋谱整理器"
    Else
        Form3.Show
    End If
End Sub
