VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000040C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "五子棋决策系统"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7380
   ForeColor       =   &H00000000&
   Icon            =   "Five.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7380
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer autoClicker 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   1440
   End
   Begin VB.Timer clicker 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   1440
   End
   Begin VB.Timer unloader 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1680
      Top             =   1440
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "Five.frx":5F32
      Left            =   1200
      List            =   "Five.frx":5F54
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   615
   End
   Begin VB.Timer showCount 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1440
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "Five.frx":5F76
      Left            =   1800
      List            =   "Five.frx":5F80
      MousePointer    =   1  'Arrow
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Default         =   -1  'True
      Height          =   330
      Left            =   3375
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   645
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "Five.frx":5F94
      Left            =   2760
      List            =   "Five.frx":5F9E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   90
      Width           =   615
   End
   Begin VB.Shape Dot 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Z 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   315
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "计算深度："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00000000&
      Index           =   0
      Visible         =   0   'False
      X1              =   360
      X2              =   5610
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Menu Mgame 
      Caption         =   "游戏(&G)"
      Begin VB.Menu Mnew 
         Caption         =   "重新开局(&N)"
      End
   End
   Begin VB.Menu Mhelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu Mstatistics 
         Caption         =   "统计(&S)"
      End
      Begin VB.Menu Mabout 
         Caption         =   "关于(&A)"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "托盘"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayHide 
         Caption         =   "显示/隐藏(&H)"
      End
      Begin VB.Menu mnuTrayRun 
         Caption         =   "计算/中止(&S)"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lastButton As Integer
Private lastX As Single
Private lastY As Single
Private click As Boolean
Private loading As Boolean
Private mainIcon As Boolean
Private auto As Boolean
Private automessage As String

Private Sub autoClicker_Timer()
    autoClicker.Enabled = False
    If auto = False Then Exit Sub
    If Command1.Enabled = True Then Command1_Click
End Sub

Private Sub clicker_Timer()
    clicker.Enabled = False
    If Command1.Enabled = True Then Command1_Click
End Sub

Private Sub Combo1_Click()
    mdeep = Val(Combo1.Text)
    If Form1.Visible = True Then Command1.SetFocus
End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "白" Then
        Combo3.BackColor = vbWhite
        Combo3.ForeColor = vbBlack
        mcolour = 2
    Else
        Combo3.BackColor = vbBlack
        Combo3.ForeColor = vbWhite
        mcolour = 1
    End If
    If Form1.Visible = True Then Command1.SetFocus
End Sub

Private Sub Combo4_Click()
    If Combo4.Text = "无禁手" Then
        If UBound(db) = 0 Or handcut = True Then
            handcut = False
            loadFile "handfree"
        End If
    Else
        If UBound(db) = 0 Or handcut = False Then
            handcut = True
            loadFile "handcut"
        End If
    End If
    If Form1.Visible = True Then Command1.SetFocus
End Sub

Private Sub loadFile(ByVal FileName As String)
    If FileName = vbNullString Then Exit Sub
    Mnew.Enabled = False
    Combo1.Enabled = False
    Combo3.Enabled = False
    Combo4.Enabled = False
    Command1.Enabled = False
    mnuTrayRun.Enabled = False
    setCaption "数据加载中"
    Form1.MousePointer = 12
    click = False
    loading = True
    showCount.Enabled = True
    'load db
    Close #1
    Erase cache
    Erase db
    memFull = False
    dbLarge = False
    ReDim db(M)
    dbMax = 0
    dbCount = 0
    trecords = 0
    dbHit = 0
    dbUse = 0
    cacheHit = 0
    cacheUse = 0
    If Dir(FileName & ".db", vbSystem + vbHidden) <> vbNullString Then
        Open FileName & ".db" For Binary As #1
            If LOF(1) \ 45 > M Then
                dbLarge = True
                ReDim db(16 * M)
            End If
            loadDB
        Close #1
    End If
    If Dir(FileName & ".dat", vbSystem + vbHidden) <> vbNullString Then SetAttr FileName & ".dat", vbNormal
    Open FileName & ".dat" For Binary As #1
    loadDB
    'keep the file open for write
    If memFull = True Then MsgBox "内存不足，不能加载所有数据", vbInformation, "数据加载错误"
    'load finished
    showCount.Enabled = False
    loading = False
    click = True
    Form1.MousePointer = 0
    setCaption "五子棋决策系统"
    mnuTrayRun.Enabled = True
    Command1.Enabled = True
    Combo1.Enabled = True
    Combo3.Enabled = True
    Combo4.Enabled = True
    Mnew.Enabled = True
    MsgBox "棋谱数量：" & trecords & vbNewLine & "有效载入：" & dbCount & vbNewLine & "最大节点数：" & dbMax & vbNewLine & "消耗内存：" & (dbCount * 45 \ M) & "M"
End Sub

Private Sub Command1_Click()
    If Command1.BackColor = &HFF00& Then
        setCaption "五子棋决策系统  计算中"
        Dot(0).Visible = False
        Mnew.Enabled = False
        Combo1.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Command1.BackColor = &HFF&
        Form1.MousePointer = 12
        click = False
        mstop = False
        showCount.Enabled = True
        nextPoint = getPoint(mcolour, mmap)
        showCount.Enabled = False
        click = True
        Form1.MousePointer = 0
        Command1.BackColor = &HFF00&
        Combo1.Enabled = True
        Combo3.Enabled = True
        Combo4.Enabled = True
        Mnew.Enabled = True
        If mstop = False And nextPoint >= 0 Then
            setCaption "五子棋决策系统  计算完成" & IIf(winning, "*", "")
            pointDown
            If winning And auto = True Then Mnew_Click
        Else
            If mstop = True Then
                setCaption "五子棋决策系统  计算中止"
                auto = False
            Else
                setCaption "五子棋决策系统  无最佳结果"
                nextPoint = findWin(IIf(mcolour = 1, 2, 1), mmap)
                If nextPoint >= 0 Then nextPoint = IIf(scanMap(mcolour, mmap, nextPoint \ mheight, nextPoint Mod mheight, True) = valueMin, -1, nextPoint)
                If nextPoint < 0 Then nextPoint = findThree(IIf(mcolour = 1, 2, 1), mmap)
                If nextPoint >= 0 Then nextPoint = IIf(scanMap(mcolour, mmap, nextPoint \ mheight, nextPoint Mod mheight, True) = valueMin, -1, nextPoint)
                If nextPoint < 0 Then nextPoint = findEmpty(mcolour, mmap)
                If nextPoint >= 0 Then pointDown
                If auto = True Then Mnew_Click
            End If
        End If
        If auto = True Then autoClicker.Enabled = True
    Else
        mstop = True
        auto = False
    End If
End Sub

Private Sub pointDown()
    mmap(nextPoint \ mheight, nextPoint Mod mheight) = mcolour
    refreshMap
    Dot(0).Left = 270 + (nextPoint \ mheight) * 25 * 15 - 30
    Dot(0).Top = 630 + (nextPoint Mod mheight) * 25 * 15 - 30
    Dot(0).Visible = True
    '以mmap(nextPoint \ mheight, nextPoint Mod mheight)为准，mcolour在refreshMap之后会变
    If mmap(nextPoint \ mheight, nextPoint Mod mheight) = 1 Then
        If checkWin(1, mmap, nextPoint \ mheight, nextPoint Mod mheight) = True And auto = False Then MsgBox "黑棋胜", vbInformation, "五子棋"
    ElseIf mmap(nextPoint \ mheight, nextPoint Mod mheight) = 2 Then
        If checkWin(2, mmap, nextPoint \ mheight, nextPoint Mod mheight) = True And auto = False Then MsgBox "白棋胜", vbInformation, "五子棋"
    End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If auto = True Then Exit Sub
    If Chr(KeyAscii) = "a" Then
        automessage = "a"
    ElseIf Chr(KeyAscii) = "u" And automessage = "a" Then
        automessage = "au"
    ElseIf Chr(KeyAscii) = "t" And automessage = "au" Then
        automessage = "aut"
    ElseIf Chr(KeyAscii) = "o" And automessage = "aut" Then
        automessage = ""
        auto = True
    Else
        automessage = ""
    End If
End Sub

Private Sub Form_DblClick()
    Form_MouseDown lastButton, 0, lastX, lastY
End Sub

Private Sub Form_Load()
    Randomize
    ReDim db(0) 'Combo4_Click需要
    mlength = 15
    mheight = 15
    ReDim mmap(mlength - 1, mheight - 1)
    ReDim order(mlength * mheight - 1, 1)
    order(0, 0) = 7
    order(0, 1) = 7
    Dim i As Long
    Dim j As Long
    Dim k As Long
    i = 1
    For j = 1 To 7
        For k = 7 - j To 7 + j
            order(i, 0) = k
            order(i, 1) = 7 - j
            i = i + 1
            order(i, 0) = k
            order(i, 1) = 7 + j
            i = i + 1
        Next k
        For k = 7 - j + 1 To 7 + j - 1
            order(i, 0) = 7 - j
            order(i, 1) = k
            i = i + 1
            order(i, 0) = 7 + j
            order(i, 1) = k
            i = i + 1
        Next k
    Next j
    ReDim sdetail(99)
    Combo1.Clear
    For i = 0 To 99
        Combo1.AddItem i
    Next i
    Combo1.Text = 0
    Combo3.Text = "黑"
    '270=120+10*15
    '630=480+10*15
    Form1.Width = Form1.Width - Form1.ScaleWidth + 270 * 2 + (mlength - 1) * 25 * 15
    Form1.Height = Form1.Height - Form1.ScaleHeight + 630 + 270 + (mheight - 1) * 25 * 15
    For i = 1 To mheight
        Load Lines(i)
        Lines(i).x1 = 270
        Lines(i).y1 = 630 + (i - 1) * 25 * 15
        Lines(i).x2 = 270 + (mlength - 1) * 25 * 15
        Lines(i).y2 = 630 + (i - 1) * 25 * 15
        Lines(i).Visible = True
    Next i
    For i = 1 To mlength
        Load Lines(mheight + i)
        Lines(mheight + i).x1 = 270 + (i - 1) * 25 * 15
        Lines(mheight + i).y1 = 630
        Lines(mheight + i).x2 = 270 + (i - 1) * 25 * 15
        Lines(mheight + i).y2 = 630 + (mheight - 1) * 25 * 15
        Lines(mheight + i).Visible = True
    Next i
    Load Dot(1)
    Dot(1).Left = 270 + 3 * 25 * 15 - 30
    Dot(1).Top = 630 + 3 * 25 * 15 - 30
    Dot(1).Visible = True
    Load Dot(2)
    Dot(2).Left = 270 + (mlength - 4) * 25 * 15 - 30
    Dot(2).Top = 630 + 3 * 25 * 15 - 30
    Dot(2).Visible = True
    Load Dot(3)
    Dot(3).Left = 270 + 3 * 25 * 15 - 30
    Dot(3).Top = 630 + (mheight - 4) * 25 * 15 - 30
    Dot(3).Visible = True
    Load Dot(4)
    Dot(4).Left = 270 + (mlength - 4) * 25 * 15 - 30
    Dot(4).Top = 630 + (mheight - 4) * 25 * 15 - 30
    Dot(4).Visible = True
    Load Dot(5)
    Dot(5).Left = 270 + ((mlength - 1) \ 2) * 25 * 15 - 30
    Dot(5).Top = 630 + ((mheight - 1) \ 2) * 25 * 15 - 30
    Dot(5).Visible = True
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        k = i * mheight + j + 1
        Load Z(k)
        Z(k).Left = 270 + i * 25 * 15 - 10 * 15
        Z(k).Top = 630 + j * 25 * 15 - 10 * 15
        Z(k).ZOrder 0
    Next j
    Next i
    Dot(0).FillColor = vbRed
    Dot(0).BorderColor = vbRed
    Dot(0).ZOrder 0
    Mnew.Enabled = False
    Combo1.Enabled = False
    Combo3.Enabled = False
    Command1.Enabled = False
    mnuTrayRun.Enabled = False
    Form1.MousePointer = 12
    click = False
    Form1.Show
    AddToTray Me, mnuTray
    mainIcon = True
    auto = False
End Sub

Private Sub refreshMap()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim black As Long
    Dim white As Long
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        k = i * mheight + j + 1
        If mmap(i, j) = 0 Then
            Z(k).Visible = False
        Else
            If mmap(i, j) = 1 Then
                Z(k).FillColor = vbBlack
                black = black + 1
            ElseIf mmap(i, j) = 2 Then
                Z(k).FillColor = vbWhite
                white = white + 1
            Else
                Z(k).FillColor = vbYellow
            End If
            Z(k).Visible = True
        End If
    Next j
    Next i
    If black > white Then
        Combo3.Text = "白"
    Else
        Combo3.Text = "黑"
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If click = False Then Exit Sub
    x = x - 270 + 12 * 15
    y = y - 630 + 12 * 15
    If x < 0 Or y < 0 Then Exit Sub
    x = x \ (25 * 15)
    y = y \ (25 * 15)
    If x >= mlength Or y >= mheight Then Exit Sub
    Dot(0).Visible = False
    If Button = mmap(x, y) Then
        mmap(x, y) = 0
    Else
        mmap(x, y) = Button
    End If
    refreshMap
    If mmap(x, y) = 1 Then
        If checkWin(1, mmap, x, y) = True Then
            MsgBox "黑棋胜", vbInformation, "五子棋"
        ElseIf scanMap(1, mmap, x, y, True) = valueMin Then
            MsgBox "黑棋禁手", vbExclamation, "五子棋"
        End If
    ElseIf mmap(x, y) = 2 Then
        If checkWin(2, mmap, x, y) = True Then MsgBox "白棋胜", vbInformation, "五子棋"
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lastButton = Button
    lastX = x
    lastY = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    mstop = True
    auto = False
    Close #1
    unloader.Enabled = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Mabout_Click()
    Form2.Show , Form1
End Sub

Private Sub Mnew_Click()
    Dot(0).Visible = False
    Dim i As Long
    Dim j As Long
    For i = 0 To mlength - 1
    For j = 0 To mheight - 1
        mmap(i, j) = 0
        Z(i * mheight + j + 1).Visible = False
    Next j
    Next i
    Combo3.Text = "黑"
    setCaption "五子棋决策系统"
End Sub

Private Sub mnuTrayClose_Click()
    Unload Form1
End Sub

Private Sub mnuTrayHide_Click()
    If Form1.WindowState = vbMinimized Then
        Form1.WindowState = vbNormal
        Form1.Show
    Else
        Form1.WindowState = vbMinimized
    End If
End Sub

Private Sub mnuTrayRun_Click()
    If Command1.BackColor = &HFF00& Then
        clicker.Enabled = True
    Else
        mstop = True
        auto = False
    End If
End Sub

Private Sub Mstatistics_Click()
    MsgBox "棋谱数量：" & dbCount & vbNewLine & "棋谱命中：" & dbHit & "/" & dbUse & vbNewLine & "缓存命中：" & cacheHit & "/" & cacheUse, vbInformation, "五子棋"
End Sub

Private Sub showCount_Timer()
    If loading Then
        setCaption "数据加载中  " & (rindex \ (records \ 100)) & "%"
    Else
        Dim status As String
        status = smain & "."
        Dim i As Long
        For i = 0 To mdeep
            status = status & sdetail(i)
        Next i
        setCaption "五子棋决策系统  " & status
    End If
    If mainIcon = True Then
        SetTrayIcon Form2.Icon
        mainIcon = False
    Else
        SetTrayIcon Form1.Icon
        mainIcon = True
    End If
End Sub

Private Sub setCaption(title As String)
    Form1.Caption = title
    SetTrayTip title
End Sub

Private Sub unloader_Timer()
    unloader.Enabled = False
    RemoveFromTray
    End
End Sub
