VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   Icon            =   "DBloader.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   630
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   15
      Top             =   15
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dataError As Boolean
Private black, white As Long

Private Sub Form_Activate()
    If Dir(files(UBound(files)), vbSystem + vbHidden) <> vbNullString Then
        SetAttr files(UBound(files)), vbNormal
        Kill files(UBound(files))
    End If
    Open files(UBound(files)) For Binary As #2
    memFull = False
    Open files(0) For Binary As #1
        If LOF(1) \ 45 > M Then
            dbLarge = True
            ReDim db(16 * M)
        Else
            dbLarge = False
            ReDim db(M)
        End If
    Close #1
    dbMax = 0
    dbCount = 0
    trecords = 0
    Dim i As Long
    Dim j As Long
    Dim b As Byte
    For i = 0 To UBound(files) - 1
        If Trim(files(i)) <> "" Then
            Label1.Caption = files(i)
            Label2.Caption = "0%"
            Open files(i) For Binary As #1
                records = LOF(1) \ 45
                trecords = trecords + records
                For rindex = 0 To records - 1
                    If rindex Mod 10000 = 0 Then Label2.Caption = Format(rindex / (records / 100), "0.00") & "%"
                    DoEvents
                    Get #1, , data
                    dataError = False
                    black = 0
                    white = 0
                    For j = 0 To 44 '每个字节存了5个数据，3进制累加
                        b = data(j)
                        checkData b \ 81
                        b = b Mod 81
                        checkData b \ 27
                        b = b Mod 27
                        checkData b \ 9
                        b = b Mod 9
                        checkData b \ 3
                        b = b Mod 3
                        checkData b
                    Next j
                    If (black = 0 And white = 0) Or black - white > 1 Or white > black Then dataError = True
                    If dataError = False Then
                        If inDB(data, True) = False Then Put #2, , data
                    End If
                Next rindex
            Close #1
        End If
    Next i
    Close #2
    Form3.Hide
    If memFull = True Then MsgBox "内存不足，不能处理所有数据", vbInformation, "数据处理错误"
    Dim message As String
    For i = 0 To UBound(files) - 1
        If Trim(files(i)) <> "" Then
            message = message & files(i) & vbNewLine
        End If
    Next i
    MsgBox message & trecords & "-->" & dbCount, vbInformation, "五子棋棋谱整理器"
    End
End Sub

Private Sub checkData(ByVal data As Byte)
    If dataError = True Then Exit Sub
    If data > 2 Then
        dataError = True
    ElseIf data = 1 Then
        black = black + 1
    ElseIf data = 2 Then
        white = white + 1
    End If
End Sub
