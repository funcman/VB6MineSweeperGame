Attribute VB_Name = "Mod"
Public m(639) As Integer  '记录地点的属性
Public num As Integer '雷数
Public opennum
'Public Time As Integer
Public start As Boolean
Public score As Integer

Sub NewGame() '新游戏初始化
    num = 0
    opennum = 0
    'Time = 0
    start = False
    score = 0
    For i = 0 To 639
        m(i) = 0
        frmMain.Tile(i).Caption = ""
    Next i
    For i = 1 To 18
        For j = 1 To 30
            If Int(5 * Rnd) = 0 Then
                m(i * 32 + j) = 1
                num = num + 1
            End If
        Next j
    Next i
End Sub

Sub main()
    Call NewGame
    frmMain.Show
End Sub

Sub hack(Index As Integer)
    If start = False Then
        GoTo endsubhack
    End If
'    Call opentile(Index - 33)
    If frmMain.Tile(Index - 32).Caption = "" Then
        Call opentile(Index - 32)
    End If
 '   Call opentile(Index - 31)
    If frmMain.Tile(Index - 1).Caption = "" Then
        Call opentile(Index - 1)
    End If
    If frmMain.Tile(Index + 1).Caption = "" Then
        Call opentile(Index + 1)
    End If
 '   Call opentile(Index + 31)
    If frmMain.Tile(Index + 32).Caption = "" Then
        Call opentile(Index + 32)
    End If
  '  Call opentile(Index + 33)
    'If M(index - 32) <= 0 And frmMain.Tile(index - 33).Caption = "" Then
     '   frmMain.Tile(index - 32).Caption = 0
      '  hack (index - 32)
    'End If
    'If M(index - 1) <= 0 And frmMain.Tile(index - 1).Caption = "" Then
     '   frmMain.Tile(index - 1).Caption = 0
      '  hack (index - 1)
    'End If
    'If M(index + 1) <= 0 And frmMain.Tile(index + 1).Caption = "" Then
  '      frmMain.Tile(index + 1).Caption = 0
 '       hack (index + 1)
  '  End If
   ' If M(index + 32) <= 0 And frmMain.Tile(index + 32).Caption = "" Then
    '    frmMain.Tile(index + 32).Caption = 0
     '   hack (index + 32)
   ' End If
endsubhack:
End Sub

Sub opentile(Index As Integer)
   ' Call hack(index)
    If start = False Then
        GoTo endsubopentile
    End If
    Select Case Index
        Case 0 To 31
            GoTo endsubopentile
        Case 32, 63, 64, 95, 96, 127, 128, 159, 160, 191, 192, 223, 224, 255, 256, 287, 288, 319, 320, 351, 352, 383, 384, 415, 416, 447, 448, 479, 480, 511, 512, 543, 544, 575, 576, 607
            GoTo endsubopentile
        Case 608 To 639
            GoTo endsubopentile
    End Select
    Dim sum As Integer
    sum = 0
    If m(Index) = 1 Then
        frmMain.Tile(Index).Caption = "●"

        mv = MsgBox("GAME OVER", 0, "：（")
        Call NewGame
        start = False
        GoTo endsubopentile
    Else
        If m(Index) = 0 Then
            m(Index) = -1
        End If
        If m(Index - 33) > 0 Then
            sum = sum + 1
        End If
        If m(Index - 32) > 0 Then
            sum = sum + 1
        End If
        If m(Index - 31) > 0 Then
            sum = sum + 1
        End If
        If m(Index - 1) > 0 Then
            sum = sum + 1
        End If
        If m(Index + 1) > 0 Then
            sum = sum + 1
        End If
        If m(Index + 31) > 0 Then
            sum = sum + 1
        End If
        If m(Index + 32) > 0 Then
            sum = sum + 1
        End If
        If m(Index + 33) > 0 Then
            sum = sum + 1
        End If
        frmMain.Tile(Index).Caption = sum
    End If
    If sum = 0 And frmMain.Tile(Index).Caption <> "●" Then
        Call hack(Index)
    End If
endsubopentile:
End Sub

