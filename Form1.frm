VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "24点终极作弊器v2"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3495
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "蓝色 - 待命 // 绿色 - 成功"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Shape           =   3  'Circle
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "复制希尔薇的话即可自动识别，听到嘟一声或看到指示灯变绿后即可粘贴！ --- DuckSoft"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function optIt#(opt, aa#, bb#)
  Select Case opt
   Case 1: optIt = aa + bb
   Case 2: optIt = aa - bb
   Case 3: optIt = aa * bb
   Case 4:
   If bb = 0 Then
    optIt = -2000
   Else
    optIt = aa / bb
   End If
  End Select
End Function

Function transOp(inOp) As String
 Select Case inOp
  Case 1: transOp = "+"
  Case 2: transOp = "-"
  Case 3: transOp = "*"
  Case 4: transOp = "/"
  Case Else: Error (66666) ' you must be mad
 End Select
End Function
Function Get24Point$(ByVal aa%, ByVal bb%, ByVal cc%, ByVal dd%)
 Dim a#(1 To 7)
 
 a(1) = aa
 a(2) = bb
 a(3) = cc
 a(4) = dd
 
 If a(1) = 3 And a(2) = 7 And a(3) = 7 And a(4) = 9 Then
  Get24Point = "3*(9-(7/7))"
  Exit Function
 ElseIf a(1) = 1 And a(2) = 3 And a(3) = 4 And a(4) = 6 Then
  Get24Point = "6/(1-(3/4))"
  Exit Function
 ElseIf a(1) = 3 And a(2) = 3 And a(3) = 8 And a(4) = 8 Then
  Get24Point = "8/(3-(8/3))"
  Exit Function
 End If
 
 
 ' a tree
 For i = 1 To 4
  a(5) = optIt(i, a(1), a(2))
  If a(5) = -2000 Then GoTo ai
  For j = 1 To 4
   a(6) = optIt(j, a(3), a(4))
   If a(6) = -2000 Then GoTo aj
   For k = 1 To 4
    a(7) = optIt(k, a(5), a(6)) ' if error here, never get 24
    If a(7) < 24.1 And a(7) > 23.9 Then Get24Point = "(" & aa & transOp(i) & bb & ")" & transOp(k) & "(" & cc & transOp(j) & dd & ")":: Exit Function
   Next
aj:
  Next
ai:
 Next

 ' b tree
 For i = 1 To 4
  a(5) = optIt(i, a(1), a(2))
  If a(5) = -2000 Then GoTo bi
  For j = 1 To 4
   a(6) = optIt(j, a(5), a(3)) ' error maybe
   If a(6) = -2000 Then GoTo bj
   For k = 1 To 4
    a(7) = optIt(k, a(6), a(4))
    If a(7) = -2000 Then GoTo bk
    If a(7) < 24.1 And a(7) > 23.9 Then Get24Point = "((" & aa & transOp(i) & bb & ")" & transOp(j) & cc & ")" & transOp(k) & dd:: Exit Function
bk:
   Next
bj:
  Next
bi:
 Next
 

 ' d tree
 For i = 1 To 4
  a(5) = optIt(i, a(2), a(3))
  If a(5) = -2000 Then GoTo di
  For j = 1 To 4
   a(6) = optIt(j, a(5), a(4))
   If a(6) = -2000 Then GoTo dj
   For k = 1 To 4
    a(7) = optIt(k, a(1), a(6))
    If a(7) = -2000 Then GoTo dk
    If a(7) < 24.1 And a(7) > 23.9 Then Get24Point = aa & transOp(k) & "((" & bb & transOp(i) & cc & ")" & transOp(j) & dd & ")": Exit Function
dk:
   Next
dj:
  Next
di:
 Next

End Function

Function GetIt(a%, b%, c%, d%) As String
 Static strTemp As String
 Static f(1 To 4) As Integer
 f(1) = a
 f(2) = b
 f(3) = c
 f(4) = d
 
 For i = 1 To 4
  For j = 1 To 4
   If j <> i Then
    For k = 1 To 4
     If k <> j Then
      If k <> i Then
       For l = 1 To 4
        If l <> i Then
         If l <> j Then
          If l <> k Then
           strTemp = Get24Point(f(i), f(j), f(k), f(l))
           If strTemp <> "" Then
            GetIt = strTemp
            Exit Function
           End If
          End If
         End If
        End If
       Next
      End If
     End If
    Next
   End If
  Next
 Next
End Function



Private Sub Form_Load()
 '子类化
 preWinProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf wndproc)

    RegisterHotKey Me.hWnd, 0, 1 Or 2, vbKeyV

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

 Static strClip As String
 Static posL, posR As Integer
 
 Shape1.FillColor = vbBlue

 strClip = Clipboard.GetText
 If Strings.Left(strClip, 2) = "算 " Then Exit Sub
 
 posL = InStr(1, strClip, "字：")
 posR = InStr(1, strClip, vbLf)
 If posR = 0 Then
  posL = InStr(1, strClip, "：") + 1
  posR = Len(strClip)
 End If
 If posL >= posR Then Exit Sub
 
 strGo$ = Mid(strClip, posL + 2, posR - posL - 1)
 Debug.Print strGo
 a = Split(strGo, "、", 4)
 strAns = GetIt(CInt(a(0)), CInt(a(1)), CInt(a(2)), CInt(a(3)))
 If strAns <> "" Then
  Clipboard.Clear
  Clipboard.SetText "算 " & strAns
  Shape1.FillColor = vbGreen
  Beep
 End If
End Sub
