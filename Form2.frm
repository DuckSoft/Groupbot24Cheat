VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "猜数字终极作弊器 彩蛋版"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "重设猜数字"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭猜数字作弊并开启24点作弊"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开启猜数字作弊并关闭24点作弊"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "恭喜你发现了这个彩蛋, enjoy it!"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, r As Integer
Const multiplier As Double = 0.5


Private Sub Command1_Click(Index As Integer)
On Error Resume Next
 Select Case Index
  Case 0:
   Form1.Timer1.Enabled = False
   Form2.Timer1.Enabled = True
   Form2.Command1(0).Enabled = False
   Form2.Command1(1).Enabled = True
  Case 1:
   Form1.Timer1.Enabled = True
   Form2.Timer1.Enabled = False
   Form2.Command1(0).Enabled = True
   Form2.Command1(1).Enabled = False
 End Select
End Sub

Private Sub ResetIt()
 l = 1
 r = 999
 Clipboard.Clear
 Clipboard.SetText "我猜 500"
 Label2.Caption = "RESET OK. INITIAL 500."
 Beep
 Beep
 Beep
 
End Sub

Private Sub Command2_Click()
 ResetIt
End Sub

Private Sub Form_Load()
 ResetIt
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Unload Form1
End Sub

Private Sub Timer1_Timer()
 Static clipStr As String
 Static numPosL, numPosR As Integer
 Static numIn As Integer
 Static boolGreater As Boolean
 
 clipStr = Clipboard.GetText()
 
 If Strings.Left(clipStr, 4) = "我猜" Then Exit Sub
 If InStr(1, clipStr, "恭喜") <> 0 Then ResetIt
 If Strings.Left(clipStr, 5) = "reset" Then ResetIt
 
 numPosL = InStr(1, clipStr, "【")
 numPosR = InStr(1, clipStr, "】")
 
 If numPosL >= numPosR Then Exit Sub
 If numPosR = 0 Then Exit Sub
 If Len(clipStr) < numPosR + 6 Then Exit Sub
 
 numIn = CInt(Mid(clipStr, numPosL + 1, numPosR - numPosL - 1))
 
 Select Case Mid(clipStr, numPosR + 6, 1)
  Case "大": boolGreater = True
  Case "小": boolGreater = False
  Case Else: Exit Sub
 End Select
 
 Label2.Caption = "INPUT " & numIn & " with " & IIf(boolGreater, "大", "小")
 
 If numIn < l Or numIn > r Then 'out of range
  Exit Sub
 Else
 
  If boolGreater = False Then
   l = numIn
  Else
   r = numIn
  End If
  Clipboard.Clear
  
  Clipboard.SetText ("我猜 " + CStr(CInt(l + (r - l) \ 2)))
  Beep
  Label2.Caption = "CALCULATED NUM IS " & CInt(l + (r - l) \ 2)
  Exit Sub
 End If
End Sub

