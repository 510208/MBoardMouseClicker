VERSION 5.00
Begin VB.Form RunOnTop 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "正在執行中"
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   4560
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "永遠顯示在最上層(建議)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Value           =   1  '核取
      Width           =   2655
   End
   Begin VB.CommandButton StopBtn 
      Appearance      =   0  '平面
      Caption         =   "停止(&S)"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "次"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "當前執行到"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "RunOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) '聲明API函數
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' Sleep函數
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long ' 偵測按下鍵的值
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '定位最上層API
'點擊程式 mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'右鍵點擊程式 mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
Const MOUSEEVENTF_LEFTDOWN = &H2 '左鍵按下
Const MOUSEEVENTF_LEFTUP = &H4 '左鍵彈起
Const MOUSEEVENTF_MIDDLEDOWN = &H20 '中鍵按下
Const MOUSEEVENTF_MIDDLEUP = &H40 '中鍵彈起
Const MOUSEEVENTF_MOVE = &H1 '移動鼠標
Const MOUSEEVENTF_ABSOLUTE = &H8000 '鼠標指定絕對座標
Const MOUSEEVENTF_RIGHTDOWN = &H8 '右鍵按下
Const MOUSEEVENTF_RIGHTUP = &H10 '右鍵彈起
Dim EndInfo As Boolean
Dim KeyBoardKey

Private Sub Form_Activate()
    Select Case QuickSet.Combo1.ListIndex
        Case 0
            KeyBoardKey = vbKeyF1
        Case 1
            KeyBoardKey = vbKeyF2
        Case 2
            KeyBoardKey = vbKeyF5
        Case 3
            KeyBoardKey = vbKeyF7
        Case 4
            KeyBoardKey = vbKeyF11
        Case 5
            KeyBoardKey = vbKeyF12
        Case Else
            KeyBoardKey = vbKeyF1
    End Select
    Dim LoopNum As Double
    LoopNum = 0
    Sleep 3000
    Do Until GetAsyncKeyState(KeyBoardKey)
        Select Case Main.Combo1.ListIndex
            Case 0
                mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 '左鍵按下,彈起
            Case 1
                mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0 '中鍵按下，彈起
            Case 2
                mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0 '右鍵按下，彈起
            Case Else
                mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 '左鍵按下,彈起
        End Select
        LoopNum = LoopNum + 1
        Label1.Caption = Str(LoopNum)
        Sleep Val(Main.BreakTxt.Text)
        If EndInfo Then
            Exit Do
        End If
    Loop
    MsgBox "完成執行" & vbNewLine & "共執行了" & LoopNum & "次", vbInformation, "訊息"
End Sub

Private Sub Form_Load()
    Main.Hide
End Sub

Private Sub StopBtn_Click()
    EndInfo = True
    Main.Show
    Me.Hide
End Sub

Private Sub TopChk_Click()
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '最上層顯示
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '取消最上層顯示
    End If
End Sub
