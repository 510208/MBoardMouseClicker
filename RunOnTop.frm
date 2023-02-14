VERSION 5.00
Begin VB.Form RunOnTop 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "永遠顯示在最上層(建議)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   1  '核取
      Width           =   2655
   End
   Begin VB.CommandButton StopBtn 
      Appearance      =   0  '平面
      Caption         =   "停止(&S)"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   840
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

Private Sub Form_Load()
    If Main.DownClockChk.Value Then
        For i = 3 To 1 Step -1
            Sleep 1000
            Label2.Caption = i
        Next i
    End If
    If Main.RunTurn.Value Then
        For Turn = 0 To Main.TurnSpin.Value
            If GetAsyncKeyState(vbKeyF1) Then
                Form.Hide
                Exit Sub
            End If
            mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Next Turn
    Else
        Do Until Not (GetAsyncKeyState(vbKeyF1))
            mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Loop
    End If
End Sub

Private Sub TopChk_Click()
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '最上層顯示
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '取消最上層顯示
    End If
End Sub
