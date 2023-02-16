VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "MBoardMouseClicker"
   ClientHeight    =   2385
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4560
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton QuickKey 
      Caption         =   "快速鍵(&K)"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton StopBtn 
      Appearance      =   0  '平面
      Cancel          =   -1  'True
      Caption         =   "退出(&S)"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton StartBtn 
      Appearance      =   0  '平面
      Caption         =   "程式執行(&S)"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "永遠顯示在最上層(建議)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Value           =   1  '核取
      Width           =   2655
   End
   Begin VB.Frame BeforeRun_Frm 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "執行前(&B)"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '平面
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   1440
         List            =   "Form1.frx":000D
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox DownClockChk 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "倒數3秒後執行(&D)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox BreakTxt 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Text            =   "100"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "被按下按鈕(&B)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin MSForms.SpinButton BreakSpin 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   255
         Size            =   "450;450"
      End
      Begin VB.Label BreakLbl 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "執行間隔(&B)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label AlertLbl 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "小提醒！執行間隔的單位是毫秒(ms)，一秒鐘等於一千毫秒，此數值必須設定，否則後果自負"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu Tool 
      Caption         =   "功能(&T)"
      Begin VB.Menu Start 
         Caption         =   "連續點擊"
         Shortcut        =   {F6}
      End
      Begin VB.Menu NinSi87 
         Caption         =   "-"
      End
      Begin VB.Menu AboutProg 
         Caption         =   "關於(&A)"
      End
   End
End
Attribute VB_Name = "Main"
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

Private Sub AboutProg_Click()
    frmAbout.Show
End Sub

Private Sub Form_Load()
    Main.Show
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '最上層顯示
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '取消最上層顯示
    End If
End Sub

Private Sub QuickKey_Click()
    QuickSet.Show
End Sub

Private Sub start_Click()
    If BreakTxt.Text <> "" And Val(BreakTxt.Text) > 5 Then
        RunOnTop.Show
    Else
        MsgBox "請注意！執行間隔必須超過5毫秒，並不可為空白！", vbCritical, "設定錯誤！"
    End If
End Sub

Private Sub StartBtn_Click()
    start_Click
End Sub

Private Sub Timer1_Timer()
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
    If GetAsyncKeyState(KeyBoardKey) Then
        Label2.Caption = "87"
    Else
        Label2.Caption = "89787"
    End If
End Sub

Private Sub StopBtn_Click()
    End
End Sub

Private Sub TopChk_Click()
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '最上層顯示
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '取消最上層顯示
    End If
End Sub

Private Sub TurnSpin_Change()
    TurnTxt = Str(TurnSpin.Value)
End Sub

Private Sub TurnTxt_Change()
    On Error GoTo error9487:
    TurnSpin.Value = Val(TurnTxt.Text)
    Exit Sub
error9487:
    MsgBox "抱歉，您給定的參數值過大！", vbCritical, "錯誤！"
End Sub

Private Sub BreakSpin_Change()
    BreakTxt = Str(BreakSpin.Value)
End Sub

Private Sub BreakTxt_Change()
    BreakSpin.Value = Val(BreakTxt.Text)
End Sub

