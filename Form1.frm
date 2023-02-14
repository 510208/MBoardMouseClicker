VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton StopBtn 
      Caption         =   "程式停止(&S)"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "程式執行(&S)"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "永遠顯示在最上層(建議)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
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
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '平面
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   3240
         List            =   "Form1.frx":000D
         TabIndex        =   15
         Text            =   "選擇按鈕"
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton RunToKB 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "執行直到按鍵被按下(&K)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton RunTurn 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "執行指定次數(&T)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CheckBox DownClockChk 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "倒數3秒後執行(&D)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox BreakTxt 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   3360
         TabIndex        =   5
         Text            =   "100"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TurnTxt 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   1200
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "被按下按鈕(&B)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin MSForms.SpinButton BreakSpin 
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   255
         Size            =   "450;450"
      End
      Begin VB.Label BreakLbl 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "執行間隔(&B)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin MSForms.SpinButton TurnSpin 
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   255
         Size            =   "450;450"
      End
      Begin VB.Label TurnLbl 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "執行次數(&T)："
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label AlertLbl 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "小提醒！執行間隔的單位是毫秒(ms)，一秒鐘等於一千毫秒，此數值必須設定，否則後果自負"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Menu Tool 
      Caption         =   "功能(&T)"
      Begin VB.Menu Start 
         Caption         =   "連續點擊"
         Shortcut        =   {F5}
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

Private Sub Form_Load()
    TurnSpin.Value = 0
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '最上層顯示
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '取消最上層顯示
    End If
End Sub

Private Sub start_Click()
    If BreakTxt.Text <> "" And Val(BreakTxt.Text) >= 5 Then
        If RunTurn.Value = True And Val(TurnTxt.Text) > 0 Or RunToKB.Value = True Then
            RunOnTop.Show
        Else
            MsgBox "請注意！確定執行次數大於0，否則將造成錯誤！", vbCritical, "設定錯誤！"
        End If
    Else
        MsgBox "請注意！執行間隔必須超過5毫秒，並不可為空白！", vbCritical, "設定錯誤！"
    End If
End Sub

Private Sub StartBtn_Click()
    start_Click
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
    TurnSpin.Value = Val(TurnTxt.Text)
End Sub

Private Sub BreakSpin_Change()
    BreakTxt = Str(BreakSpin.Value)
End Sub

Private Sub BreakTxt_Change()
    BreakSpin.Value = Val(BreakTxt.Text)
End Sub

