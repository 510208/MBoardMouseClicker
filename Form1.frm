VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton StopBtn 
      Caption         =   "�{������(&S)"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "�{������(&S)"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�û���ܦb�̤W�h(��ĳ)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Value           =   1  '�֨�
      Width           =   2655
   End
   Begin VB.Frame BeforeRun_Frm 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "����e(&B)"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '����
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   3240
         List            =   "Form1.frx":000D
         TabIndex        =   15
         Text            =   "��ܫ��s"
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton RunToKB 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���檽�����Q���U(&K)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton RunTurn 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "������w����(&T)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CheckBox DownClockChk 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "�˼�3������(&D)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox BreakTxt 
         Appearance      =   0  '����
         Height          =   270
         Left            =   3360
         TabIndex        =   5
         Text            =   "100"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TurnTxt 
         Appearance      =   0  '����
         Height          =   270
         Left            =   1200
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "�Q���U���s(&B)�G"
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
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���涡�j(&B)�G"
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
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "���榸��(&T)�G"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label AlertLbl 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�p�����I���涡�j�����O�@��(ms)�A�@��������@�d�@��A���ƭȥ����]�w�A�_�h��G�ۭt"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Menu Tool 
      Caption         =   "�\��(&T)"
      Begin VB.Menu Start 
         Caption         =   "�s���I��"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) '�n��API���
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' Sleep���
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long ' �������U�䪺��
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '�w��̤W�hAPI
'�I���{�� mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'�k���I���{�� mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
Const MOUSEEVENTF_LEFTDOWN = &H2 '������U
Const MOUSEEVENTF_LEFTUP = &H4 '����u�_
Const MOUSEEVENTF_MIDDLEDOWN = &H20 '������U
Const MOUSEEVENTF_MIDDLEUP = &H40 '����u�_
Const MOUSEEVENTF_MOVE = &H1 '���ʹ���
Const MOUSEEVENTF_ABSOLUTE = &H8000 '���Ы��w����y��
Const MOUSEEVENTF_RIGHTDOWN = &H8 '�k����U
Const MOUSEEVENTF_RIGHTUP = &H10 '�k��u�_

Private Sub Form_Load()
    TurnSpin.Value = 0
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '�̤W�h���
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '�����̤W�h���
    End If
End Sub

Private Sub start_Click()
    If BreakTxt.Text <> "" And Val(BreakTxt.Text) >= 5 Then
        If RunTurn.Value = True And Val(TurnTxt.Text) > 0 Or RunToKB.Value = True Then
            RunOnTop.Show
        Else
            MsgBox "�Ъ`�N�I�T�w���榸�Ƥj��0�A�_�h�N�y�����~�I", vbCritical, "�]�w���~�I"
        End If
    Else
        MsgBox "�Ъ`�N�I���涡�j�����W�L5�@��A�ä��i���ťաI", vbCritical, "�]�w���~�I"
    End If
End Sub

Private Sub StartBtn_Click()
    start_Click
End Sub

Private Sub TopChk_Click()
    If TopChk.Value Then
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '�̤W�h���
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '�����̤W�h���
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

