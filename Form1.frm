VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Main 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CheckBox Check1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�û���ܦb�̤W�h(��ĳ)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Value           =   1  '�֨�
      Width           =   2775
   End
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�û���ܦb�̤W�h(��ĳ)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Value           =   1  '�֨�
      Width           =   2775
   End
   Begin VB.Frame BeforeRun_Frm 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "����e(&B)"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox BreakTxt 
         Appearance      =   0  '����
         Height          =   270
         Left            =   3360
         TabIndex        =   5
         Text            =   "100"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TurnTxt 
         Appearance      =   0  '����
         Height          =   270
         Left            =   1200
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin MSForms.SpinButton BreakSpin 
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   240
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
         Top             =   240
         Width           =   1335
      End
      Begin MSForms.SpinButton TurnSpin 
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
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
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�p�����I���涡�j�����O�@��(ms)�A�@��������@�d�@��A���ƭȥ����]�w�A�_�h��G�ۭt"
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   2880
      TabIndex        =   9
      Top             =   840
      Width           =   1575
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
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long ' �H
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
    For Turn = 0 To 100
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Next Turn
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

