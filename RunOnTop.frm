VERSION 5.00
Begin VB.Form RunOnTop 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   Caption         =   "���b���椤"
   ClientHeight    =   1035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   4560
   Begin VB.CheckBox TopChk 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�û���ܦb�̤W�h(��ĳ)(&T)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Value           =   1  '�֨�
      Width           =   2655
   End
   Begin VB.CommandButton StopBtn 
      Appearance      =   0  '����
      Caption         =   "����(&S)"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "��e�����"
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
                mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 '������U,�u�_
            Case 1
                mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0 '������U�A�u�_
            Case 2
                mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0 '�k����U�A�u�_
            Case Else
                mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 '������U,�u�_
        End Select
        LoopNum = LoopNum + 1
        Label1.Caption = Str(LoopNum)
        Sleep Val(Main.BreakTxt.Text)
        If EndInfo Then
            Exit Do
        End If
    Loop
    MsgBox "��������" & vbNewLine & "�@����F" & LoopNum & "��", vbInformation, "�T��"
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
        IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '�̤W�h���
    Else
        IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3) '�����̤W�h���
    End If
End Sub
