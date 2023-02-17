VERSION 5.00
Begin VB.Form ClickTest 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "連點測試器"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4455
   StartUpPosition =   3  '系統預設值
   Begin VB.Label RightLbl 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "右鍵點擊次數："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label CentNumLbl 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label CentLbl 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "中鍵點擊次數："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LeftLbl 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "左鍵點擊次數："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "ClickTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LeftCount, MiddleCount, RightCount As Double
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
            LeftCount = LeftCount + 1
        Case vbMiddleButton
            MiddleCount = MiddleCount + 1
        Case vbRightButton
            RightCount = RightCount + 1
    End Select
    LeftLbl.Caption = Str(LeftCount)
    CentNumLbl.Caption = Str(MiddleCount)
    RightLbl.Caption = Str(RightCount)
End Sub
