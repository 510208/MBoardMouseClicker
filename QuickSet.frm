VERSION 5.00
Begin VB.Form QuickSet 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "快速鍵設定"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   2985
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton OK 
      Caption         =   "確定(&O)"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '平面
      Height          =   300
      ItemData        =   "QuickSet.frx":0000
      Left            =   1440
      List            =   "QuickSet.frx":0016
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "開始/停止(&S)："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "QuickSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
    QuickSet.Hide
End Sub
