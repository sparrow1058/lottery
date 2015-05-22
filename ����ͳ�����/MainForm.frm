VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Main form"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9795
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "双色球"
      Height          =   1815
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "22选5"
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
C225.Show
MainForm.Hide
End Sub

Private Sub Command2_Click()
Form2.Show
Me.Hide
End Sub
