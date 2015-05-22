VERSION 5.00
Begin VB.Form SaveForm 
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   10065
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   1215
      Left            =   1800
      TabIndex        =   1
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   5655
   End
End
Attribute VB_Name = "SaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path + "save.txt" For Output As #1
  
Do While Not EOF(1)
    Line Input #1, Tempstr
    List1.AddItem Left(Tempstr, 21)
Loop


End Sub

Private Sub Form_Load()
Label1.Caption = Left(Form2.AllData.List(Form2.AllData.ListCount - 1), 5)
End Sub

