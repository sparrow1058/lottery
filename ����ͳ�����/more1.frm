VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   14130
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List3 
      Height          =   4920
      Left            =   2880
      TabIndex        =   2
      Top             =   4920
      Width           =   3255
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   9780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 0 To HZForm.List1.ListCount - 1
    List1.AddItem HZForm.List1.List(i)
Next i
List1.ListIndex = List1.ListCount - 1
End Sub
