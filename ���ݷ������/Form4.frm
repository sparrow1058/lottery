VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "分区遗漏统计"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   LinkTopic       =   "Form4"
   ScaleHeight     =   8550
   ScaleWidth      =   12570
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "区间遗漏 统计"
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegLost(7) As String
Dim RegLostStr As String

Private Sub Form_Load()
Dim i As Integer
For i = Form1.List1.ListCount - 20 To Form1.List1.ListCount - 1
 List1.AddItem Left(Form1.List1.List(i), 44)
Next i
End Sub
Private Sub ShowList2()
Dim tempstr As String
Dim i As Integer
For i = List1.ListCount - 1 To List1.ListCount - 20 Step -1
 tempstr = List1.List(i)

Next i


End Sub

Private Function CheckRegLost(IStr As String)



End Function
Private Function RegLostClear()
Dim i As Integer
For i = 1 To 7
 RegLost(i) = "          "
Next i

RegLostStr = ""
End Function

Private Function RegLostAdd()
Dim i As Integer
For i = 1 To 7
    RegLostStr = RegLostStr + "|" + RegLost(i)
Next i
End Function

