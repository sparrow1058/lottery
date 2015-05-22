VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form6"
   ScaleHeight     =   8670
   ScaleWidth      =   12735
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "和值分布图"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   7680
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   1920
      ScaleHeight     =   7035
      ScaleWidth      =   9795
      TabIndex        =   3
      Top             =   480
      Width           =   9855
   End
   Begin VB.ListBox List2 
      Height          =   7440
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllNums6 As Integer

Private Sub ShowSumSortPic()
Dim tempstr As String
Dim i, j As Integer
Dim MaxMin6 As MaxMin
Dim Nums(1000) As Integer
Dim NClow As Integer

Picture1.DrawWidth = 2
For i = 0 To List1.ListCount - 1
 
     Nums(i) = Val(List1.List(i))
Next i
MaxMin6 = GetMaxMin(Nums())


Picture1.Scale (MMin1.min - 1, MMin2.min - 1)-(MMin1.max + 1, MMin2.max + 1)
Picture_Y = MMin2.max
For i = 0 To List2.ListCount - 1
tempstr = List2.List(i)
' Picture1.CurrentX = Val(Left(tempstr, 3))
' Picture1.CurrentY = MMin2.max - NClow
 'Picture1.Print Left(tempstr, 3)
 Picture1.Line (Val(Left(tempstr, 3)), MMin2.max)-(Val(Left(tempstr, 3)), MMin2.max - Val(Right(tempstr, 2))), vbGreen
 Picture1.Print Right(tempstr, 2)


Next i
Picture1.DrawWidth = 13
'Picture1.PSet (95, 8), vbRed
For i = 1 To 20
tempstr = Right(Left(List1.List(List1.ListCount - i), 24), 3)
    For j = 0 To List2.ListCount - 1
 '       If Val(tempstr) = Num1(j) Then
  '          Picture1.PSet (Num1(j), MMin2.max - Num2(j)), RGB((i \ 5) * 255, (i \ 10) * 255, (i \ 15) * 255)
   '         Picture1.Print Str(i)
    '        Exit For
    '    End If
        
    Next j
    'Label6.Caption = Label6.Caption + Str(Num1(j))
    
Next i

End Sub
Private Sub Form_Load()
Dim i As Integer
For i = Form1.List1.ListCount - 1 To 0 Step -1
 List1.AddItem Right(Left(Form1.List1.List(i), 24), 3)
 AllNums6 = AllNums6 + 1
Next i

Call ShowSumLost
' List1.ListIndex = List1.ListCount - 1
End Sub


Private Sub ShowSumLost()
 Dim i, j As Integer
 For i = 0 To List1.ListCount - 1
  For j = i + 1 To List1.ListCount - 1
    If List1.List(i) = List1.List(j) Then
      List2.AddItem Str(j - i)
     Exit For
    End If
   Next j
    If j = List1.ListCount - 1 Then
     List2.AddItem "Too long"
    End If
Next i
 
 


End Sub


Private Sub List1_Click()
Dim i As Integer
Dim Count As Integer
Dim tempstr As String
If List1.ListIndex > List2.ListCount - 1 Then
 List2.ListIndex = List1.ListIndex
End If
For i = List1.ListIndex + 1 To List1.ListCount - 1
 If (List1.List(List1.ListIndex) = List1.List(i)) Then
  
   tempstr = tempstr + Str(i - Count - List1.ListIndex) + "  "
    Count = i
 End If
 Next i
 Label1.Caption = tempstr

End Sub
