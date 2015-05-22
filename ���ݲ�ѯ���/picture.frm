VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   18600
   FillColor       =   &H000080FF&
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   11310
   ScaleWidth      =   18600
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1680
      TabIndex        =   4
      Top             =   9360
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      ForeColor       =   &H0000C000&
      Height          =   9135
      Left            =   120
      ScaleHeight     =   9075
      ScaleWidth      =   18315
      TabIndex        =   3
      Top             =   0
      Width           =   18375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   9360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14640
      TabIndex        =   1
      Top             =   9480
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const YMAX = 50
Private Sub Command1_Click()
Picture1.Cls
Call ShowSumNum
End Sub

Private Sub Command2_Click()
Picture1.Cls
Call ShowBlueNum
End Sub

Private Sub Form_Load()
Call ShowSumNum

'Picture1.Line (0, Picture1.Height / 2)-(Picture1.Width, Picture1.Height / 2), vbWhite
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 200 Then
    Label1.Caption = "SUM: " + Str(Int(X)) + "    " + "COO: " + Str(SumReg(Int(X)))
End If

End Sub

Private Sub ShowSumNum()
Dim i As Integer
Picture1.Scale (48, 0)-(154, YMAX + 1)
Picture1.DrawWidth = 1
For i = 0 To YMAX
    Picture1.Line (50, i)-(154, i), &HCC00
 
Next i
For i = 0 To YMAX
   Picture1.CurrentX = 0
   Picture1.CurrentY = i
   Picture1.Print Str(YMAX - i)
Next i

For i = 50 To 160
  If (i Mod 10) = 0 Then
    Picture1.Line (i, 0)-(i, YMAX), &HFF66CC
End If
 Next i
Picture1.DrawWidth = 5
For i = 50 To 154
       
       Picture1.PSet (i, YMAX - SumReg(i)), vbRed
       Picture1.CurrentY = Picture1.CurrentY - 1
       Picture1.CurrentX = Picture1.CurrentX - 1
       Picture1.Print Str(i)
        
Next i
Picture1.DrawWidth = 15
For i = TotalNum - 5 To TotalNum - 1
    Picture1.PSet (SumNum(i), YMAX), vbRed
    Picture1.Print Str(SumNum(i)) + Str(TotalNum - i)
Next i
Picture1.DrawWidth = 10
For i = TotalNum - 10 To TotalNum - 6
    Picture1.PSet (SumNum(i), YMAX - 1), vbRed
Next i
Picture1.DrawWidth = 10
For i = TotalNum - 15 To TotalNum - 11
    Picture1.PSet (SumNum(i), YMAX - 1), vbBlue
Next i
For i = TotalNum - 20 To TotalNum - 16
    Picture1.PSet (SumNum(i), YMAX - 2), vbRed
Next i
For i = TotalNum - 25 To TotalNum - 21
    Picture1.PSet (SumNum(i), YMAX - 2), vbBlue
Next i
End Sub
Private Sub ShowBlueNum()
Dim i As Integer
        Picture1.DrawWidth = 1
Picture1.Scale (0, 0)-(32, 26)
For i = 0 To 26
    Picture1.Line (0, i)-(32, i), &HCC00
Next i
For i = 0 To 32
    If (i Mod 4) = 0 Then
        Picture1.DrawWidth = 2
        Picture1.Line (i, 0)-(i, 26), &HFFFF
    Else
        Picture1.DrawWidth = 1
  
   ' If (i Mod 2) = 0 Then
   '     Picture1.Line (i, 0)-(i, 26), &HFF66CC
  '  Else
    '    Picture1.Line (i, 0)-(i, 26), &HFFFF66
 ' End If
    End If
 Next i
 Picture1.DrawWidth = 8
 For i = 0 To 25
 
    Picture1.PSet (BlueNum(TotalNum - 26 + i), i), vbRed
    Picture1.Print Str(BlueNum(TotalNum - 26 + i))
 Next i
    
End Sub

