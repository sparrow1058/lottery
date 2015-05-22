VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   ScaleHeight     =   7950
   ScaleWidth      =   11175
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   7200
      Width           =   8895
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   7560
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   6000
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "大小奇偶分布"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "区间分布"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "遗漏区间分布"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "遗漏分布"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
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
      Height          =   6000
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   5295
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BData(2000) As Integer
Dim BTotal As Integer
Dim LostFlag(16) As Integer

Private Sub ShowList2(OP1 As Integer)
Dim i, j, k As Integer
Dim BRegion(4) As Integer
Dim LRegion(5) As Integer
Dim ODDBIG(4) As Integer
Dim BigSmall(2) As Integer
Dim OddEven(2) As Integer
Dim LineStr As String
Dim TempLost As String
List2.Clear
If OP1 = 0 Then
For i = List1.ListCount - 1 To List1.ListCount - 256 Step -1
 LineStr = LineStr + Mid(List1.List(i), 4, 2) + "-"
 j = j + 1
 If j = 16 Then
   j = 0
    List2.AddItem LineStr
   ' List2.AddItem "---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|"
   List2.AddItem ""
    LineStr = ""
 End If
Next i
List2.AddItem LineStr
End If

If OP1 = 1 Then
For i = 0 To List1.ListCount - 1
 TempLost = Val(Mid(List1.List(i), 4, 2))
  If TempLost < 6 Then
    LRegion(1) = LRegion(1) + 1
  ElseIf TempLost > 5 And TempLost < 11 Then
    LRegion(2) = LRegion(2) + 1
  ElseIf TempLost > 10 And TempLost < 21 Then
    LRegion(3) = LRegion(3) + 1
  ElseIf TempLost > 20 And TempLost < 31 Then
    LRegion(4) = LRegion(4) + 1
  ElseIf TempLost > 30 Then
   LRegion(5) = LRegion(5) + 1
 End If
 
 
 j = j + 1
 If j = 16 Then
   j = 0
  For k = 1 To 5
   LineStr = LineStr + Format(Str(LRegion(k)), "00") + "  |  "
   LRegion(k) = 0
  Next k
 
 List2.AddItem LineStr
 LineStr = ""
 End If

Next i
  For k = 1 To 5
   LineStr = LineStr + Format(Str(LRegion(k)), "00") + "  |  "
   Next k
List2.AddItem LineStr
End If




If OP1 = 2 Then
 For i = 0 To List1.ListCount - 1
  BRegion((Val(Left(List1.List(i), 2)) - 1) \ 4) = BRegion((Val(Left(List1.List(i), 2)) - 1) \ 4) + 1
  j = j + 1
   If j = 16 Then
      For k = 0 To 3
       LineStr = LineStr + Format(Str(BRegion(k)), "0") + "  | "
       BRegion(k) = 0
      Next k
      j = 0
      List2.AddItem LineStr
      LineStr = ""
  End If
  
Next i
       For k = 0 To 3
       LineStr = LineStr + Format(Str(BRegion(k)), "0") + "  | "
       BRegion(k) = 0
      Next k
 List2.AddItem LineStr

End If


If OP1 = 3 Then
''大小奇偶分布  1 小鸡 2 小鸥 3 大鸡 4 大鸥
For i = 0 To List1.ListCount - 1
  TempLost = Val(Left(List1.List(i), 2))
 If TempLost < 9 Then
    BigSmall(1) = BigSmall(1) + 1
    If TempLost Mod 2 Then
        ODDBIG(1) = ODDBIG(1) + 1
    Else
        ODDBIG(2) = ODDBIG(2) + 1
    End If
Else
    BigSmall(2) = BigSmall(2) + 1
    If TempLost Mod 2 Then
        ODDBIG(3) = ODDBIG(3) + 1
    Else
        ODDBIG(4) = ODDBIG(4) + 1
    End If
End If

    If TempLost Mod 2 Then
        OddEven(1) = OddEven(1) + 1
    Else
        OddEven(2) = OddEven(2) + 1
    End If





j = j + 1
 If j = 16 Then
    j = 0
    For k = 1 To 4
    LineStr = LineStr + Format(ODDBIG(k), "00") + " | "
    ODDBIG(k) = 0
    Next k
    LineStr = LineStr + "   | " + Format(Str(BigSmall(1)), "00") + " | " + Format(Str(BigSmall(2)), "00") + "  |  " + Format(Str(OddEven(1)), "00") + " | " + Format(Str(OddEven(2)), "00")
    BigSmall(1) = 0
    BigSmall(2) = 0
    OddEven(1) = 0
    OddEven(2) = 0
    List2.AddItem LineStr
    LineStr = ""
End If
Next i
   For k = 1 To 4
       LineStr = LineStr + Format(ODDBIG(k), "00") + " | "
    Next k
    List2.AddItem LineStr + "   | " + Format(Str(BigSmall(1)), "00") + " | " + Format(Str(BigSmall(2)), "00") + "  |  " + Format(Str(OddEven(1)), "00") + " | " + Format(Str(OddEven(2)), "00")
'    List2.AddItem "SmallOdd|SmallEven|BigOdd|BigEven|"
List2.AddItem ""
    List2.AddItem "小鸡小鸥 大鸡 大鸥 |      大小比    |   奇偶比"

End If






End Sub
Private Sub ShowList3()
  Dim LineS(5) As String
Dim i As Integer
 For i = 1 To 16
  List3.AddItem Format(Str(i), "00") + "    " + Format(Str(LostFlag(i)), "00")
Next i
   For i = 1 To 16
    If LostFlag(i) < 6 Then
        LineS(1) = LineS(1) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 5 And LostFlag(i) < 10 Then
        LineS(2) = LineS(2) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 9 And LostFlag(i) < 20 Then
        LineS(3) = LineS(3) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    
    If LostFlag(i) > 19 And LostFlag(i) < 30 Then
        LineS(4) = LineS(4) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 29 Then
        LineS(5) = LineS(5) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    Next i
   For i = 1 To 5
    List3.AddItem Format(Str(i), "0") + " " + LineS(i)
  Next i
    

End Sub
Private Sub Form_Load()
Dim TempStr As String
Dim i As Integer
Open App.Path & "\GoldData.txt" For Input As #1
While Not EOF(1)
    Line Input #1, TempStr
    BData(i) = Val(Right(Replace(TempStr, " ", ""), 2))
    i = i + 1
Wend
 BTotal = i
Close #1
'List1.AddItem Format(Str(BData(0)), "00")
'Call CheckLost(BData(0))
For i = 0 To BTotal - 1
  List1.AddItem Format(Str(BData(i)), "00") + " " + Format(Str(LostFlag(BData(i))), "00") + " " + Format(Str((BData(i) - 1) \ 4), "0")
 Call CheckLost(BData(i))
    
Next i
List1.ListIndex = List1.ListCount - 1
Form3.Caption = "总计" + " " + Str(List1.ListCount) + "期"
Call ShowList3
Call Option1_Click(0)

Text1.Text = TRHead

End Sub
Private Sub CheckLost(InNum As Integer)
 Dim i As Integer
 For i = 1 To 16
   LostFlag(i) = LostFlag(i) + 1
Next i
 LostFlag(InNum) = 0
End Sub

Private Sub Option1_Click(Index As Integer)
ShowList2 (Index)
End Sub
