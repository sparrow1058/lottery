VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form3"
   ScaleHeight     =   7290
   ScaleWidth      =   11190
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List4 
      Height          =   3840
      Left            =   8160
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
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
      Height          =   3120
      Left            =   8160
      TabIndex        =   6
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   6000
      Width           =   5775
      Begin VB.OptionButton Option1 
         Caption         =   "余3"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3区间"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "尾号分布"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "差值分布"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "大小奇偶分布"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
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
      Width           =   5895
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

Private Sub AddSubSum()
Dim tempstr As String
Dim i As Integer
For i = List2.ListCount - 1 To List2.ListCount - 3
 tempstr = tempstr + List2.List(i)
 Next i
 
 



End Sub

Private Sub ShowList2(OP1 As Integer)
Dim i, j, k As Integer
Dim BRegion(4) As Integer
Dim LRegion(5) As Integer
Dim ODDBIG(4) As Integer
Dim BigSmall(2) As Integer
Dim OddEven(2) As Integer
Dim TempValue As Integer
Dim LineStr As String
Dim TempLost As String
Dim RangeNum(0 To 31)       '-15 ---+15
Dim SNum(20) As Integer
Dim SubSum(20) As Integer

'for zbr  type
Dim TempBlue As Integer
Dim zbr As ZhiBiao

List2.Clear
If OP1 = 0 Then
For i = List1.ListCount - 1 To List1.ListCount - 256 Step -1
 LineStr = Mid(List1.List(i), 4, 2) + "-" + LineStr
Next i
 
 
For j = 1 To Len(LineStr) Step 30


   
    List2.AddItem Mid(LineStr, j, 30)
   ' List2.AddItem "---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|"

 
 Next j
 
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


If OP1 = 4 Then
 j = 0
 LineStr = ""
For i = List1.ListCount - 300 To List1.ListCount - 1
    TempValue = Val(Left(List1.List(i), 2)) - Val(Left(List1.List(i - 1), 2))
    RangeNum(TempValue + 15) = RangeNum(TempValue + 15) + 1
    LineStr = LineStr + " " + Format(Str(TempValue), "@@@")
    j = j + 1
    If j = 10 Then
        List2.AddItem LineStr
         j = 0
        LineStr = ""
    End If
Next i
List2.AddItem LineStr
j = 0
LineStr = ""
For i = 0 To 15
 LineStr = Format(Str(i), "@@@") + " " + Format(Str(RangeNum(i + 15)), "@@@") + "      |     " + Format(Str(0 - i), "@@@") + " " + Format(Str(RangeNum(15 - i)), "@@@")

 List4.AddItem LineStr



Next i

End If

'尾号分布统计
If OP1 = 5 Then
For i = List1.ListCount - 100 To List1.ListCount - 1
 TempBlue = Val(Left(List1.List(i), 2))
 Call CheckZB(TempBlue, zbr)
 List2.AddItem ShowZBR(zbr)
 
Next i

End If

'3区间分布
If OP1 = 6 Then
For i = List1.ListCount - 100 To List1.ListCount - 1
 TempBlue = Val(Left(List1.List(i), 2))
 If TempBlue < 6 Then
 LineStr = "       ▲   |              |               |   "  '★▲"
 ElseIf TempBlue > 10 Then
  LineStr = "            |              |      ▲      |   "
Else
  LineStr = "            |     ▲      |               |   "
End If
List2.AddItem LineStr
LineStr = ""

Next i
List2.AddItem "  01-05   |    06-10    |    11-16    |     3区间分布”"
List2.ListIndex = List2.ListCount - 1
End If

'余3 分布统计
 If OP1 = 7 Then
For i = List1.ListCount - 100 To List1.ListCount - 1
 TempBlue = Val(Left(List1.List(i), 2))
 TempBlue = TempBlue Mod 3
 If TempBlue = 0 Then
 LineStr = "       ▲   |              |               |   "  '★▲"
 ElseIf TempBlue = 2 Then
  LineStr = "            |              |      ▲      |   "
Else
  LineStr = "            |     ▲      |               |   "
End If
List2.AddItem LineStr
LineStr = ""

Next i
List2.AddItem "3  6  9 12 15 | 1 4 7 10 13 16  |2 5 8 11 14 |     3区间分布”"
List2.ListIndex = List2.ListCount - 1


End If
End Sub
Private Sub ShowList3()
  Dim LineS(5) As String
Dim i As Integer

   For i = 1 To 16

    If LostFlag(i) < 10 Then
         List3.AddItem Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 9 And LostFlag(i) < 20 Then
        LineS(1) = LineS(1) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    
    If LostFlag(i) > 19 And LostFlag(i) < 30 Then
        LineS(2) = LineS(2) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 29 And LostFlag(i) < 40 Then
        LineS(3) = LineS(3) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    If LostFlag(i) > 39 Then
        LineS(4) = LineS(4) + Format(Str(i), "00") + "-" + Format(Str(LostFlag(i)), "00") + " "
    End If
    
    
    Next i
  
   For i = 1 To 4
    List3.AddItem LineS(i)
  Next i
    

End Sub
Private Sub Form_Load()
Dim tempstr As String
Dim i As Integer
Open App.Path & "\GoldData.txt" For Input As #1
While Not EOF(1)
    Line Input #1, tempstr
    BData(i) = Val(Right(Replace(tempstr, " ", ""), 2))
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

