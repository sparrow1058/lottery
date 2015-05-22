VERSION 5.00
Begin VB.Form LoseForm 
   Caption         =   "遗漏偏差数据统计"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   14865
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Tail --"
      Height          =   615
      Index           =   1
      Left            =   10440
      TabIndex        =   15
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tail ++"
      Height          =   615
      Index           =   0
      Left            =   10440
      TabIndex        =   14
      Top             =   9480
      Width           =   1335
   End
   Begin VB.ListBox List7 
      Height          =   6180
      Left            =   8160
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hot S"
      Height          =   1095
      Left            =   13440
      TabIndex        =   9
      Top             =   9000
      Width           =   855
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HMC"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      TabIndex        =   8
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Blue Ball"
      Height          =   615
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Red Ball"
      Height          =   615
      Index           =   0
      Left            =   8880
      TabIndex        =   6
      Top             =   8880
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   3300
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List5 
      Height          =   3300
      Left            =   6720
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox List6 
      Height          =   6180
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   3480
      Width           =   8055
   End
   Begin VB.ListBox List4 
      Height          =   4380
      Left            =   8880
      TabIndex        =   2
      Top             =   4560
      Width           =   5775
   End
   Begin VB.ListBox List3 
      Height          =   4380
      Left            =   8880
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.ListBox List1 
      Height          =   3300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "LoseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HMCSamll As Integer
Public TailSet As Integer

Private Sub ShowList2()
Dim TempStr As String
Dim List2str As String
Dim List3str As String
Dim List4Str As String
Dim i, j As Integer
Dim DoubleNum(1 To 5) As String
Dim TempNum(1 To 6) As Integer
Dim NumSum10(9) As Integer
Dim ANum(100) As Integer
Dim Sum10(9) As Integer
For i = 0 To List1.ListCount - 2
    TempStr = Mid(List1.List(i), 22, 18)
    For j = 1 To 6
        TempNum(j) = Mid(TempStr, j * 3 - 2, 2)
        If Val(TempNum(j)) < 10 Then
            Sum10(Val(TempNum(j))) = Sum10(Val(TempNum(j))) + 1
        End If
    Next j
    
    For j = 1 To 5
        If TempNum(1) = TempNum(j + 1) Then
            DoubleNum(1) = Str(TempNum(1))
         End If
    Next j
    For j = 1 To 4
        If TempNum(2) = TempNum(j + 2) Then
            DoubleNum(2) = Str(TempNum(2))
         End If
    Next j
    For j = 1 To 3
        If TempNum(3) = TempNum(j + 3) Then
            DoubleNum(3) = Str(TempNum(3))
         End If
    Next j
    For j = 1 To 2
        If TempNum(4) = TempNum(j + 4) Then
            DoubleNum(4) = Str(TempNum(4))
         End If
    Next j
    If TempNum(5) = TempNum(6) Then
        DoubleNum(5) = Str(TempNum(5))
  End If
  
    
    'For j = 1 To 6
     '   SUM10(tempnum(j)) = SUM10(tempnum(j)) + 1
        
    
    
    For j = 1 To 5
        List2str = List2str + DoubleNum(j) + " "
       
        If DoubleNum(j) <> "" And Val(DoubleNum(j)) < 10 Then
            NumSum10(Val(DoubleNum(j))) = NumSum10(Val(DoubleNum(j))) + 1
        End If
         DoubleNum(j) = ""
      Next j
    
    If i Mod 10 = 0 Then '''error
     For j = 0 To 9
        List3str = List3str + Str(j) + ":" + Str(NumSum10(j)) + "|"
        List4Str = List4Str + Str(j) + ":" + Format(Str(Sum10(j)), "00") + "|"
        NumSum10(j) = 0
        Sum10(j) = 0
     Next j
        List3.AddItem List3str
        'List3.AddItem "---------------------------------------------------------------------"
       List4.AddItem List4Str
        List3str = ""
        List4Str = ""
    End If
    
    
    List2.AddItem List2str
    List2str = ""
  
    
    
Next i

     For j = 0 To 9
        List3str = List3str + Str(j) + ":" + Str(NumSum10(j)) + "|"
         List4Str = List4Str + Str(j) + ":" + Format(Str(Sum10(j)), "00") + "|"
        NumSum10(j) = 0
         Sum10(j) = 0
     Next j
        List3.AddItem List3str
        List3.AddItem "---------------------------------------------------------------------"
        List3.AddItem "剩余期数统计 " + Str((i - 1) Mod 10)
        List4.AddItem List4Str
    


End Sub




Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    Call ShellExecute(hwnd, "open", "http://taobao.starlott.com/ssq/jb.html", vbNullString, vbNullString, 1)
ElseIf Index = 1 Then
    Call ShellExecute(hwnd, "open", "http://www.cpyjy.com/ssq_lan.html", vbNullString, vbNullString, 1)
End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
For i = 0 To 2
 If Option1(i).Value = True Then
    HMCSamll = i + 3
    Exit For
 End If
Next i
 
HMC.Show
End Sub



Private Sub Command3_Click(Index As Integer)
Dim i As Integer
If Index = 0 Then
  TailSet = TailSet + 1
Else
  TailSet = TailSet - 1
End If
If TailSet > 9 Then
 TailSet = 0
End If
If TailSet < 0 Then
 TailSet = 9
End If
For i = 0 To List6.ListCount - 1
 List6.Selected(i) = False
Next i

If TailSet = 0 Then
 List6.Selected(10 + TailSet - 1) = True
 List6.Selected(20 + TailSet - 1) = True
  List6.Selected(30 + TailSet - 1) = True
End If

If TailSet > 0 And TailSet < 4 Then
 List6.Selected(TailSet - 1) = True
 List6.Selected(10 + TailSet - 1) = True
 List6.Selected(20 + TailSet - 1) = True
  List6.Selected(30 + TailSet - 1) = True
End If
If TailSet > 3 Then
 List6.Selected(TailSet - 1) = True
 List6.Selected(10 + TailSet - 1) = True
 List6.Selected(20 + TailSet - 1) = True

End If





 
  
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To Form2.mainlist2.ListCount - 1
    List1.AddItem Form2.mainlist2.List(i)
Next i
Call ShowList2
List1.ListIndex = List1.ListCount - 2

For i = 0 To Form2.yllist.ListCount - 1
    List5.AddItem Form2.yllist.List(i)
Next i
'‘List6.AddItem "淘宝红球统计" '"http://taobao.starlott.com/ssq/jb.html"
'List6.AddItem "彩票篮球统计" '"http://www.cpyjy.com/ssq_lan.html"
TailSet = 0
Call ShowMTime(0)
Call ShowList7
End Sub

Private Sub ShowList7()
Dim i, j As Integer
Dim Sum7 As Integer
Dim TempStr As String
For i = 0 To List6.ListCount - 1
 TempStr = Replace(List6.List(i), "|", "")
 TempStr = Replace(TempStr, " ", "")
 For j = 4 To Len(TempStr) Step 2
     Sum7 = Sum7 + Val(Mid(TempStr, j, 2))
 Next j
 List7.AddItem Str(Sum7)
 Sum7 = 0
 Next i
 
 

End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then
    List3.ListIndex = 0
End If

If List1.ListIndex < List1.ListCount - 1 And List1.ListIndex > 0 Then
    List2.ListIndex = List1.ListIndex
    List3.ListIndex = ((List1.ListIndex - 1) / 10) + 1
   ' List4.ListIndex = List3.ListIndex
   ' LList.Caption = Str(List1.ListIndex)
    ShowMTime (List1.ListCount - 1 - List1.ListIndex)
If List6.ListCount = 33 Then
For i = 0 To List6.ListCount - 1
    List6.Selected(i) = False
Next i
TempStr = Left(Replace(Trim(List1.List(List1.ListIndex)), " ", ""), 12)

For i = 1 To 12 Step 2      'leaf
  SelNum = Val(Mid(TempStr, i, 2))
 List6.Selected(SelNum - 1) = True
 Next i
End If
End If


End Sub
Private Function ShowLostCheck()

End Function


Private Function ShowMTime(reduce As Integer)
Dim SaveData(1 To 1000, 1 To 33) As Integer
Dim LoseNum(1 To 1000, 1 To 33) As Integer
Dim TData(7) As String
Dim NumStr(1 To 33) As String
Dim NumTimes As Integer
Dim TempStr As String
Dim i, j As Integer
Dim TotalNum As Integer
TotalNum = Form2.AllData.ListCount - reduce
List6.Clear
For i = 1 To TotalNum
    TempStr = Form2.AllData.List(i - 1)
    Call GetNum(TempStr, TData())
 For j = 1 To 6
    SaveData(i, Val(TData(j))) = 1
 Next j
Next i
For j = 1 To 33
       If SaveData(1, j) = 1 Then
            LoseNum(1, j) = 0
        
        Else
            LoseNum(1, j) = 1
        
        End If
    For i = 2 To TotalNum
        If SaveData(i, j) = 1 Then
            LoseNum(i, j) = 0
        Else
            LoseNum(i, j) = LoseNum(i - 1, j) + 1
        End If
             Next i


Next j
TempStr = ""
    For j = 1 To 33
       NumStr(j) = Format(Str(LoseNum(TotalNum, j)), "00") + "| "
    For i = TotalNum To 2 Step -1
            If LoseNum(i, j) = 0 Then
               NumTimes = NumTimes + 1
                NumStr(j) = Format(Str(LoseNum(i - 1, j)), "00") + "| " + NumStr(j)
            End If
          
            If NumTimes > 18 Then
                NumTimes = 0
                Exit For
            End If
            
           
    Next i
           List6.AddItem Format(Str(j), "00") + ": |" + NumStr(j)
        
    Next j
    
    
    
  ' saveData(i,val(mid(tempstr,j


End Function

Private Sub List5_Click()
Dim i As Integer
Dim SelNum As Integer
Dim TempStr As String
Call ShowMTime(0)
If List6.ListCount = 33 Then
For i = 0 To List6.ListCount - 1
    List6.Selected(i) = False
Next i
TempStr = RTrim(List5.List(List5.ListIndex))
For i = 5 To Len(TempStr) Step 3
  SelNum = Val(Mid(TempStr, i, 2))
 List6.Selected(SelNum - 1) = True
 Next i
End If
End Sub

Private Sub List6_DblClick()
Dim i As Integer
Dim DCNum As Integer
For i = 0 To List6.ListCount - 1
    List6.Selected(i) = False
Next i

  DCNum = List6.ListIndex / 5
  If DCNum < 6 Then
  For i = 1 To 5 Step 2
   List6.Selected(DCNum * 5 + i - 1) = True
  Next i
  Else
   List6.Selected(31) = True
  End If
End Sub
