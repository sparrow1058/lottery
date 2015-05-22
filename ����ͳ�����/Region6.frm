VERSION 5.00
Begin VB.Form Region6 
   Caption         =   "详细区间数值统计"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form3"
   ScaleHeight     =   8595
   ScaleWidth      =   11145
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List10 
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List9 
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List8 
      Height          =   240
      Left            =   8520
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.ListBox List7 
      Height          =   1680
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox List6 
      Height          =   3300
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   10455
   End
   Begin VB.ListBox List5 
      Height          =   1860
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   10455
   End
   Begin VB.ListBox List4 
      Height          =   1140
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ListBox List3 
      Height          =   1140
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   7920
      Width           =   9615
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "30 期数据统计"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   10680
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "10 期数据统计"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   0
      Left            =   10680
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   855
   End
End
Attribute VB_Name = "Region6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ShowLabel3()
Dim i, j As Integer
Dim NType(8) As Integer
Dim NumFlag As Integer
Dim Tempstr As String
For i = 0 To List2.ListCount - 1
 NumFlag = 0
 Tempstr = Replace(List2.List(i), " ", "")
  Tempstr = Replace(Tempstr, "|", "")
  For j = 1 To 7
   If Val(Mid(Tempstr, j, 1)) = 4 Then
    NumFlag = NumFlag + 1000
   ElseIf Val(Mid(Tempstr, j, 1)) = 3 Then
    NumFlag = NumFlag + 100
  ElseIf Val(Mid(Tempstr, j, 1)) = 2 Then
    NumFlag = NumFlag + 10
  ElseIf Val(Mid(Tempstr, j, 1)) = 1 Then
    NumFlag = NumFlag + 1
  End If
  Next j

Select Case (NumFlag Mod 10)
  Case 6:
    NType(0) = NType(0) + 1
  Case 4:
    NType(1) = NType(1) + 1
  Case 3:
    NType(2) = NType(2) + 1
  Case 2:
   If (NumFlag \ 1000) = 1 Then
    NType(3) = NType(3) + 1  ''1*2 +4
   Else
    NType(4) = NType(4) + 1   ''1*2+ 2*2
   End If
  Case 1:
   If (NumFlag \ 100) = 1 Then
     NType(5) = NType(5) + 1   ' 1 +2 +3
   Else
     NType(6) = NType(6) + 1 ''1+5
   End If
  Case 0
    If (NumFlag \ 100) = 2 Then
     NType(7) = NType(7) + 1   ' 3 +3
    Else
     NType(8) = NType(8) + 1 ' 2+2 +2

   End If
  
  
End Select


Next i
Label3.Caption = " 1*6  | 1*4+2 | 1*3+3 | 1*2+4 |1*2+2*2| 1+2+3 |  1+5  |  3*2  |  2*3  |" + vbCrLf

For i = 0 To 8

Label3.Caption = Label3.Caption + "  " + Format(Str(NType(i)), "000") + "  |"

Next i


End Sub
Private Sub ShowList6()
Dim i, j, k As Integer
Dim Num1, Num2, Num3 As Integer
Dim List5Str(1 To 3) As String
Dim List6Str As String
For i = List5.ListCount - 2 To 3 Step -3
 List5Str(3) = TrimList(List5.List(i))
 List5Str(2) = TrimList(List5.List(i - 1))
 List5Str(1) = TrimList(List5.List(i - 2))
 For j = 1 To 35
 k = k + 1
 Num3 = Val(Mid(List5Str(3), 2 * j - 1, 2))
 Num2 = Val(Mid(List5Str(2), 2 * j - 1, 2))
 Num1 = Val(Mid(List5Str(1), 2 * j - 1, 2))
 List6Str = List6Str + Format(Str(Num1 + Num2 + Num3), "00") + " "
 If k = 5 Then
  List6Str = List6Str + "|"
  k = 0
 End If
 
 
 Next j
 List6.AddItem List6Str
 List6Str = ""
' List5Str(1) = Replace(List5Str, " ", "")
 'list5str(

Next i



End Sub
Private Function TrimList(Tempstr As String) As String
Tempstr = Replace(Tempstr, " ", "")
Tempstr = Replace(Tempstr, "|", "")
TrimList = Tempstr
End Function

Private Sub Form_Load()
Dim Tempstr As String
Dim Temp05 As String
Dim TData(7) As String
Dim TempRegStr As String
Dim Region6(6) As Integer
Dim RegionSum(6) As Integer
Dim RegionSum10(6) As Integer
Dim RegList5(6) As Integer
Dim RegList52(6, 5) As Integer
Dim RegList510(6, 5) As Integer
Dim i, j, k, k2, M05 As Integer
For i = (Form2.AllData.ListCount Mod 10) To Form2.AllData.ListCount - 1
 k = k + 1
 Tempstr = Form2.AllData.List(i)
 
 List1.AddItem Tempstr
 Call GetNum(Tempstr, TData())
 For j = 1 To 6
  'Region6((Val(tdata(j)) - 1#) / 5) = Region6((Val(tdata(j)) - 1#) / 5) + 1
 If Val(TData(j)) < 6 Then
    Region6(0) = Region6(0) + 1
 ElseIf Val(TData(j)) < 11 Then
    Region6(1) = Region6(1) + 1
 ElseIf Val(TData(j)) < 16 Then
    Region6(2) = Region6(2) + 1
 ElseIf Val(TData(j)) < 21 Then
    Region6(3) = Region6(3) + 1
 ElseIf Val(TData(j)) < 26 Then
    Region6(4) = Region6(4) + 1
 ElseIf Val(TData(j)) < 31 Then
    Region6(5) = Region6(5) + 1
 Else
    Region6(6) = Region6(6) + 1
 End If
 Next j
For j = 0 To 6
    RegList5(j) = RegList5(j) + Region6(j)
    RegList510(j, Region6(j)) = RegList510(j, Region6(j)) + 1
    RegList52(j, Region6(j)) = RegList52(j, Region6(j)) + 1
    TempRegStr = TempRegStr + Str(Region6(j)) + "|"
    RegionSum(j) = RegionSum(j) + Region6(j)
    
    Region6(j) = 0
Next j

List2.AddItem TempRegStr
TempRegStr = ""

'If (i Mod 5) = 0 Then
If k = 5 Then
 k2 = k2 + 1
    For j = 0 To 6
     TempRegStr = TempRegStr + Str(RegionSum(j)) + "|"
     RegionSum10(j) = RegionSum10(j) + RegionSum(j)
    RegionSum(j) = 0
    Next j
    List3.AddItem TempRegStr
    k = 0
End If
TempRegStr = ""
If k2 = 2 Then

    For j = 0 To 6
     TempRegStr = TempRegStr + Format(Str(RegionSum10(j)), "00") + "|"
     RegionSum10(j) = 0
     For M05 = 0 To 4
      Temp05 = Temp05 + Format(Str(RegList510(j, M05)), "00") + " "
      RegList510(j, M05) = 0
     Next M05
      
      Temp05 = Temp05 + "|"
    Next j
    List4.AddItem TempRegStr
    List5.AddItem Temp05
    Temp05 = ""
    k2 = 0
End If
TempRegStr = ""
Tempstr = ""
Next i
If Not k Then
For j = 0 To 6
     TempRegStr = TempRegStr + Str(RegionSum(j)) + "|"
     RegionSum10(j) = RegionSum10(j) + RegionSum(j)
     Tempstr = Tempstr + Format(Str(RegionSum10(j)), "00") + "|"
Next j    'tempstr = tempstr + Format(Str(RegionSum10(j)), "00") + "|"
    List3.AddItem TempRegStr
    List4.AddItem Tempstr
End If
'List4.AddItem "---------------------"
'List4.AddItem "____________________|"

'List4.AddItem "05|10|15|20|25|30|33|"
List1.ListIndex = List1.ListCount - 1

List5.AddItem "_01-05区间分布_|_06-10区间分布_|_11-15区间分布_|_16-20区间分布_|_21-25区间分布_|_26-30区间分布_|_31-33区间分布_|"

For j = 0 To 6
Tempstr = ""
 For i = 0 To 4
  Tempstr = Tempstr + Str(RegList52(j, i)) + " "
 Next i
 List7.AddItem Str(j + 1) + " " + Tempstr + Str(RegList5(j)) + "  "


Next j
List4.ListIndex = List4.ListCount - 1
List5.ListIndex = List5.ListCount - 1
Call ShowList6
List8.Visible = False
List8.Top = List5.Top
List8.Width = List5.Width
List8.Left = List5.Left
List8.Height = List5.Height

List9.Top = List6.Top
List9.Width = List6.Width - 3500
List9.Left = List6.Left
List9.Height = List6.Height

List10.Top = List6.Top
List10.Width = 3500
List10.Left = List9.Left + List9.Width
List10.Height = List9.Height



'Call ShowList9

End Sub


Private Sub List1_Click()

List2.ListIndex = List1.ListIndex
Label1.Caption = Str(List1.ListIndex + 1)
List3.ListIndex = List1.ListIndex / 5  ''leaf
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
Dim i As Integer
Dim Tempstr  As String
Label3.Caption = ""
For i = List2.ListIndex - 9 To List2.ListIndex
   Tempstr = Replace(List2.List(i), "|", "")
   Tempstr = Replace(Tempstr, " ", "")
'  Label3.Caption = Label3.Caption + TempStr + " "
 
 Next i
End Sub

Private Sub List2_DblClick()
If List5.Visible = True Then
   List5.Visible = False
   List6.Visible = False
   List8.Visible = True
   List9.Visible = True
   List10.Visible = True
   
  
   Call ShowList8
     Call ShowList9
   Call ShowList10
Else
   List5.Visible = True
   List6.Visible = True
   List8.Visible = False
    List9.Visible = False
    List9.Clear
   List8.Clear
    List10.Clear
    List10.Visible = False
    
End If

End Sub
Private Sub ShowList10()
Dim i, j, k, Total As Integer
Const All = 30
Dim Linestr As String
Dim TempData(All) As String
Dim TempData4(All) As String
Dim DataSum(All) As Integer
Dim SumStr(All) As String
Dim Result(All) As Str_Sum
Dim Result4(All) As Str_Sum
Dim Tempstr As String
Total = List2.ListCount - 1
For i = 0 To All
'For i = List2.ListCount - 30 To List2.ListCount - 1
   Tempstr = Replace(List2.List(Total - 30 + i), "|", "")
   Tempstr = Replace(Tempstr, " ", "")
   TempData(i) = Str(Val(Mid(Tempstr, 1, 1)) + Val(Mid(Tempstr, 2, 1)) + Val(Mid(Tempstr, 3, 1))) _
                + Str(Val(Mid(Tempstr, 4, 1)) + Val(Mid(Tempstr, 5, 1)) + Val(Mid(Tempstr, 6, 1))) + Mid(Tempstr, 7, 1)
   
   TempData(i) = Replace(TempData(i), " ", "")
   TempData4(i) = Str(Val(Mid(Tempstr, 1, 1)) + Val(Mid(Tempstr, 2, 1))) + Str(Val(Mid(Tempstr, 3, 1)) + Val(Mid(Tempstr, 4, 1))) _
                + Str(Val(Mid(Tempstr, 5, 1)) + Val(Mid(Tempstr, 6, 1)))
   
   TempData4(i) = Replace(TempData4(i), " ", "")
   
 Next i
 
 Call SumCount(TempData(), Result())
 Call SumCount(TempData4(), Result4())
 For i = 0 To All
 If Result4(i).Sum1 Then
    Linestr = Result4(i).str1 + "  " + Str(Result4(i).Sum1)
  End If
 If Result(i).Sum1 Then
    Linestr = Linestr + "  " + Result(i).str1 + "  " + Str(Result(i).Sum1)
 End If
 If Not (Linestr = "") Then
    List10.AddItem Linestr
     Linestr = ""
 End If
Next i
 
 
 
 
 
 
 
  

End Sub
Private Sub ShowList9()
Dim i, j, k As Integer
Dim Tempstr As String
Dim Reg22(0 To 33, 3) As Integer
Dim Reg22Str(3) As String
Dim Reg222Str(2) As String
Dim Reg222(0 To 333, 2) As Integer
Dim Reg35Str(10) As String
 For i = List2.ListCount - 18 To List2.ListCount - 1
 Tempstr = Replace(List2.List(i), " ", "")
 Tempstr = Replace(Tempstr, "|", "")
 Reg22(Val(Mid(Tempstr, 1, 2)), 1) = Reg22(Val(Mid(Tempstr, 1, 2)), 1) + 1
 Reg22(Val(Mid(Tempstr, 3, 2)), 2) = Reg22(Val(Mid(Tempstr, 3, 2)), 2) + 1
 Reg22(Val(Mid(Tempstr, 5, 2)), 3) = Reg22(Val(Mid(Tempstr, 5, 2)), 3) + 1

Reg222(Val(Mid(Tempstr, 1, 3)), 1) = Reg222(Val(Mid(Tempstr, 1, 3)), 1) + 1
Reg222(Val(Mid(Tempstr, 4, 3)), 2) = Reg222(Val(Mid(Tempstr, 4, 3)), 2) + 1
  Next i
  For j = 1 To 3
  For i = 0 To 22
   
   If (i Mod 10) < 4 Then
   ' If Reg22(i, j) <> 0 Then
      Reg22Str(j) = Reg22Str(j) + Format(Str(i), "00") + " " + Format(Str(Reg22(i, j)), "0") + " | "
      
  ' End If
  End If
  Reg22(i, j) = 0
  Next i
   List9.AddItem Reg22Str(j)
  
  Next j
  List9.AddItem "----------------------------------------------------------------------------------"
  
  For j = 1 To 2
   For i = 0 To 233
     
     If (i Mod 10) < 4 And ((i \ 10) Mod 10) < 4 Then
            k = (i Mod 10) + ((i \ 10) Mod 10) + (i \ 100)
            Reg35Str(k) = Reg35Str(k) + Format(Str(i), "000") + " " + Format(Str(Reg222(i, j)), "0") + "| "
      End If
       Next i
       For k = 1 To 6
            List9.AddItem Str(k) + " | " + Reg35Str(k)
            Reg35Str(k) = ""
       Next k
        List9.AddItem "---------------------------------------------------------------------------"
    
    Next j
  


End Sub
Private Sub ShowList8()
Dim i, j, k As Integer
Dim Tempstr, LineStr1, LineStr2, LineStr3, LineStr4 As String
Dim L8Str As String
Dim RegionApp(400, 8) As String
Dim AppearStr1(40), AppearStr2(40, 2) As String
Dim AppearTimes1(40), AppearTimes2(40, 2) As Integer
Dim APCount1, APCount2(2) As Integer
Dim RegAppSum(6) As Integer
Dim Totals As Integer
Totals = List2.ListCount
For i = 0 To List2.ListCount - 1
 Tempstr = Replace(List2.List(i), " ", "")
 Tempstr = Replace(Tempstr, "|", "")
 LineStr1 = Str(Val(Mid(Tempstr, 1, 1)) + Val(Mid(Tempstr, 2, 1))) + " : " + Str(Val(Mid(Tempstr, 3, 1)) + Val(Mid(Tempstr, 4, 1))) + " : " + Str(Val(Mid(Tempstr, 5, 1)) + Val(Mid(Tempstr, 6, 1)) + Val(Mid(Tempstr, 7, 1)))
 LineStr2 = Mid(Tempstr, 1, 2) + " : " + Mid(Tempstr, 3, 2) + " : " + Mid(Tempstr, 5, 2) + " : " + Mid(Tempstr, 7, 1)
 LineStr3 = Mid(Tempstr, 1, 3) + " : " + Mid(Tempstr, 4, 3) + " : " + Mid(Tempstr, 7, 1)
 LineStr4 = Str(Val(Mid(Tempstr, 1, 1)) + Val(Mid(Tempstr, 2, 1)) + Val(Mid(Tempstr, 3, 1))) + " : " + Str(Val(Mid(Tempstr, 4, 1)) + Val(Mid(Tempstr, 5, 1)) + Val(Mid(Tempstr, 6, 1))) + " : " + Str(Val(Mid(Tempstr, 7, 1)))

 List8.AddItem LineStr1 + " |  " + LineStr2 + " | " + LineStr3 + " | " + LineStr4
 LineStr4 = Replace(LineStr4, " ", "")
 LineStr4 = Replace(LineStr4, ":", "")
 RegionApp(i, 5) = LineStr4
 
 RegionApp(i, 0) = Mid(Tempstr, 1, 2)
 RegionApp(i, 1) = Mid(Tempstr, 3, 2)
 RegionApp(i, 2) = Mid(Tempstr, 5, 2)
 
 RegionApp(i, 3) = Mid(Tempstr, 1, 3)
 RegionApp(i, 4) = Mid(Tempstr, 4, 3)
 
 
Next i
For i = 0 To Totals - 1          '//leaf continue app
 For j = i + 1 To Totals - 1
    For k = 0 To 5
 If (RegionApp(i, k) <> "-") Then
  If RegionApp(i, k) = RegionApp(j, k) Then
    RegAppSum(k) = RegAppSum(k) + 1
    RegionApp(j, k) = "-"
  End If
 End If
  Next k
 Next j
  For k = 0 To 5
   If (RegionApp(i, k) <> "-") Then
       L8Str = RegionApp(i, k) + "  " + Str(RegAppSum(k)) ''leaf
      RegAppSum(k) = 0
   End If
   Next k
  ' List8.AddItem L8Str
   L8Str = ""
 
Next i

List8.ListIndex = List8.ListCount - 1

End Sub


Private Sub List3_Click()
'Label2.Caption = Str(List3.ListIndex)
'List5.ListIndex = List3.ListIndex
End Sub

