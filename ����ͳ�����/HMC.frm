VERSION 5.00
Begin VB.Form HMC 
   Caption         =   "数据热冷门分布统计"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form4"
   ScaleHeight     =   8280
   ScaleWidth      =   10125
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   9135
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   5535
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   5520
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "HMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const HMCNum = 20
Dim LoseHMC(33, HMCNum) As Integer
Dim HMCStr(33, HMCNum) As String
Dim LoseT(33)  As Integer
Private Sub ShowList4()
Dim i, j, k As Integer
Dim TempStr, TS2 As String
Dim HotN(30) As Integer
Dim MidN, CoolN As Integer
For i = 0 To 50
 TempStr = Replace(List2.List(i), " ", "")
 For j = 1 To 6
  TS2 = Mid(TempStr, j * 3 - 2, 3)
  If Left(TS2, 1) = "C" Then
    CoolN = CoolN + 1
  ElseIf Left(TS2, 1) = "M" Then
    MidN = MidN + 1
  Else
    HotN(Val(Right(TS2, 2))) = HotN(Val(Right(TS2, 2))) + 1
 End If
 Next j
k = k + 1
If k = 3 Then       ''the Num
  k = 0
  TempStr = ""
  For j = 1 To 10
   If Not HotN(j) Then
    TempStr = TempStr + "H" + Format(Str(j), "0") + ":" + Format(Str(HotN(j)), "0") + " "
    HotN(j) = 0
   End If
  Next j
  List4.AddItem TempStr + "M:" + Str(MidN) + " " + "C:" + Str(CoolN)
  MidN = 0
  CoolN = 0
  
 End If

Next i


End Sub

Private Sub Form_Load()

Dim i, j As Integer
Dim TempStr As String
For i = LoseForm.List1.ListCount - 2 To 0 Step -1
    List1.AddItem LoseForm.List1.List(i)
Next i

For i = 1 To 33
TempStr = LoseForm.List6.List(i - 1)
For j = 1 To HMCNum
 
 LoseHMC(i, j) = Mid(TempStr, 2 + j * 4, 2)
 
Next j
Next i
Call GetHMCStr

Call ShowList2
Call ShowList3
Call ShowList4

'' Tmpe =33: |05| 04| 04| 18| 04| 14| 03| 08| 00| 00| 00| 03| 07| 22| 02| 00| 00| 01| 01| 04|
End Sub
Private Sub ShowList2()
Dim i, j As Integer
Dim TempStr As String
Dim Data(6) As Integer

For i = 0 To List1.ListCount - 1
 TempStr = List1.List(i)
 Call GetData1(TempStr, Data())
  TempStr = ""
 For j = 1 To 6
  LoseT(Data(j)) = LoseT(Data(j)) + 1
  If LoseT(Data(j)) <= 20 Then
  TempStr = TempStr + HMCStr(Data(j), 20 - LoseT(Data(j))) + " "
  End If
 Next j
   List2.AddItem TempStr

Next i

End Sub
Private Function GetData1(Lstr As String, Data() As Integer)
Dim i As Integer
 For i = 1 To 6
  Data(i) = Val(Mid(Lstr, i * 3 - 2, 2))
 Next i
End Function

Private Function GetHMCStr()
Dim i, j As Integer
For i = 1 To 33
If LoseHMC(i, 1) > 10 Then
  HMCStr(i, 1) = "C01"
 ElseIf LoseHMC(i, 1) < 6 Then
  HMCStr(i, 1) = " M01"
 Else
  HMCStr(i, 1) = "H01"
End If
For j = 2 To HMCNum

'' 检验大于10次长期遗漏
If LoseHMC(i, j) >= 10 Then

 If Left(HMCStr(i, j - 1), 1) = "C" Then
   HMCStr(i, j) = "C" + Format(Str(Val(Right(HMCStr(i, j - 1), 2) + 1)), "00")
 Else
    HMCStr(i, j) = "C01"
 End If
End If

''检验小于10次，大于5次遗漏
If LoseHMC(i, j) > LoseForm.HMCSamll And LoseHMC(i, j) < 10 Then
 If Left(HMCStr(i, j - 1), 1) = "M" Then
        HMCStr(i, j) = "M" + Format(Str(Val(Right(HMCStr(i, j - 1), 2) + 1)), "00")
 Else
        HMCStr(i, j) = "M01"
 End If
End If
 ''检验少于5次的遗漏
If LoseHMC(i, j) <= LoseForm.HMCSamll Then
 If Left(HMCStr(i, j - 1), 1) = "H" Then
   HMCStr(i, j) = "H" + Format(Str(Val(Right(HMCStr(i, j - 1), 2) + 1)), "00")
 'ElseIf Left(HMCStr(i, j - 1), 1) = "C" Then
  '  HMCStr(i, j) = "H00"
 ''leaf  此处有点问题，考虑解决方法
 Else
    HMCStr(i, j) = "H01"
 End If
End If
Next j

Next i

End Function

Private Sub ShowList3()
Dim i As Integer
Dim str0, str1, str2, str3, str4, str5, str6, str7, str8, str9, StrO, strM, strC As String
 For i = 1 To 33
 If Left(HMCStr(i, 20), 1) = "H" Then
  If HMCStr(i, 20) = "H01" Then
   If LoseHMC(i, 20) = 0 Then
       str0 = str0 + Format(Str(i), "00") + " "
      
    Else
       str1 = str1 + Format(Str(i), "00") + " "
   End If
  ElseIf HMCStr(i, 20) = "H02" Then
   str2 = str2 + Format(Str(i), "00") + " "
  ElseIf HMCStr(i, 20) = "H03" Then
   str3 = str3 + Format(Str(i), "00") + " "
  ElseIf HMCStr(i, 20) = "H04" Then
   str4 = str4 + Format(Str(i), "00") + " "
  ElseIf HMCStr(i, 20) = "H05" Then
   str5 = str5 + Format(Str(i), "00") + " "
 ElseIf HMCStr(i, 20) = "H06" Then
   str6 = str6 + Format(Str(i), "00") + " "
'ElseIf HMCStr(i, 20) = "H07" Then
 '  str7 = str7 + Format(Str(i), "00") + " "
'ElseIf HMCStr(i, 20) = "H08" Then
 '  str8 = str8 + Format(Str(i), "00") + " "
'ElseIf HMCStr(i, 20) = "H09" Then
 '  str9 = str9 + Format(Str(i), "00") + " "
 
  Else
   StrO = StrO + Format(Str(i), "00") + " "
  End If
 End If
  
  
 If Left(HMCStr(i, 20), 1) = "M" Then
   strM = strM + Format(Str(i), "00") + " "
 End If
 If Left(HMCStr(i, 20), 1) = "C" Then
   strC = strC + Format(Str(i), "00") + " "
 End If
Next i
List3.AddItem "H0T0: " + str0
List3.AddItem "H0T1: " + str1
List3.AddItem "H0T2: " + str2
List3.AddItem "H0T3: " + str3
List3.AddItem "H0T4: " + str4
List3.AddItem "H0T5: " + str5
List3.AddItem "H0T6: " + str6
'List3.AddItem "H0T7: " + str7
'List3.AddItem "H0T8: " + str8
'List3.AddItem "H0T9: " + str9


List3.AddItem "HOTO: " + StrO
List3.AddItem "MIDD: " + strM
List3.AddItem "COOL: " + strC


End Sub


Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
Dim TempStr As String
Dim i As Integer
For i = 0 To 32
 LoseForm.List6.Selected(i) = False
Next i
TempStr = Replace(List3.List(List3.ListIndex), " ", "")
If Len(TempStr) > 6 Then
  TempStr = Right(TempStr, Len(TempStr) - 5)
End If
For i = 1 To Len(TempStr) / 2
 LoseForm.List6.Selected(Val(Mid(TempStr, 2 * i - 1, 2)) - 1) = True
Next i
 
End Sub

