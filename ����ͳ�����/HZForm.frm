VERSION 5.00
Begin VB.Form HZForm 
   Caption         =   "和值，奇偶，区间分布"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   13965
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   420
      Index           =   3
      Left            =   4800
      TabIndex        =   27
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox List4 
      Height          =   2580
      Index           =   2
      Left            =   12120
      TabIndex        =   25
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   420
      Index           =   2
      Left            =   4800
      TabIndex        =   24
      Top             =   3360
      Width           =   3255
   End
   Begin VB.ListBox List4 
      Height          =   2580
      Index           =   1
      Left            =   9600
      TabIndex        =   23
      Top             =   0
      Width           =   4335
   End
   Begin VB.ListBox List2 
      Height          =   3300
      Index           =   1
      Left            =   4800
      TabIndex        =   22
      Top             =   0
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "100期数据统计"
      Height          =   3855
      Left            =   7440
      TabIndex        =   13
      Top             =   5400
      Width           =   6495
      Begin VB.ListBox List5 
         Height          =   1860
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5 Check"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9360
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.ListBox List42 
         Height          =   1680
         Left            =   4800
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
      Begin VB.ListBox List6 
         Height          =   1860
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Height          =   1860
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.ListBox List41 
         Height          =   3660
         Left            =   3360
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.ListBox List22 
         Height          =   3660
         Left            =   2160
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.ListBox List21 
         Height          =   3660
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "区间分布统计"
         Height          =   375
         Index           =   2
         Left            =   8400
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "和值分布"
      Height          =   1575
      Left            =   0
      TabIndex        =   10
      Top             =   5400
      Width           =   7455
      Begin VB.Label Label1 
         Height          =   1215
         Left            =   -120
         TabIndex        =   11
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "奇偶，大小分布"
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   7080
      Width           =   7455
      Begin VB.Label Label5 
         Caption         =   "区间分布(B:S):"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label3 
         Caption         =   "奇偶分布(O:E):"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label4 
         Caption         =   "大小分布(B:S):"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   6855
      End
   End
   Begin VB.ListBox List4 
      Height          =   5280
      Index           =   0
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "21 to 183"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "次数102"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ListBox List3 
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "HZForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SumtimesBack(21 To 183) As Integer
Dim OddTimes(0 To 1000) As Integer
Dim BigTimes(0 To 1000) As Integer
Dim Coo5 As Boolean
Private Sub ShowReg()
Dim i, j, k As Integer
Dim Reg1(6), Reg2(6), Reg3(6) As Integer
Dim Reg1A(6), Reg2A(6), Reg3A(6) As Integer

Dim TempStr1, TempStr2, TempStr3 As String
Dim TempStr As String
 For i = (List4(0).ListCount Mod 10) + 1 To List4(0).ListCount - 1
   TempStr = Trim(List4(0).List(i))
   TempStr = Replace(TempStr, " ", "")
   TempStr = Replace(TempStr, ":", "")
   Reg1(Val(Mid(TempStr, 1, 1))) = Reg1(Val(Mid(TempStr, 1, 1))) + 1
   Reg2(Val(Mid(TempStr, 2, 1))) = Reg2(Val(Mid(TempStr, 2, 1))) + 1
   Reg3(Val(Mid(TempStr, 3, 1))) = Reg2(Val(Mid(TempStr, 3, 1))) + 1
   k = k + 1
   If k = 10 Then
    k = 0
    For j = 0 To 6
     TempStr1 = TempStr1 + Str(Reg1(j))
     TempStr2 = TempStr2 + Str(Reg2(j))
     TempStr3 = TempStr3 + Str(Reg3(j))
     Reg1A(j) = Reg1A(j) + Reg1(j)
     Reg2A(j) = Reg2A(j) + Reg2(j)
     Reg3A(j) = Reg3A(j) + Reg3(j)
     Reg1(j) = 0
     Reg2(j) = 0
     Reg3(j) = 0
    Next j
       List4(1).AddItem TempStr1 + "|" + TempStr2 + "|" + TempStr3
    TempStr1 = ""
    TempStr2 = ""
    TempStr3 = ""
   End If
   

    
   
  Next i
  
  'For the Lost numbers
  If k > 0 Then
  For j = 0 To 6
     TempStr1 = TempStr1 + Str(Reg1(j))
     TempStr2 = TempStr2 + Str(Reg2(j))
     TempStr3 = TempStr3 + Str(Reg3(j))
     Reg1A(j) = Reg1A(j) + Reg1(j)
     Reg2A(j) = Reg2A(j) + Reg2(j)
     Reg3A(j) = Reg3A(j) + Reg3(j)
     Reg1(j) = 0
     Reg2(j) = 0
     Reg3(j) = 0
    Next j
       List4(1).AddItem TempStr1 + "|" + TempStr2 + "|" + TempStr3
  End If
  
  
   List4(1).AddItem "--------------|---------------------"
  List4(1).AddItem "01-11区间分布 |12-21区间分布|22-33区间分布|"
  List4(1).ListIndex = List4(1).ListCount - 1
   
End Sub

Private Sub ShowODDEVEN()
Dim i, j, k As Integer
Dim TempStr, TempStr1, TempStr2 As String
Dim ODDN(6), BIGN(6) As Integer
Dim OddAN(6), BigAN(6) As Integer
 For i = (List2(0).ListCount Mod 10) + 1 To List2(0).ListCount - 1  'just Lost 9 Nums For the last
   
    TempStr = Right(List2(0).List(i), 10)
    ODDN(Val(Mid(TempStr, 1, 1))) = ODDN(Val(Mid(TempStr, 1, 1))) + 1
    BIGN(Val(Mid(TempStr, 7, 1))) = BIGN(Val(Mid(TempStr, 7, 1))) + 1
   k = k + 1
  If k = 10 Then
   k = 0
  
  For j = 0 To 6
   TempStr1 = TempStr1 + Str(ODDN(j))
   TempStr2 = TempStr2 + Str(BIGN(j))
    OddAN(j) = OddAN(j) + ODDN(j)
    BigAN(j) = BigAN(j) + BIGN(j)
    ODDN(j) = 0
    BIGN(j) = 0
  Next j
   List2(1).AddItem TempStr1 + "|" + TempStr2
    TempStr1 = ""
   TempStr2 = ""
   End If
 
 Next i
 'For the LostNumbers
 If k > 0 Then
  For j = 0 To 6
   TempStr1 = TempStr1 + Str(ODDN(j))
   TempStr2 = TempStr2 + Str(BIGN(j))
    OddAN(j) = OddAN(j) + ODDN(j)
    BigAN(j) = BigAN(j) + BIGN(j)
    ODDN(j) = 0
    BIGN(j) = 0
  Next j
   List2(1).AddItem TempStr1 + "|" + TempStr2
 End If
 
 
 
 List2(1).AddItem "--------------|----------------"
List2(1).AddItem " 0 1 2 3 4 5 6| 0 1 2 3 4 5 6 "
List2(1).ListIndex = List2(1).ListCount - 1

For i = 0 To 6
 TempStr1 = TempStr1 + Str(OddAN(i))
 TempStr2 = TempStr2 + Str(BigAN(i))
Next i
 List2(2).AddItem TempStr1
 List2(3).AddItem TempStr2

End Sub









Private Function ShowSum(Sum As Integer, SumStr As String)
''compare with 102
Dim i As Integer
If Sum < 102 Then
    For i = 40 To Sum Step 2
        SumStr = SumStr + " "
    Next i
    For i = Sum To 101 Step 2
        SumStr = SumStr + "#"
    Next i
    SumStr = SumStr + "|"
End If
If Sum > 102 Then
    For i = 40 To 101 Step 2
        SumStr = SumStr + " "
    Next i
    SumStr = SumStr + "|"
    For i = 103 To Sum Step 2
        SumStr = SumStr + "@"
    Next i
 End If


End Function

Private Sub Check1_Click()
 If Check1.Value = 0 Then
   Coo5 = False
 ElseIf Check1.Value = 1 Then
    Coo5 = True
End If
 
 
 Call ShowList24


End Sub


Private Sub ShowList24()
Dim i, j, k As Integer
Dim LSUM5, SSUM5, EVEN5, ODD5 As Integer
Dim LSUM10, SSUM10, EVEN10, ODD10 As Integer
Dim Sum1, Sum2, Sum3 As Integer
Dim Sum110, Sum210, Sum310 As Integer
Dim ToNum As Integer
If Coo5 = True Then
    ToNum = 99
Else
    ToNum = 98
End If
List21.Clear
List22.Clear
List41.Clear
List42.Clear
List6(0).Clear
List6(1).Clear




Dim TempStr As String
For i = List2(0).ListCount - 1 - ToNum To List2(0).ListCount - 1
    TempStr = List2(0).List(i)
 LSUM5 = Val(Mid(TempStr, 8, 1)) + LSUM5
SSUM5 = Val(Mid(TempStr, 11, 1)) + SSUM5
EVEN5 = Val(Mid(TempStr, 14, 1)) + EVEN5
ODD5 = Val(Mid(TempStr, 17, 1)) + ODD5
 
    TempStr = List4(0).List(i)
   Sum1 = Val(Mid(TempStr, 2, 1)) + Sum1
   Sum2 = Val(Mid(TempStr, 5, 1)) + Sum2
   Sum3 = Val(Mid(TempStr, 8, 1)) + Sum3
 
 j = j + 1
 k = k + 1
 'If (i Mod 5) = 3 Then
 If j = 5 Then
 List21.AddItem Format(Str(LSUM5), "00") + "   :   " + Format(Str(SSUM5), "00")
 List22.AddItem Format(Str(EVEN5), "00") + "   :   " + Format(Str(ODD5), "00")
 LSUM10 = LSUM10 + LSUM5
 SSUM10 = SSUM10 + SSUM5
 EVEN10 = EVEN10 + EVEN5
 ODD10 = ODD10 + ODD5
 
 
 
 
 LSUM5 = 0
 SSUM5 = 0
 EVEN5 = 0
 ODD5 = 0
     List41.AddItem Format(Str(Sum1), "00") + " | " + Format(Str(Sum2), "00") + " | " + Format(Str(Sum3), "00")
    Sum110 = Sum110 + Sum1
    Sum210 = Sum210 + Sum2
    Sum310 = Sum310 + Sum3
    Sum1 = 0
    Sum2 = 0
    Sum3 = 0
    j = 0
 
 End If
 
' If (i Mod 10) = 9 Then '''l总数为0-399
 If k = 10 Then
     
     List6(0).AddItem Format(Str(LSUM10), "00") + "   :   " + Format(Str(SSUM10), "00")
     List6(1).AddItem Format(Str(EVEN10), "00") + "   :   " + Format(Str(ODD10), "00")
     LSUM10 = 0
    SSUM10 = 0
    EVEN10 = 0
    ODD10 = 0
    List42.AddItem Format(Str(Sum110), "00") + " | " + Format(Str(Sum210), "00") + " | " + Format(Str(Sum310), "00")
   Sum110 = 0
   Sum210 = 0
   Sum310 = 0
   
   k = 0
    
 End If
    
Next i

'While (Not Coo5)
If Not Coo5 Then
List21.AddItem Format(Str(LSUM5), "00") + "   :   " + Format(Str(SSUM5), "00")
  
 List22.AddItem Format(Str(EVEN5), "00") + "   :   " + Format(Str(ODD5), "00")
 
 LSUM10 = LSUM10 + LSUM5
 SSUM10 = SSUM10 + SSUM5
 EVEN10 = EVEN10 + EVEN5
 ODD10 = ODD10 + ODD5
   List6(0).AddItem Format(Str(LSUM10), "00") + "   :   " + Format(Str(SSUM10), "00")
     List6(1).AddItem Format(Str(EVEN10), "00") + "   :   " + Format(Str(ODD10), "00")
      List41.AddItem Format(Str(Sum1), "00") + " | " + Format(Str(Sum2), "00") + " | " + Format(Str(Sum3), "00")
    
    
    Sum110 = Sum110 + Sum1
    Sum210 = Sum210 + Sum2
    Sum310 = Sum310 + Sum3
     List42.AddItem Format(Str(Sum110), "00") + " | " + Format(Str(Sum210), "00") + " | " + Format(Str(Sum310), "00")
      
'Wend
End If
End Sub



Private Sub Form_Load()
Dim i, j As Integer
Dim TempStr As String
Dim SumStr As String
Dim TempSum As Integer      ''the sum good as 102
Dim Sumtimes(21 To 183) As Integer
Dim NumTimes(21 To 183) As Integer
Dim AreaTimes(1, 1000) As Integer
Dim EndTimes(9, 1000) As Integer
Dim EndNum As Integer
Dim EndStr As String
Dim AreaNum1, AreaNum0 As Integer
Dim OddNum As Integer
Dim BigNum As Integer
Dim TData(7) As String
For i = 0 To Form2.AllData.ListCount - 1
    List1.AddItem Form2.AllData.List(i)
    TempStr = Form2.AllData.List(i)
    Call GetNum(TempStr, TData())
    For j = 1 To 6
        TempSum = TempSum + Val(TData(j))
        If Val(TData(j)) Mod 2 Then
            OddNum = OddNum + 1
         End If
         If Val(TData(j)) > 16 Then
            BigNum = BigNum + 1
        End If
         If Val(TData(j)) < 12 Then
            AreaTimes(0, i) = AreaTimes(0, i) + 1
        End If
        If Val(TData(j)) > 22 Then
            AreaTimes(1, i) = AreaTimes(1, i) + 1
        End If
        EndTimes((Val(TData(j)) Mod 10), i) = EndTimes((Val(TData(j)) Mod 10), i) + 1
    
    Next j
    OddTimes(i) = OddNum
    BigTimes(i) = BigNum
    List2(0).AddItem "O:E = " + Str(OddNum) + ":" + Str(6 - OddNum) + "|" + Str(BigNum) + ":" + Str(6 - BigNum)
    Call ShowSum(TempSum, SumStr)
    Sumtimes(TempSum) = Sumtimes(TempSum) + 1
    SumtimesBack(TempSum) = SumtimesBack(TempSum) + 1
    List3.AddItem Format(Str(TempSum), "000") + " " + Format(Str(TempSum - 102), "00") + SumStr
    List4(0).AddItem Str(AreaTimes(0, i)) + ":" + Str(6 - AreaTimes(0, i) - AreaTimes(1, i)) + ":" + Str(AreaTimes(1, i))
    
    TempSum = 0
    OddNum = 0
    BigNum = 0
    SumStr = ""
Next i
Call DaPaixu(Sumtimes(), NumTimes())    ''leaf
For i = 21 To 60
    Label1.Caption = Label1.Caption + Format(Str(NumTimes(i)), "000") + "-" + Format(Str(Sumtimes(i)), "00") + " "
Next i
For i = 1 To 10
    OddNum = OddNum + OddTimes(List1.ListCount - i)
    BigNum = BigNum + BigTimes(List1.ListCount - i)
    If i > 4 Then
        Label3.Caption = Label3.Caption + Str(i) + ":" + Str(OddNum) + "-" + Str(6 * i - OddNum)
        Label4.Caption = Label4.Caption + Str(i) + ":" + Str(BigNum) + "-" + Str(6 * i - BigNum)
    End If
Next i
For i = 1 To 8
    AreaNum0 = AreaNum0 + AreaTimes(0, List1.ListCount - i)
    AreaNum1 = AreaNum1 + AreaTimes(1, List1.ListCount - i)
    If i > 4 Then
    Label5.Caption = Label5.Caption + Str(i) + "-" + Str(AreaNum0) + ":" + Str(6 * i - AreaNum0 - AreaNum1) + ":" + Str(AreaNum1) + " "
    End If
Next i
List1.ListIndex = List1.ListCount - 1
For j = 0 To 9
For i = 1 To 8
    EndNum = EndNum + EndTimes(j, List1.ListCount - i)
    If i > 4 Then
       EndStr = EndStr + Str(EndNum)
    End If
Next i
    
    List5.AddItem Str(j) + ": " + EndStr
    EndNum = 0
    EndStr = ""
Next j

' Coo5 = True
Check1.Value = 1

Call ShowODDEVEN
Call ShowReg

End Sub

Private Sub List1_Click()
List2(0).ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex

End Sub



Private Sub List2_Click(Index As Integer)
List1.ListIndex = List2(0).ListIndex
List3.ListIndex = List2(0).ListIndex
End Sub

Private Sub List21_Click()
'List6(0).ListIndex = (List21.ListIndex - 0.5) / 2
End Sub

Private Sub List22_Click()
'List6(1).ListIndex = List22.ListIndex / 2
End Sub

Private Sub List3_Click()
List2(0).ListIndex = List3.ListIndex
List1.ListIndex = List3.ListIndex
List4(0).ListIndex = List3.ListIndex
End Sub



Private Sub Text1_Change()
If Len(Text1.Text) > 3 Then
    Text1.Text = ""
End If
If Val(Text1.Text) > 20 And Val(Text1.Text) < 184 Then
    Label2.Caption = SumtimesBack(Val(Text1.Text))
Else
    Label2.Caption = "输入数据超出范围21-183"
End If
End Sub
