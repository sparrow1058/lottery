VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "出现总次数区间统计"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form5"
   ScaleHeight     =   8505
   ScaleWidth      =   11100
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List8 
      BackColor       =   &H0080FF80&
      Height          =   2940
      Left            =   3360
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox List7 
      Height          =   1680
      Left            =   240
      TabIndex        =   9
      Top             =   6720
      Width           =   10575
   End
   Begin VB.ListBox List6 
      Height          =   2400
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   10575
   End
   Begin VB.Frame Frame1 
      Caption         =   "总计数据出现区间"
      Height          =   2055
      Left            =   5160
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
      Begin VB.ListBox List5 
         Height          =   1500
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "1R 2R 3R 4R 5R 6R 7R"
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "1R|2R|3R|4R|5R|6R|"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.ListBox List4 
      Height          =   1860
      Left            =   8040
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin VB.ListBox List3 
      BackColor       =   &H0080FF80&
      Height          =   2940
      Left            =   1680
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   2940
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   5520
      TabIndex        =   10
      Top             =   5640
      Width           =   5295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RegionN(1 To 33) As Integer
Private Sub ShowList4()
Dim List4Str(1 To 7) As String
Dim i As Integer
 For i = 1 To 33
  List4Str(RegionN(i)) = List4Str(RegionN(i)) + Format(Str(i), "00") + " "
 Next i
For i = 1 To 7
  List4.AddItem Format(Str(i), "00") + ": " + List4Str(i)
Next i


End Sub

Private Sub ShowList6()
Dim i, j, k As Integer
Dim TempDatat(33) As Integer
Dim TempStr1, TempStr2 As String
Dim Temp10(10) As String
 
 TempStr1 = Replace(List1.List(1), " ", "")
 For i = 1 To 33
  TempDatat(i) = Val(Mid(TempStr1, i * 3 - 2, 2))
 Next i

For i = List2.ListCount - 1 To 0 Step -1
 TempStr2 = Replace(List2.List(i), " ", "")
  For j = 1 To 6
  
  ' TempDatat(Val(Mid(TempStr2, j * 2 - 1, 2))) = TempDatat(Val(Mid(TempStr2, j * 2 - 2, 2))) - 1 'error
  TempDatat(Val(Mid(TempStr2, j * 2 - 1, 2))) = TempDatat(Val(Mid(TempStr2, j * 2 - 1, 2))) - 1
  
  
  Next j
 For k = 1 To 33
   Temp10(i) = Temp10(i) + Format(Str(TempDatat(k)), "00") + "|"
   If (k Mod 5) = 0 Then
     Temp10(i) = Temp10(i) + "  "
   End If
   
 Next k
 Next i
For i = 0 To 10
 List6.AddItem Temp10(i)
Next i

 
End Sub
Private Function BTOS(InputStr As String)
Dim Tempstr As String
Dim i, j As Integer
Dim MaxAppear, MaxNum As Integer
Dim TempData(1, 33) As Integer
For i = 1 To 33
 TempData(1, i) = Val(Mid(InputStr, i * 2 - 1, 2))
 TempData(0, i) = i
Next i
    For j = 1 To 32
     For i = 1 To 33 - j
       
        If TempData(1, i) < TempData(1, i + 1) Then
            MaxAppear = TempData(1, i + 1)
            MaxNum = TempData(0, i + 1)
            TempData(1, i + 1) = TempData(1, i)
            TempData(0, i + 1) = TempData(0, i)
            TempData(1, i) = MaxAppear
            TempData(0, i) = MaxNum
        End If
     Next i
     Next j
   InputStr = ""
For i = 1 To 33
 InputStr = InputStr + Format(Str(TempData(0, i)), "00")
Next i

InputStr = Replace(InputStr, " ", "")





End Function

Private Sub ShowList3New()
Dim Tempstr, Linestr As String
Dim List1Str2 As String
Dim RegionNew(33) As Integer
Dim L5Data(6, 7) As Integer
Dim i, j, k  As Integer
Dim DataSum As Integer

For k = 0 To List2.ListCount - 1
 Tempstr = Replace(List2.List(k), " ", "")
        List1Str2 = Replace(List6.List(k), " ", "")
       List1Str2 = Replace(List1Str2, "|", "")
       Call BTOS(List1Str2)
For i = 1 To 6
      For j = 1 To 5
          'RegionN(i, j) = Val(Mid(List1.List(2), (i - 1) * 15 + 3 * j - 2, 2))

               
       RegionNew(Val(Mid(List1Str2, (i - 1) * 10 + 2 * j - 1, 2))) = i
      
      Next j
Next i
   RegionNew(Val(Mid(List1Str2, 61, 2))) = 7
   RegionNew(Val(Mid(List1Str2, 63, 2))) = 7
   RegionNew(Val(Mid(List1Str2, 65, 2))) = 7
    RegionNew(28) = 1

For i = 1 To 6
   Linestr = Linestr + Format(RegionNew(Mid(Tempstr, i * 2 - 1, 2)), "0") + " | "
  '  RegionNew(Mid(TempStr, i * 2 - 1, 2)) = 1
 Next i
 List3.AddItem Linestr
  Linestr = ""

Next k

For i = 0 To List3.ListCount - 1
 Tempstr = Replace(List3.List(i), " ", "")
 Tempstr = Replace(Tempstr, "|", "")
 
  For j = 1 To 6
    L5Data(j, Mid(Tempstr, j, 1)) = L5Data(j, Mid(Tempstr, j, 1)) + 1
  Next j
Next i
For i = 1 To 7
 Tempstr = ""
 For j = 1 To 6
  DataSum = DataSum + L5Data(j, i)
  Tempstr = Tempstr + Str(L5Data(j, i)) + "|"
 Next j
 List5.AddItem Tempstr + "  " + Format(Str(DataSum), "00")
 DataSum = 0
Next i
 





End Sub



Private Sub ShowList3()

Dim Data(6) As Integer
Dim DataA(6, 7) As Integer
Dim i, j, k As Integer
Dim RegVal, LineVal As Integer
Dim Tempstr, Linestr As String
Dim List1Str2 As String
List1Str2 = Replace(List1.List(2), " ", "")

 For i = 1 To 6
      For j = 1 To 5
          'RegionN(i, j) = Val(Mid(List1.List(2), (i - 1) * 15 + 3 * j - 2, 2))
       RegionN(Val(Mid(List1Str2, (i - 1) * 15 + 3 * j - 2, 2))) = i
    
      Next j
Next i
   RegionN(Val(Mid(List1Str2, 91, 2))) = 7
   RegionN(Val(Mid(List1Str2, 94, 2))) = 7
   RegionN(Val(Mid(List1Str2, 97, 2))) = 7


For i = 0 To List2.ListCount - 1
 Tempstr = Replace(List2.List(i), " ", "")
  For j = 1 To 6
    Data(j) = Mid(Tempstr, 2 * j - 1, 2)
    DataA(j, RegionN(Data(j))) = DataA(j, RegionN(Data(j))) + 1
    Linestr = Linestr + Str(RegionN(Data(j))) + " | "
    LineVal = LineVal + RegionN(Data(j))
  Next j
  'List3.AddItem LineStr + Str(LineVal)         'remove  as test
  Linestr = ""
  LineVal = 0
Next i

  For j = 1 To 7
   Linestr = ""
   
   For i = 1 To 6
     RegVal = RegVal + DataA(i, j)
    
     Linestr = Linestr + Format(Str(DataA(i, j)), "00") + "|"
   Next i
   'List5.AddItem LineStr + Str(RegVal)  'remove
   RegVal = 0
   Next j
  

  
  

End Sub

Private Sub Form_Load()
Dim i, j As Integer
Dim Tempstr, TempStr2 As String
For i = 0 To 3
Tempstr = Replace(Form2.labelList.List(i), " ", "")

 For j = 0 To 6
    TempStr2 = TempStr2 + Mid(Tempstr, j * 15 + 1, 15) + "  "
 Next j
 List1.AddItem TempStr2
 TempStr2 = ""
 
Next i
For i = 0 To 10
 List2.AddItem Trim(Right(Form2.AllData.List(Form2.AllData.ListCount - 11 + i), 21))
Next i

Call ShowList3
Call ShowList4
Call ShowList6
Call ShowList3New
Call ShowList7
Call ShowList8
'Call ChangeList1
End Sub
Private Function GetSumRegion(Tempstr As String, SumReg() As Integer)
Dim i, j As Integer
Dim TMax, TMin As Integer
Dim TRegn As Integer
Dim Tsum(1 To 33) As String
For i = 1 To 33
 'TSum(i, 0) = i
 Tsum(i) = Val(Mid(Tempstr, i * 2 - 1, 2))
Next i
TMin = Tsum(1)
TMax = Tsum(1)
For i = 2 To 33
 If TMax < Tsum(i) Then
    TMax = Tsum(i)       '74
  End If
  If TMin > Tsum(i) Then
    TMin = Tsum(i)       '40
  End If
Next i

  TRegn = (TMax - TMin) / 6
For j = 0 To 6
 
 For i = 1 To 33
    If Tsum(i) >= (TMin + TRegn * j) And Tsum(i) < (TMin + TRegn * (j + 1)) Then
        SumReg(i) = j
    End If
Next i
Next j

End Function
Private Sub ShowList8()
Dim i, j As Integer
Dim Tempstr As String
Dim Linestr As String
Dim OutSum(1 To 33) As Integer
For i = 0 To List6.ListCount - 2
Tempstr = List6.List(i)
Tempstr = Replace(Tempstr, " ", "")
Tempstr = Replace(Tempstr, "|", "")
Call GetSumRegion(Tempstr, OutSum())
 Tempstr = Replace(List2.List(i + 1), " ", "")
For j = 1 To 7
    Linestr = Linestr + Str(OutSum(Val(Mid(Tempstr, 2 * j - 1, 2)))) + " "
Next j
   List8.AddItem Linestr
    Linestr = ""
Next i
    
End Sub

Private Sub ShowList7()
Dim i As Integer
Dim Tempstr As String
Dim OutSum(1 To 33) As Integer
Tempstr = List1.List(1)
Tempstr = Replace(Tempstr, " ", "")
Tempstr = Replace(Tempstr, "|", "")
Call GetSumRegion(Tempstr, OutSum())
For i = 1 To 33
    Label3.Caption = Label3.Caption + Str(OutSum(i)) + " "
Next i
Call ShowList7Old
End Sub

Private Sub ShowList7Old()
Dim i, j As Integer
Dim TMax, TMin As Integer
Dim TRegn As Integer
Dim Tsum(33) As Integer
Dim Linestr(6), NumStr(6) As String
Dim Tempstr As String
Tempstr = List1.List(1)
Tempstr = Replace(Tempstr, " ", "")
Tempstr = Replace(Tempstr, "|", "")
For i = 1 To 33
 'TSum(i, 0) = i
 Tsum(i) = Val(Mid(Tempstr, i * 2 - 1, 2))
Next i
TMin = Tsum(1)
TMax = Tsum(1)
For i = 2 To 33
 If TMax < Tsum(i) Then
    TMax = Tsum(i)       '74
  End If
  If TMin > Tsum(i) Then
    TMin = Tsum(i)       '40
  End If
Next i

  TRegn = (TMax - TMin) / 6
For j = 0 To 6
 
 For i = 1 To 33
      If Tsum(i) >= (TMin + TRegn * j) And Tsum(i) < (TMin + TRegn * (j + 1)) Then
        Linestr(j) = Linestr(j) + " " + Format(Str(i), "00") + "-" + Format(Str(Tsum(i)), "00") + "|"
        NumStr(j) = NumStr(j) + Format(Str(i), "00") + " "
      End If
     Next i
  Next j
 
   For i = 0 To 6
   List7.AddItem Linestr(i) + "   " + NumStr(i)
  Next i
        

End Sub

