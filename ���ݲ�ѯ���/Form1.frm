VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15630
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   15630
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   8400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   30
      Min             =   21
      Max             =   2000
      SelStart        =   100
      Value           =   100
   End
   Begin VB.ListBox List3 
      Height          =   7260
      Left            =   10320
      TabIndex        =   10
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   7335
      Begin VB.OptionButton Option1 
         Caption         =   "和值统计"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "大小奇偶"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "尾号模式"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "区间模式"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "查号模式"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      Height          =   7620
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "万能按钮"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1935
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
      Height          =   9420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label1"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   8880
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim data(2000) As String
Dim Regs(2000) As String
Dim RegList(2000) As String
Dim RegSlider(2000) As String
Dim TailList(2000) As String
Dim L3Str(2000)  As String
Dim setList(2000) As Integer
Dim OEBSList(500) As Integer


Dim SliderMax As Integer

Dim SWModel As Integer
Dim FindNums As Integer
Dim Total3 As Integer


Private Sub Command1_Click()
'Form3.Show
If SWModel = 0 Then
    Call TestNum

End If
If SWModel = 1 Then
   Call ShowRegNotAppear

End If

End Sub
Private Sub TestNum()
 Dim i As Integer
 Dim a1, a2, a3, a4, a5, a6 As Integer
 For a1 = 1 To 2
     For a2 = 1 To 2
        For a3 = 1 To 2
            For a4 = 1 To 2
               ' For a5 = 1 To 2
                   If a1 + a2 + a3 + a4 + a5 = 6 Then
                    List2.AddItem Str(a1) + Str(a2) + Str(a3) + Str(a4) + Str(a5)
                    End If
               ' Next a5
            Next a4
        Next a3
    Next a2
Next a1
            
 
 
 

End Sub

Private Sub ShowRegNotAppear()
Dim R1, R2, R3, R4, R5, R6 As Integer
Dim i As Integer
Dim tempstr As String
List3.Clear
For R1 = 0 To 2
    For R2 = 0 To 2
        For R3 = 0 To 2
            For R4 = 0 To 2
                For R5 = 0 To 2
                    For R6 = 0 To 2
                        If (R1 + R2 + R3 + R4 + R5 + R6) = 6 Then
                            tempstr = Str(R1) + Str(R2) + Str(R3) + Str(R4) + Str(R5) + Str(R6)
                            Label1.BackColor = vbRed
                            tempstr = Replace(tempstr, " ", "")
                           For i = 1 To TotalNum
                            If tempstr = Mid(Regs(i), 1, Len(tempstr)) Then
                              ' GoTo findone
                               Exit For
                            End If
                            Label1.Caption = tempstr + "   " + Str(i)
                           Next i

                           If i = TotalNum + 1 Then
                                List3.AddItem tempstr
                           End If
                          Label1.BackColor = vbGreen
                        End If
                      Next R6
                  Next R5
              Next R4
          Next R3
      Next R2
  Next R1
                      
'Label1.BackColor = vbRed


End Sub


Private Sub Form_Load()
Dim i As Integer
SWModel = 0
Open App.Path & "/data.txt" For Input As #1
While Not EOF(1)

Line Input #1, data(i)
List1.AddItem data(i)
data(i) = Replace(data(i), " ", "")
BlueNum(i) = Right(data(i), 2)
TotalNum = TotalNum + 1
i = i + 1
Wend
Close #1
FindNums = 0
List1.ListIndex = List1.ListCount - 1
End Sub


Private Sub List1_Click()
If SWModel = 3 Then
 List2.ListIndex = List1.ListIndex
End If
End Sub

Private Sub List2_Click()
If SWModel = 0 Then
    List1.ListIndex = setList(List2.ListIndex)
End If
If SWModel = 3 Then


End If


End Sub

Private Sub Option1_Click(Index As Integer)
SWModel = Index
Text1.Text = ""
Form1.Caption = "软件模式" + Str(Index) + " " + Option1(Index).Caption

Select Case Index
Case 0
    List2.Clear
Case 1
    Call Model1
Case 2
    Call ShowTail
    Call ShowTailList3
Case 3
    Call ShowOEBS
Case 4
    Call ShowSum
    Form3.Show
End Select


End Sub
Private Sub Model0()
Dim i, j As Integer
Dim tempstr As String
Dim ss As String
List2.Clear
tempstr = Replace(Text1.Text, " ", "")
If Len(tempstr) >= 12 Then
    Text1.Text = ""
For i = 1 To Len(tempstr) Step 2
    Text1.Text = Text1.Text + " " + Mid(tempstr, i, 2)
Next i
'    Text1.Text = Text1.Text + "+" + Right(tempstr, 2)
For i = 0 To TotalNum - 1
   List1.ListIndex = i
    If (Mid(tempstr, 1, 12) = Mid(data(i), 6, 12)) Then
        List2.AddItem List1.List(i)
        setList(j) = i
    End If
Next i

End If
List1.ListIndex = setList(0)
End Sub

Private Sub Model1()
Dim i, j As Integer
Dim k  As Integer
Dim tempstr As String
Dim iPre As Integer
List2.Clear
For i = List1.ListCount - 1 To 0 Step -1
    tempstr = Right(Replace(List1.List(i), " ", ""), 14)
    Regs(TotalNum - i) = CheckR6(tempstr)
 Next i
    
For i = 0 To TotalNum
    iPre = i
  If (Len(Regs(i)) > 0) Then
    RegList(i) = Regs(i) + "  " + Str(i)
  End If
  For j = i + 1 To TotalNum
   
    If (Regs(i) = Regs(j)) And (Len(Regs(i)) > 0) Then
        RegList(i) = RegList(i) + "   " + Str(j - iPre)
        Regs(j) = ""
        iPre = j
    End If
  Next j
Next i

For j = 1 To i
 If Len(RegList(j)) > 7 Then
 List2.AddItem RegList(j)
 RegSlider(k) = RegList(j)
 k = k + 1
 End If
Next j
 Slider1.Max = k
 SliderMax = k
 Slider1.Value = k / 2

End Sub
Public Function CheckR6(InStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    For i = 0 To 6
        CheckR6 = CheckR6 + Format(Str(Reg(i)), "0")
    Next i
End Function

Private Sub Slider1_Change()
Dim i, ISet As Integer



If (SWModel = 1) Then
List3.Clear
ISet = Slider1.Value

If (ISet > 20 And ISet < SliderMax - 20) Then
For i = ISet - 15 To ISet + 15
    List3.AddItem RegSlider(i)
Next i
End If
End If

If (SWModel = 2) Then



End If

End Sub

Private Sub Text1_Change()

If SWModel = 0 Then
 If Len(Text1.Text) > 12 Then
    Text1.Text = ""
 End If
If (Len(Text1.Text) Mod 2) = 0 And Len(Text1.Text) > 1 Then
    CheckSameNum (Text1.Text)
End If
End If
If SWModel = 1 Then
    If Len(Text1.Text) > 7 Then
     Text1.Text = ""
     End If
     If Text1.Text = "n" Then
        Call ShowNoneReg
     Else
        Call CheckRegList(Text1.Text)
     End If
    
End If

If SWModel = 2 Then
    If Len(Text1.Text) > 6 Then
        Text1.Text = ""
     End If
      If Len(Text1.Text) > 2 Then
        Call CheckTailList(Text1.Text)
     End If
    If Text1.Text = "" Then
     Call ShowTailList3
    End If
    
End If



If SWModel = 3 Then
    Call OEBSList3

End If



Label1.Caption = "总计 :   " + Str(FindNums)

End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Private Sub CheckRegList(iStr As String)
Dim i, j As Integer
List2.Clear
For i = 0 To TotalNum
 If iStr = Mid(Regs(i), 1, Len(iStr)) Then
  List2.AddItem RegList(i)
  j = j + 1
 End If
Next i
FindNums = j
End Sub

Private Sub CheckSameNum(iStr As String)
Dim i, j, k, X, N As Integer
Dim nflag As Integer

N = (Len(iStr) + 1) \ 2
       
List2.Clear
For i = 0 To TotalNum
  nflag = 0
 For j = 1 To N
   
   For k = 0 To 5

    If Mid(data(i), 6 + k * 2, 2) = Mid(iStr, 2 * j - 1, 2) Then
       'List2.AddItem List1.List(i)
       nflag = nflag + 1
       Exit For
   End If
   Next k
  Next j
If nflag = N Then
    List2.AddItem List1.List(i)
    setList(X) = i
    X = X + 1
 End If

Next i
FindNums = X
End Sub


Private Function CheckTail(iStr As String) As String
    Dim i As Integer
    Dim TailFlag(10) As Integer
    Dim tempstr As String
    For i = 0 To 5
      TailFlag(Val(Mid(iStr, 2 * i + 1, 2)) Mod 10) = 1 + TailFlag(Val(Mid(iStr, 2 * i + 1, 2)) Mod 10)
    Next i
   For i = 0 To 9
    If TailFlag(i) > 0 Then
        CheckTail = CheckTail + " " + Str(i)
    End If
   Next i
    
    
    
End Function
Private Sub ShowTail()
Dim i As Integer
Dim tempstr As String
Dim Tails(2000) As String
List2.Clear
For i = List1.ListCount - 1 To 0 Step -1
    tempstr = Right(Replace(List1.List(i), " ", ""), 14)
    Tails(TotalNum - i) = CheckTail(tempstr)
    List2.AddItem CheckTail(tempstr)
 Next i


End Sub
Private Sub CheckTailList(iStr As String)
   Dim i, Total As Integer
   Dim tempstr, SaveStr As String

   List3.Clear
   For i = 0 To Total3 - 1
    tempstr = L3Str(i)
    tempstr = Mid(tempstr, 1, InStr(tempstr, " ") - 1)
    If InStr(tempstr, iStr) Then
     List3.AddItem L3Str(i)
     End If
  Next i
        



End Sub
Private Sub ShowTailList3()
Dim i, j As Integer
Dim iPre As Integer
Dim tempstr As String
List3.Clear
For i = 0 To List2.ListCount - 1
     TailList(i) = Replace(List2.List(i), " ", "")
Next i
For i = 0 To List2.ListCount - 2
    iPre = i
    tempstr = Format(TailList(i), "@@@@@@")
    If Len(tempstr) > 1 Then
    tempstr = tempstr + Format(Str(i), "@@@@")
    For j = i + 1 To List2.ListCount - 1
      If TailList(i) = TailList(j) Then
        tempstr = tempstr + Format(Str(j - iPre), "@@@@")
        TailList(j) = ""
        iPre = j
      End If
    Next j
    End If
    If Len(tempstr) > 3 Then
        List3.AddItem tempstr
    End If
    
Next i
Total3 = 0
   For i = 0 To List3.ListCount - 1
    L3Str(i) = LTrim(List3.List(i))
    Total3 = Total3 + 1
   Next i

End Sub

Private Sub ShowOEBS()
Dim i, j As Integer
Dim OEBSdata(3000) As String
Dim tempstr As String
Dim linestr As String
Label1.Caption = "小奇小偶 大奇大偶"
List2.Clear
For i = 0 To TotalNum - 1
    OEBSdata(i) = CheckOEBS(Mid(data(i), 6, 12))
'    List2.AddItem tempstr
Next i

For i = TotalNum - 1 To 0 Step -1
    tempstr = Replace(OEBSdata(i), " ", "")
   linestr = tempstr + Str(TotalNum - 1 - i)
    For j = i - 1 To 0 Step -1
        If tempstr = Replace(OEBSdata(j), " ", "") Then
            linestr = linestr + Str(i - j)
            Exit For
        End If
    Next j
        List2.AddItem linestr
Next i



End Sub
Private Function CheckOEBS(iStr As String) As String
Dim i As Integer
Dim OEBS(4) As Integer
Dim Temp As Integer
For i = 0 To 5
    Temp = Val(Mid(iStr, 2 * i + 1, 2))
  If Temp Mod 2 Then
    If Temp < 17 Then
        OEBS(0) = OEBS(0) + 1   'small odd
    Else
        OEBS(2) = OEBS(2) + 1   'big odd
    End If
   Else
      If Temp > 16 Then
        OEBS(3) = OEBS(3) + 1   'small even
    Else
        OEBS(1) = OEBS(1) + 1   'big  even
    End If
  
  End If
Next i
For i = 0 To 3
   CheckOEBS = CheckOEBS + " :  " + Str(OEBS(i))
    
Next i
End Function

Private Sub OEBSList3()
  Dim i As Integer
   Dim tempstr As String
  List3.Clear
    For i = 0 To List2.ListCount - 1
        tempstr = Replace(List2.List(i), " ", "")
        tempstr = Replace(tempstr, ":", "")
        If Text1.Text = Left(tempstr, Len(Text1.Text)) Then
            List3.AddItem List2.List(i)
        End If
   Next i
    

End Sub

Private Sub ShowSum()
    Dim i As Integer


    For i = 0 To TotalNum - 1
         SumNum(i) = Val(Mid(data(i), 6, 2)) + Val(Mid(data(i), 8, 2)) + Val(Mid(data(i), 10, 2)) + Val(Mid(data(i), 12, 2)) + Val(Mid(data(i), 14, 2)) + Val(Mid(data(i), 16, 2))
         SumReg(SumNum(i)) = SumReg(SumNum(i)) + 1
     '    List2.AddItem SumNum(i)
            
    Next i
    For i = 21 To 183
        List3.AddItem Format(Str(i), "000") + " |        " + Format(SumReg(i), "000")
    Next i
    
End Sub

Private Sub ShowNoneReg()
    

End Sub




