VERSION 5.00
Begin VB.Form LongTime 
   Caption         =   "长期彩票趋势预测长期"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   13485
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "伴侣数字概率"
      Height          =   7695
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   9855
      Begin VB.ListBox List4 
         Height          =   6180
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.ListBox List3 
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   9375
      End
      Begin VB.ListBox List1 
         Height          =   6180
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   9375
      End
      Begin VB.Label Label2 
         Caption         =   "次数:"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   7200
         Width           =   9495
      End
      Begin VB.Label Label2 
         Caption         =   "数字:"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   6720
         Width           =   9495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "重复中将模式表"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.ListBox List2 
         Height          =   6720
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "比列越小,重复中将概率越大"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   7080
         Width           =   2535
      End
   End
End
Attribute VB_Name = "LongTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalNum As Integer
Dim SaveData(1 To 1000, 1 To 33) As Integer
Dim FollNum(1 To 33, 1 To 33) As Integer
'Dim RedNum(1 To 1000, 1 To 16) As Integer

Private Sub ShowRepeat()
Dim i, j As Integer
Dim Ditto(1 To 33, 1 To 7) As Integer
Dim bilvStr(1 To 33) As String
'For j = 1 To 33
  ' For i = 1 To 500
 '
Dim tempstr As String

Dim tdata(7) As String
List2.Clear


For i = 1 To TotalNum
   ' TempStr = Form2.AllData.List(Form2.AllData.ListCount - TotalNum - 1 + i)
 tempstr = Form2.AllData.List(Form2.AllData.ListCount - TotalNum - 1 + i)
'For i = 80 To 1 Step -1

    Call GetNum(tempstr, tdata())
    For j = 1 To 6      ''1 to 6
        SaveData(i, Val(tdata(j))) = 1
    Next j
    Next i
For j = 1 To 33
    For i = 1 To TotalNum - 4
        If SaveData(i, j) = 1 And SaveData(i + 1, j) = 1 And SaveData(i + 2, j) = 1 And SaveData(i + 3, j) = 1 And SaveData(i + 4, j) = 1 Then
            Ditto(j, 5) = Ditto(j, 5) + 1
             'Form2.MainList.ListIndex = i
        End If
        If SaveData(i, j) = 1 And SaveData(i + 1, j) = 1 And SaveData(i + 2, j) = 1 And SaveData(i + 3, j) = 1 Then
            Ditto(j, 4) = Ditto(j, 4) + 1
           
        End If
       
        If SaveData(i, j) = 1 And SaveData(i + 1, j) = 1 And SaveData(i + 2, j) = 1 Then
            Ditto(j, 3) = Ditto(j, 3) + 1
        End If
        
        If SaveData(i, j) = 1 And SaveData(i + 1, j) = 1 Then
            Ditto(j, 2) = Ditto(j, 2) + 1
        End If
        If SaveData(i, j) = 1 Then
            Ditto(j, 1) = Ditto(j, 1) + 1
        End If
    
    Next i
    For i = TotalNum - 3 To TotalNum
         If SaveData(i, j) = 1 Then
            Ditto(j, 1) = Ditto(j, 1) + 1
        End If
        If SaveData(i - 1, j) = 1 And SaveData(i, j) = 1 Then
            Ditto(j, 2) = Ditto(j, 2) + 1
        End If
        If SaveData(i - 2, j) = 1 And SaveData(i - 1, j) = 1 And SaveData(i, j) = 1 Then
            Ditto(j, 3) = Ditto(j, 3) + 1
        End If
        If SaveData(i - 3, j) = 1 And SaveData(i - 2, j) = 1 And SaveData(i - 1, j) = 1 And SaveData(i, j) = 1 Then
            Ditto(j, 4) = Ditto(j, 4) + 1
        End If
        
    Next i
    Ditto(j, 4) = Ditto(j, 4) - 2 * Ditto(j, 5)
    Ditto(j, 3) = Ditto(j, 3) - 2 * Ditto(j, 4) - 3 * Ditto(j, 5)
    Ditto(j, 2) = Ditto(j, 2) - 2 * Ditto(j, 3) - 3 * Ditto(j, 4) - 4 * Ditto(j, 5)
    Ditto(j, 1) = Ditto(j, 1) - 2 * Ditto(j, 2) - 3 * Ditto(j, 3) - 4 * Ditto(j, 4) - 5 * Ditto(j, 5)
    bilvStr(j) = bilvStr(j) + Str(Format(Ditto(j, 1) / (Ditto(j, 5) + Ditto(j, 4) + Ditto(j, 3) + Ditto(j, 2)), "0.00")) + " "

Next j
For j = 1 To 33
    List2.AddItem Format(Str(j), "00") + ":" + Format(Str(Ditto(j, 1)), "00") + " " + Format(Str(Ditto(j, 2)), "00") + " " + Format(Str(Ditto(j, 3)), "00") _
    + " " + Format(Str(Ditto(j, 4)), "00") + " " + Format(Str(Ditto(j, 5)), "00") + " " + bilvStr(j)

Next j

End Sub
Private Sub ShowFollow()

Dim i, j, k As Integer
Dim TempNum(1 To 6) As Integer
Dim tempstr As String
List3.AddItem "01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 "
For i = 1 To TotalNum
   k = 1
    For j = 1 To 33
        If SaveData(i, j) = 1 Then
          TempNum(k) = j
          k = k + 1
        End If
    Next j
    For k = 1 To 5
        FollNum(TempNum(k), TempNum(k + 1)) = FollNum(TempNum(k), TempNum(k + 1)) + 1
    Next k
    For k = 1 To 4
        FollNum(TempNum(k), TempNum(k + 2)) = FollNum(TempNum(k), TempNum(k + 2)) + 1
    Next k
    For k = 1 To 3
        FollNum(TempNum(k), TempNum(k + 3)) = FollNum(TempNum(k), TempNum(k + 3)) + 1
    Next k
    For k = 1 To 2
        FollNum(TempNum(k), TempNum(k + 4)) = FollNum(TempNum(k), TempNum(k + 4)) + 1
    Next k
        FollNum(TempNum(1), TempNum(6)) = FollNum(TempNum(1), TempNum(6)) + 1
Next i
For i = 1 To 33
    For j = 1 To 33
        FollNum(j, i) = FollNum(i, j)
        tempstr = tempstr + Format(FollNum(i, j), "00") + " "
    Next j
    List1.AddItem tempstr
    List4.AddItem Format(Str(i), "00")
    tempstr = ""

Next i

End Sub


Private Sub Form_Load()
TotalNum = 300
Call ShowRepeat
Call ShowFollow
Call ShowSel
End Sub
Private Sub ShowSel()
Dim tempstr As String
Dim i As Integer
tempstr = Right(Replace(Form2.yllist.List(0), " ", ""), 12)
For i = 1 To 6
 List2.Selected(Val(Mid(tempstr, 2 * i - 1, 2)) - 1) = True
Next i

End Sub

Private Sub List1_Click()
Dim TempFoll(1 To 33) As Integer
Dim TempNum(1 To 33) As Integer
Dim i As Integer
Label2(0).Caption = "数字:"
Label2(1).Caption = "次数:"
For i = 1 To 33
    TempFoll(i) = FollNum(List1.ListIndex + 1, i)
Next i
Call DaPaixu(TempFoll(), TempNum())
For i = 1 To 33
    Label2(1).Caption = Label2(1).Caption + Format(TempFoll(i), "00") + " "
    Label2(0).Caption = Label2(0).Caption + Format(TempNum(i), "00") + " "
Next i

End Sub
