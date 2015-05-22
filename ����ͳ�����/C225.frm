VERSION 5.00
Begin VB.Form C225 
   Caption         =   "22 C 5"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   13770
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox LoseList 
      Height          =   2220
      Left            =   2400
      TabIndex        =   15
      Top             =   5520
      Width           =   8655
   End
   Begin VB.ListBox List6 
      Height          =   2220
      Left            =   11040
      TabIndex        =   14
      Top             =   4200
      Width           =   2415
   End
   Begin VB.ListBox LessList 
      Height          =   1320
      Left            =   2400
      TabIndex        =   12
      Top             =   4200
      Width           =   8655
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÒÅÂ©Æ«²î"
      Height          =   1695
      Left            =   2640
      TabIndex        =   5
      Top             =   8160
      Width           =   11055
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   2
         Left            =   5400
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   2
         Left            =   5400
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   240
      Left            =   2400
      TabIndex        =   4
      Top             =   3960
      Width           =   8655
   End
   Begin VB.ListBox List4 
      Height          =   2040
      Left            =   11040
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   11040
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   4020
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.ListBox List1 
      Height          =   6360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "C225"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalNum As Integer
Private Function ShowData()
Dim i, j As Integer
Dim tempstr As String
Dim LineStr As String
Dim Tempdata As Integer
Dim AllData(1 To 500, 1 To 22) As Integer
Dim LoseNum(1 To 500, 1 To 22) As Integer
Dim Less10(10) As Integer
For i = 1 To TotalNum
  For j = 1 To 5
    tempstr = List1.List(i - 1)
    Tempdata = Val(Mid(tempstr, 3 * j + 4, 2))
    AllData(i, Tempdata) = 1
  Next j
Next i
For j = 1 To 22
    If AllData(1, j) = 1 Then
        LoseNum(1, j) = 0
    Else
        LoseNum(1, j) = 1
    End If
    For i = 2 To TotalNum
        If AllData(i, j) = 0 Then
            LoseNum(i, j) = LoseNum(i - 1, j) + 1
        Else
            LoseNum(i, j) = 0
        End If
    Next i
Next j
For i = 1 To TotalNum
    tempstr = ""
    For j = 1 To 22
        tempstr = tempstr + Format(Str(LoseNum(i, j)), "00") + "| "
     Next j
     List2.AddItem Format(Str(i), "000") + ":" + tempstr
Next i

For i = 0 To 10
 For j = 1 To 22
    If LoseNum(TotalNum - i, j) = 0 Then
    If LoseNum(TotalNum - i - 1, j) < 10 Then
        Less10(LoseNum(TotalNum - i - 1, j)) = Less10(LoseNum(TotalNum - i - 1, j)) + 1
    Else
        Less10(10) = Less10(10) + 1
    End If
    End If
 Next j
    If i >= 4 Then
        
    For j = 0 To 10
        LineStr = LineStr + Format(Str(j), "00") + ":" + Format(Str(Less10(j)), "00") + " "
    Next j
    LessList.AddItem "×Ü¼Æ" + Format(Str(i + 1), "00") + "ÆÚÒÅÂ©Êý¾Ý£º" + LineStr
    LineStr = ""
    End If

Next i


End Function
Private Function ShowOE()
Dim tempstr As String
Dim i, j As Integer
Dim TempSum(1 To 10) As Integer
Dim TempOdd(1 To 10) As Integer
Dim OddSum As Integer
List6.Clear
For i = 1 To 10
    tempstr = List1.List(TotalNum - i)
    For j = 1 To 5
        TempSum(i) = TempSum(i) + Val(Mid(tempstr, 3 * j + 4, 2))
        If Val(Mid(tempstr, 3 * j + 4, 2)) Mod 2 Then
            TempOdd(i) = TempOdd(i) + 1
        End If
    Next j
Next i
    'Odd5 = TempOdd(1) + TempOdd(2) + TempOdd(3) + TempOdd(4) + TempOdd(5)
For i = 1 To 10
    OddSum = OddSum + TempOdd(i)
    List6.AddItem Format(Str(i), "00") + ":" + Format(Str(TempSum(i)), "00") + "| " + Format(Str(TempOdd(i)), "00") + ":" + Format(Str(5 - TempOdd(i)), "00") + " | " + Format(Str(OddSum), "00") + "--" + Format(Str(5 * i - OddSum), "00")
Next i


End Function
Private Function ShowLoseList()
Dim i, j As Integer
Dim Tempdata0, Tempdata1, TempSum, Less10 As Integer
Dim tempstr0, tempstr1 As String
Dim tempstr, LineStr As String
For i = 1 To 10
    tempstr0 = List2.List(TotalNum - 11 + i)
    tempstr1 = List2.List(TotalNum - 12 + i)
    For j = 1 To 22
        Tempdata0 = Val(Mid(tempstr0, 4 * j + 1, 2))
        Tempdata1 = Val(Mid(tempstr1, 4 * j + 1, 2))
    If Tempdata0 = 0 Then
        tempstr = tempstr + Format(Str(j), "00") + " "
        LineStr = LineStr + Mid(tempstr1, 4 * j + 1, 2) + " "
        TempSum = TempSum + Tempdata1
        If Tempdata1 < 10 Then
            Less10 = Less10 + 1
        End If
    End If
    Next j
    LoseList.AddItem tempstr + " | " + LineStr + " | " + Str(Less10) + " | " + Format(Str(TempSum), "00") + Str(Format((TempSum / 5), "00.00"))
    tempstr = ""
    LineStr = ""
    Less10 = 0
    TempSum = 0
Next i
End Function



Private Sub Form_Load()
Dim i As Integer
Dim tempstr As String
Dim Tempdata As Integer
Dim Less10(10) As String
Open App.Path + "\data\22c5.txt" For Input As #1
  
Do While Not EOF(1)
    Line Input #1, tempstr
    List1.AddItem Left(tempstr, 21)
Loop
Close #1
TotalNum = List1.ListCount
tempstr = ""
For i = 1 To 22
    tempstr = tempstr + Format(Str(i), "00") + "| "
Next i
List5.AddItem "    " + tempstr


Call ShowData
List1.ListIndex = List1.ListCount - 1
List2.ListIndex = List2.ListIndex - 1
tempstr = List2.List(List2.ListCount - 1)
For i = 1 To 22
    Tempdata = Val(Mid(tempstr, i * 4 + 1, 2))
    If Tempdata < 10 Then
        Less10(Tempdata) = Less10(Tempdata) + Format(Str(i), "00") + " "
    Else
        Less10(10) = Less10(10) + Format(Str(i), "00") + " "
    End If
Next i
For i = 0 To 10
    List3.AddItem Format(Str(i), "00") + ": " + Less10(i)
Next i

Call ShowOE
Call ShowLoseList

End Sub

Private Sub Form_Unload(Cancel As Integer)
MainForm.Show
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
Dim tempstr As String
Dim tempstr0, tempstr1, tempstr2 As String
Dim Tempdata As Integer
Dim OddNum, EventNum As Integer
Dim Temp5(1 To 5) As Integer
Dim Less10(10) As String
Dim i As Integer
List1.ListIndex = List2.ListIndex
For i = 1 To 5
    Temp5(i) = Val(Mid(List1.List(List1.ListIndex), 3 * i + 4, 2))
    If Temp5(i) Mod 2 Then
        OddNum = OddNum + 1
       
    Else
        EventNum = EventNum + 1
    End If
Next i
Label3.Caption = "O:E" + Str(OddNum) + ":" + Str(EventNum)

    



List4.Clear
tempstr = List2.List(List2.ListIndex)

tempstr1 = List2.List(List2.ListIndex - 1)

tempstr2 = List2.List(List2.ListIndex - 2)

If List2.ListIndex = List2.ListCount - 1 Then
    tempstr0 = tempstr
Else
    tempstr0 = List2.List(List2.ListIndex + 1)
End If

For i = 0 To 2
    Label1(i).Caption = ""
    Label2(i).Caption = ""
Next i


For i = 1 To 22
    Tempdata = Val(Mid(tempstr, i * 4 + 1, 2))
    If Tempdata < 10 Then
        Less10(Tempdata) = Less10(Tempdata) + Format(Str(i), "00") + " "
        If Tempdata = 0 Then
            Label1(1).Caption = Label1(1).Caption + Format(Str(i), "00") + " "
            Label2(1).Caption = Label2(1).Caption + Mid(tempstr1, i * 4 + 1, 2) + " "
         ''
        End If
    Else
        Less10(10) = Less10(10) + Format(Str(i), "00") + " "
    End If
    Tempdata = Val(Mid(tempstr1, i * 4 + 1, 2))
        If Tempdata = 0 Then
            Label1(0).Caption = Label1(0).Caption + Format(Str(i), "00") + " "
            Label2(0).Caption = Label2(0).Caption + Mid(tempstr2, i * 4 + 1, 2) + " "
         ''
        End If
     Tempdata = Val(Mid(tempstr0, i * 4 + 1, 2))
        If Tempdata = 0 Then
            Label1(2).Caption = Label1(2).Caption + Format(Str(i), "00") + " "
            Label2(2).Caption = Label2(2).Caption + Mid(tempstr, i * 4 + 1, 2) + " "
         ''
        End If





Next i
For i = 0 To 10
    List4.AddItem Format(Str(i), "00") + ": " + Less10(i)
Next i

End Sub
