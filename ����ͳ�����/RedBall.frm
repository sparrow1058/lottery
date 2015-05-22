VERSION 5.00
Begin VB.Form RedBall 
   Caption         =   "RedBall"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   12465
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox NowList 
      Height          =   3480
      Left            =   9000
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox LoseList 
      Enabled         =   0   'False
      Height          =   3480
      Left            =   0
      TabIndex        =   14
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   1800
      TabIndex        =   3
      Top             =   5040
      Width           =   9135
      Begin VB.Label Label1 
         Caption         =   "9期平准值："
         Height          =   255
         Index           =   10
         Left            =   7320
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "8期平准值："
         Height          =   255
         Index           =   9
         Left            =   5640
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   " 7期平准值："
         Height          =   255
         Index           =   8
         Left            =   3720
         TabIndex        =   11
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "6期平准值："
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "5期平准值："
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "60期平准值："
         Height          =   255
         Index           =   5
         Left            =   7200
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "50期平准值："
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "40期平准值："
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "30期平准值："
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "20期平准值："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.ListBox RedAll 
      Height          =   4560
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   240
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   7215
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Height          =   1695
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "shu"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Index           =   1
      Left            =   10080
      TabIndex        =   17
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "shu"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Index           =   0
      Left            =   9000
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "RedBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalNum As Integer
Dim RedNum(1 To 1000, 1 To 16) As Integer
Dim LoseRedNum(1 To 1000, 1 To 16) As Integer

Private Function GetRedNum(TotalNum As Integer)
Dim i, j As Integer
Dim tempstr As String
Dim RedAllData(1 To 1000, 1) As String
Dim LoseSum As Integer

Dim tdata(7) As String
For i = 1 To TotalNum
    For j = 1 To 16
        RedNum(i, j) = 0
    Next j
Next i
For i = 1 To TotalNum
    tempstr = Form2.AllData.List(Form2.AllData.ListCount - TotalNum - 1 + i)
    Call GetNum(tempstr, tdata())
    RedNum(i, Val(tdata(7))) = 1
    RedAllData(i, 0) = tdata(0)
    RedAllData(i, 1) = tdata(7)
  
    
    
    
Next i
For j = 1 To 16
       If RedNum(1, j) = 1 Then
            LoseRedNum(1, j) = 0
        Else
            LoseRedNum(1, j) = 1
        End If
    For i = 2 To TotalNum
        If RedNum(i, j) = 1 Then
            LoseRedNum(i, j) = 0
        Else
            LoseRedNum(i, j) = LoseRedNum(i - 1, j) + 1
        End If
     Next i
Next j
tempstr = ""
For i = 1 To TotalNum
    For j = 1 To 16
        tempstr = tempstr + Format(Str(LoseRedNum(i, j)), "00") + "| "
    Next j
    List1.AddItem Format(Str(i), "000") + ": " + tempstr
    tempstr = ""
Next i
    
For i = 1 To TotalNum
    RedAll.AddItem RedAllData(i, 0) + " " + RedAllData(i, 1)
Next i

List1.ListIndex = List1.ListCount - 1
RedAll.ListIndex = RedAll.ListCount - 1
'Label1(0).Caption = Label1(0).Caption + vbCrLf
For i = 1 To 80
    For j = 1 To 16
        
        
        If LoseRedNum(TotalNum - i + 1, j) = 0 Then
             LoseSum = LoseSum + LoseRedNum(TotalNum - i, j)
        ''10期遗漏数字：
        'If i <= 10 Then
            Label1(0).Caption = Str(LoseRedNum(TotalNum - i, j)) + " " + Label1(0).Caption
        '    'Label1(0).Caption = Label1(0).Caption + Str(j) + ":" + Str(LoseRedNum(TotalNum - i, j)) + " "
        'End If
        If i > 4 And i < 10 Then
            Label1(i + 1).Caption = Label1(i + 1).Caption + Format(Str(LoseSum / i), "00.00")
        End If
        Select Case i
            Case 10
                ' Label1(0).Caption = Label1(0).Caption + Format(Str(LoseSum / i), "00.00")
            Case 20
                Label1(1).Caption = Label1(1).Caption + Format(Str(LoseSum / i), "00.00")
            Case 30
                Label1(2).Caption = Label1(2).Caption + Format(Str(LoseSum / i), "00.00")
            Case 40
                Label1(3).Caption = Label1(3).Caption + Format(Str(LoseSum / i), "00.00")
            Case 50
                Label1(4).Caption = Label1(4).Caption + Format(Str(LoseSum / i), "00.00")
            Case 60
                Label1(5).Caption = Label1(5).Caption + Format(Str(LoseSum / i), "00.00")
   
        End Select
        
        
        
        End If
        
        
        
       
    Next j
Next i

Dim LoseStr(9) As String
Dim tempmax(1 To 6) As String
For j = 1 To 16
    'Select Case LoseNum(80, j)
   ' Case 0
    If LoseRedNum(TotalNum, j) < 10 Then
        LoseStr(LoseRedNum(TotalNum, j)) = LoseStr(LoseRedNum(TotalNum, j)) + Format(Str(j), "00") + " "
    End If
    
  
    
    
    
    If LoseRedNum(TotalNum, j) >= 10 And LoseRedNum(TotalNum, j) < 60 Then
        tempmax(LoseRedNum(TotalNum, j) \ 10) = tempmax(LoseRedNum(TotalNum, j) \ 10) + Format(Str(j), "00") + " "
    End If
    If LoseRedNum(TotalNum, j) >= 60 Then
        tempmax(6) = tempmax(6) + Format(Str(j), "00") + " "
    End If
 
 Next j
 
 
For j = 0 To 9
    LoseList.AddItem Str(j) + ": " + LoseStr(j)
Next j
For i = 1 To 6
    LoseList.AddItem Str(i * 10) + ": " + tempmax(i)
Next i



End Function

Private Sub Form_Load()
 TotalNum = 100
 Call GetRedNum(TotalNum)
 List2.AddItem "     01| 02| 03| 04| 05| 06| 07| 08| 09| 10| 11| 12| 13| 14| 15| 16|"
End Sub

Private Sub List1_Click()
If List1.ListIndex > 1 Then
NowList.Clear
RedAll.ListIndex = List1.ListIndex
Dim i As Integer
Dim tempstr As String
Dim tempdata(1 To 16) As Integer
Dim Ctempstr(9) As String
Dim CtempMax(1 To 6) As String
tempstr = List1.List(List1.ListIndex)
For i = 1 To 16
    tempdata(i) = Val(Mid(tempstr, 4 * i + 2, 2))
    If tempdata(i) < 10 Then
        Ctempstr(tempdata(i)) = Ctempstr(tempdata(i)) + Format(Str(i), "00") + " "
        If tempdata(i) = 0 Then
            Label2(0).Caption = Format(Str(i), "00") + ":"
            Label2(1).Caption = Mid(List1.List(List1.ListIndex - 1), 4 * i + 2, 2)
        End If
    End If
    If tempdata(i) >= 10 And tempdata(i) < 60 Then
        CtempMax(tempdata(i) \ 10) = CtempMax(tempdata(i) \ 10) + Format(Str(i), "00") + " "
    End If
    If tempdata(i) >= 60 Then
        CtempMax(6) = CtempMax(6) + Format(Str(i), "00") + " "
    End If
    
    
       ' Ctempstr(10) = Ctempstr(10) + Format(Str(i), "00")
    
Next i
For i = 0 To 9
    NowList.AddItem Format(Str(i), "00") + ":" + Ctempstr(i)
Next i
For i = 1 To 6
    NowList.AddItem Format(Str(i * 10), "00") + ":" + CtempMax(i)
Next i

End If
End Sub
