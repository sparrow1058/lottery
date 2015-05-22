VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form5"
   ScaleHeight     =   7110
   ScaleWidth      =   9195
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   3735
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2760
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2760
         TabIndex        =   13
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2760
         TabIndex        =   7
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TailNum(9) As Integer
Dim TailStr As String
Dim LeftTailStr As String

Private Sub Command1_Click(Index As Integer)
List1.Clear
Call ShowCombinData(Index + 2)


End Sub

Private Sub Command2_Click(Index As Integer)
List2.Clear
Call ShowLeftData(Index + 2)
End Sub

Private Sub Form_Load()
Dim TempStr As String
Dim LineStr As String
Dim i As Integer
For i = 0 To 9
 TailNum(i) = 0
 Next i
 TailStr = ""
 LeftTailStr = ""

TempStr = Left(Form1.List1.List(Form1.List1.ListCount - 1), 19)
LineStr = Left(TempStr, 5)
For i = 0 To 5
 LineStr = LineStr + " " + Mid(TempStr, 6 + 2 * i, 2)
 TailNum(Val(Mid(TempStr, 7 + 2 * i, 1))) = 1
Next i
LineStr = LineStr + " + " + Right(TempStr, 2)
Label1.Caption = LineStr
LineStr = ""
For i = 0 To 9
If TailNum(i) = 1 Then
 TailStr = TailStr + Format(Str(i), "@")
 Else
 LeftTailStr = LeftTailStr + Str(i)
End If
Next i
Label2.Caption = TailStr
TailStr = Replace(TailStr, " ", "")
LeftTailStr = Replace(LeftTailStr, " ", "")
End Sub

Private Sub ShowCombinData(nums As Integer)
Dim AllNums As Integer
Dim i As Integer
Dim flagstr As String
If Len(TailStr) < nums Then
 Exit Sub
End If
For i = 1 To nums
 flagstr = flagstr + "1"
Next i
For i = nums + 1 To Len(TailStr)
 flagstr = flagstr + "0"
 Next i
 AllNums = CCInOut(Len(TailStr), nums)
For i = 1 To AllNums


List1.AddItem Combin1Num(TailStr, flagstr)
flagstr = ChangeStr(flagstr)
Next i
Label3.Caption = Str(AllNums)
End Sub
Private Sub ShowLeftData(nums As Integer)
Dim AllNums As Integer
Dim i As Integer
Dim flagstr As String
If Len(LeftTailStr) < nums Then
 Exit Sub
End If
For i = 1 To nums
 flagstr = flagstr + "1"
Next i
For i = nums + 1 To Len(LeftTailStr)
 flagstr = flagstr + "0"
 Next i
 AllNums = CCInOut(Len(LeftTailStr), nums)
For i = 1 To AllNums


List2.AddItem Combin1Num(LeftTailStr, flagstr)
flagstr = ChangeStr(flagstr)
Next i
Label4.Caption = Str(AllNums)
End Sub
