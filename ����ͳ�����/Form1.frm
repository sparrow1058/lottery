VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10230
   StartUpPosition =   3  'µ¡¤f¯Ê¬Ù
   Begin VB.CommandButton Command2 
      Caption         =   "tongji"
      Height          =   975
      Left            =   8280
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.CommandButton Command1 
         Caption         =   "zuhe"
         Height          =   615
         Left            =   7320
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   6240
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function zuhe8(ChooseStr() As String, Result As Variant)

'Result 0
Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)

'Next i


End Function
Private Function zuhe9(ChooseStr() As String, Result As Variant)

Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)



End Function
Private Function zuhe10(ChooseStr() As String, Result As Variant)

Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)

'Next i


End Function


Private Sub Command1_Click()
Dim i As Integer
Dim ChoStr(9) As String
Dim Result(3, 5) As Variant
For i = 0 To 9
    ChoStr(i) = Text1(i).Text
Next i
Call zuhe8(ChoStr(), Result)

For i = 0 To 3
Text2.Text = Text2.Text + Result(i, 0) + " " + Result(i, 1) + " " + Result(i, 2) + " " + Result(i, 3) + " " + Result(i, 4) + " " + Result(i, 5) + vbCrLf
Next i
    
    



End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 9
Text1(i).Font.Size = 12
Next i
End Sub

Private Sub Text1_Change(Index As Integer)
If (Len(Text1(Index).Text) = 2) And Index < 9 Then
    Text1(Index + 1).SetFocus
End If
If Len(Text1(9).Text) > 2 Then
    Text1(9).Text = ""
End If

End Sub

