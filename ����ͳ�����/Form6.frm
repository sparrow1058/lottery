VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "滤号器"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form6"
   ScaleHeight     =   7275
   ScaleWidth      =   9225
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "随机号码2"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "长期出现次数分布"
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4"
         Height          =   180
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6"
         Height          =   180
         Index           =   5
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7"
         Height          =   180
         Index           =   6
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "详细区间分布选择"
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   42
         Top             =   1680
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   27
            Left            =   1560
            TabIndex        =   46
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   26
            Left            =   960
            TabIndex        =   45
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   25
            Left            =   480
            TabIndex        =   44
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   24
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   37
         Top             =   1440
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   23
            Left            =   1560
            TabIndex        =   41
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   22
            Left            =   960
            TabIndex        =   40
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   21
            Left            =   480
            TabIndex        =   39
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   31
         Top             =   480
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   35
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   6
            Left            =   960
            TabIndex        =   34
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   33
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   11
            Left            =   1560
            TabIndex        =   30
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   10
            Left            =   960
            TabIndex        =   29
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   28
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   15
            Left            =   1560
            TabIndex        =   25
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   14
            Left            =   960
            TabIndex        =   24
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   23
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   16
         Top             =   1200
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   19
            Left            =   1560
            TabIndex        =   20
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   18
            Left            =   960
            TabIndex        =   19
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   17
            Left            =   480
            TabIndex        =   18
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   6
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   5
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Label Label1 
         Caption         =   "01-05 06-10 11-15 16-20 21-25 26-30 31-33"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "过滤号码"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Image Image1 
         Height          =   690
         Index           =   0
         Left            =   120
         Picture         =   "Form6.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   690
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 13
 Load Image1(i)
 Image1(i).Left = (i - 1) * 690
 Image1(i).Visible = True
Next i
For i = 14 To 26
 Load Image1(i)
 Image1(i).Top = 930
 Image1(i).Left = (i - 14) * 690
 Image1(i).Visible = True
Next i

End Sub

