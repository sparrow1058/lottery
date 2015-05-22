VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "统计随机窗口"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10200
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      Caption         =   "尾号选择窗口"
      Height          =   2655
      Left            =   3240
      TabIndex        =   102
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame6 
         Height          =   2415
         Left            =   2280
         TabIndex        =   104
         Top             =   240
         Width           =   1335
         Begin VB.CheckBox Check3 
            Caption         =   "8"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   114
            Top             =   1680
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "6"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   113
            Top             =   1320
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "4"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   112
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   111
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "9"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   109
            Top             =   1680
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "7"
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   108
            Top             =   1320
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "5"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   107
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "3"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   106
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox Check3 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   105
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox List4 
         Height          =   2400
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   3720
         TabIndex        =   115
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   99
      Top             =   7200
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CheckBox Check2 
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   100
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4095
      Left            =   3240
      TabIndex        =   94
      Top             =   3240
      Width           =   6615
      Begin VB.CommandButton Command3 
         Caption         =   "随机号码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Height          =   1215
         Left            =   0
         Picture         =   "Form7.frx":29C12
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Height          =   1335
         Left            =   0
         Picture         =   "Form7.frx":2F716
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   0
         Width           =   1335
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   3975
         Left            =   1320
         TabIndex        =   96
         Top             =   0
         Width           =   4140
         Begin VB.Image Image1 
            Height          =   690
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   3975
         Left            =   5760
         TabIndex        =   95
         Top             =   0
         Width           =   690
         Begin VB.Image BlueBall 
            Height          =   690
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   690
         End
         Begin VB.Image BlueBall 
            Height          =   690
            Index           =   2
            Left            =   0
            Top             =   480
            Width           =   690
         End
         Begin VB.Image BlueBall 
            Height          =   690
            Index           =   3
            Left            =   0
            Top             =   1080
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "总计出现数据"
      Height          =   2775
      Left            =   10320
      TabIndex        =   47
      Top             =   2760
      Width           =   6495
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   6
         Left            =   3960
         TabIndex        =   86
         Top             =   2160
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   27
            Left            =   1800
            TabIndex        =   90
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   26
            Left            =   1320
            TabIndex        =   89
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   25
            Left            =   840
            TabIndex        =   88
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   24
            Left            =   360
            TabIndex        =   87
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "7:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   91
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   5
         Left            =   1680
         TabIndex        =   80
         Top             =   2160
         Width           =   2295
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   23
            Left            =   1800
            TabIndex        =   84
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   22
            Left            =   1320
            TabIndex        =   83
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   21
            Left            =   840
            TabIndex        =   82
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   20
            Left            =   360
            TabIndex        =   81
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "6:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   85
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   74
         Top             =   1680
         Width           =   2295
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   15
            Left            =   1800
            TabIndex        =   78
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   14
            Left            =   1320
            TabIndex        =   77
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   13
            Left            =   840
            TabIndex        =   76
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   75
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "4:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   79
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   4
         Left            =   3960
         TabIndex        =   68
         Top             =   1680
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   19
            Left            =   1800
            TabIndex        =   72
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   18
            Left            =   1320
            TabIndex        =   71
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   70
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   69
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "5:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   2
         Left            =   3960
         TabIndex        =   62
         Top             =   1200
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   11
            Left            =   1800
            TabIndex        =   66
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   10
            Left            =   1320
            TabIndex        =   65
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   64
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   63
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "3:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   1
         Left            =   3960
         TabIndex        =   56
         Top             =   720
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   60
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   59
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   58
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   57
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "2:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Height          =   495
         Index           =   0
         Left            =   3960
         TabIndex        =   50
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   54
            Top             =   120
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   53
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   52
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   51
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "1:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.ListBox List3 
         Height          =   1500
         Left            =   1680
         TabIndex        =   49
         Top             =   240
         Width           =   2175
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "遗漏统计"
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.ListBox List1 
         Height          =   2040
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Caption         =   "遗漏次数"
         Height          =   2055
         Left            =   0
         TabIndex        =   1
         Top             =   2400
         Width           =   2775
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   32
            Left            =   1560
            TabIndex        =   46
            Top             =   1560
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   31
            Left            =   1200
            TabIndex        =   45
            Top             =   1560
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   30
            Left            =   840
            TabIndex        =   43
            Top             =   1560
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   29
            Left            =   2400
            TabIndex        =   42
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   28
            Left            =   2040
            TabIndex        =   41
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   27
            Left            =   1680
            TabIndex        =   39
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   26
            Left            =   2400
            TabIndex        =   38
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   25
            Left            =   2040
            TabIndex        =   37
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   24
            Left            =   1680
            TabIndex        =   35
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   23
            Left            =   2400
            TabIndex        =   34
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   22
            Left            =   2040
            TabIndex        =   33
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   21
            Left            =   1680
            TabIndex        =   31
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   20
            Left            =   2400
            TabIndex        =   30
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   19
            Left            =   2040
            TabIndex        =   29
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   18
            Left            =   1680
            TabIndex        =   27
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   17
            Left            =   2400
            TabIndex        =   26
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   15
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   14
            Left            =   1080
            TabIndex        =   22
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   13
            Left            =   720
            TabIndex        =   21
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   19
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   11
            Left            =   1080
            TabIndex        =   18
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   17
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   15
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   14
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   13
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   11
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "10:"
            Height          =   255
            Index           =   10
            Left            =   600
            TabIndex        =   44
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "9:"
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   40
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "8:"
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   36
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "7:"
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   32
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "6:"
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   28
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "5:"
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "4:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "3:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "2:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "1:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "0:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2655
         Left            =   0
         TabIndex        =   92
         Top             =   4440
         Width           =   2775
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "The Num"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   3240
      TabIndex        =   93
      Top             =   2760
      Width           =   6495
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AllTask As Integer
Dim BallAll(60) As Boolean
Dim Manual As Boolean
Private Sub ZuHe(TStr As String, outstr() As String)
 Dim tempstr As String
 Dim i As Integer
 Dim Data(1 To 15) As String
 TStr = Replace(TStr, " ", "")
 If Len(TStr) > 16 Then
   TStr = Left(TStr, 16)
  End If
 For i = 1 To Len(TStr) / 2
  Data(i) = Mid(TStr, 2 * i - 1, 2)
 Next i
 If Len(TStr) = 12 Then
   outstr(1) = Data(1) + Data(2) + Data(3) + Data(4) + Data(5) + Data(6)
 End If
 If Len(TStr) = 14 Then         ''7待选 5选5中
  outstr(1) = Data(1) + Data(2) + Data(3) + Data(4) + Data(5) + Data(7)

  
  
  outstr(2) = Data(1) + Data(2) + Data(3) + Data(4) + Data(6) + Data(7)
  outstr(3) = Data(1) + Data(2) + Data(3) + Data(5) + Data(6) + Data(7)
  outstr(4) = Data(1) + Data(2) + Data(4) + Data(5) + Data(6) + Data(7)
  outstr(5) = Data(1) + Data(3) + Data(4) + Data(5) + Data(6) + Data(7)
'  outstr(6) = Data(2) + Data(3) + Data(4) + Data(5) + Data(6) + Data(7)
 End If
 
  
  
 If Len(TStr) = 16 Then         ''6选5中数据
  outstr(1) = Data(1) + Data(2) + Data(3) + Data(5) + Data(7) + Data(8)
  outstr(2) = Data(1) + Data(2) + Data(4) + Data(6) + Data(7) + Data(8)
  outstr(3) = Data(1) + Data(3) + Data(4) + Data(5) + Data(6) + Data(8)
  outstr(4) = Data(2) + Data(3) + Data(4) + Data(5) + Data(6) + Data(7)
 End If
End Sub
Private Sub ShowBall(StartN As Integer, AllData As String)
Dim NID As Integer
Dim TTop As Integer
Dim StartNum As Integer


TTop = (StartN - 1) * 690
StartNum = StartN * 6
AllData = Replace(AllData, " ", "")
If AllData = "" Then
 Exit Sub
End If


Dim i As Integer
For i = 1 To 6
  
  NID = Val(Mid(AllData, 2 * i - 1, 2)) + 100
  If BallAll(StartNum + i) = False Then
   Load Image1(StartNum + i)
   BallAll(StartNum + i) = True
  End If
  
Image1(StartNum + i).Picture = LoadResPicture(NID, 0)
Image1(StartNum + i).Top = TTop
Image1(StartNum + i).Left = (i - 1) * 690
Image1(StartNum + i).Visible = True
Next i

End Sub


Private Sub ShowFrame1()
Dim i As Integer
Dim tempstr As String
 Load Form2
For i = 0 To 10
 List1.AddItem Form2.yllist.List(i)
Next i
End Sub



Private Sub ShowTheSum()
Dim tempstr, TempStr1, TempStr2 As String
Dim i As Integer
Dim Num33(1 To 33) As Integer
If Label3.Caption = "" Or Label4.Caption = "" Then
  Label5.Caption = ""
Else
    TempStr1 = Replace(Label3.Caption, " ", "")
    TempStr1 = Replace(TempStr1, vbCrLf, "")
    TempStr2 = Replace(Label4.Caption, " ", "")
    TempStr2 = Replace(TempStr2, vbCrLf, "")
  For i = 1 To Len(TempStr1) / 2
      Num33(Val(Mid(TempStr1, 2 * i - 1, 2))) = Num33(Val(Mid(TempStr1, 2 * i - 1, 2))) + 1
  Next i
  For i = 1 To Len(TempStr2) / 2
      Num33(Val(Mid(TempStr2, 2 * i - 1, 2))) = Num33(Val(Mid(TempStr2, 2 * i - 1, 2))) + 1
  Next i
 For i = 1 To 33
  If Num33(i) = 2 Then
      tempstr = tempstr + Format(Str(i), "00") + " "
  End If
 Next i
  Label5.Caption = tempstr
End If






End Sub
Private Sub ShowFrame8()
Dim i As Integer
Load Form5
For i = 12 To 1 Step -1
 List2.AddItem Replace(Form5.List3.List(Form5.List3.ListCount - i), " ", "")
Next i
For i = 0 To 6
 List3.AddItem Form5.List4.List(i)
Next i


End Sub

Private Sub Check1_Click(Index As Integer)
 
 Dim i As Integer
 Label3.Caption = ""
 For i = 0 To 32 Step 3
   If Check1(i).Value = 1 Or Check1(i + 1).Value = 1 Or Check1(i + 2).Value = 1 Then
     Label3.Caption = Label3.Caption + Right(List1.List(i / 3), Len(List1.List(i / 3)) - 3) + vbCrLf
   End If
 Next i
Call ShowTheSum

End Sub


Private Sub Check2_Click(Index As Integer)
Dim i As Integer
Dim tempstr As String
 For i = 0 To 32
  If Check2(i).Value = 1 Then
   tempstr = tempstr + Format(Str(i + 1), "00")
  End If
 Next i
If Len(tempstr) > 14 Then
 Command2.Picture = LoadResPicture(134, 0)
Else
  Command2.Picture = LoadResPicture(135, 0)
End If

End Sub

Private Sub Check3_Click(Index As Integer)
Dim i As Integer
Dim tempstr As String
Dim allstr As String
Label4.Caption = ""

For i = 0 To 9
    If Check3(i).Value = 1 Then
      ' tempstr= string(i) +string(10+i) +" "+ string (20+i)+" "+ String(30+i)
        If i > 0 Then
            tempstr = Format(Str(i), "00") + " "
        End If
        tempstr = tempstr + Format(Str(i + 10), "00") + " "
        tempstr = tempstr + Format(Str(i + 20), "00") + " "
        If i < 4 Then
            tempstr = tempstr + Format(Str(i + 30), "00") + " "
        End If
        
    allstr = allstr + tempstr + vbCrLf
    End If

    tempstr = ""
Next i
Label4.Caption = allstr
Call ShowTheSum

End Sub

Private Sub Command1_Click()
 Dim outstr(5) As String
 Dim tempstr As String
 Dim i As Integer
 
If Len(Label5.Caption) > 20 Then
 
 Call ZuHe(Label5.Caption, outstr())
 For i = 1 To 4
 Call ShowBall(i, outstr(i))
 Next i
End If
 
 
 



' BlueBall(1).Picture = LoadResPicture(111, 0)
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyW Then
 Form2.Show
End If
If Shift = 2 And KeyCode = vbKeyB Then
 'baby.Show
End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 Form2.Show
End If
End Sub

Private Sub Command2_Click()
 If Frame3.Visible = False Then
  Frame3.Visible = True
  Form7.Height = Form7.Height + Frame3.Height
  
 End If
 
 Dim tempstr As String
 Dim i As Integer
 Dim outstr(5) As String
 For i = 0 To 32
  If Check2(i).Value = 1 Then
   tempstr = tempstr + Format(Str(i + 1), "00")
  End If
 Next i
 
 If Len(tempstr) > 11 Then
    
     Call ZuHe(tempstr, outstr())
 For i = 1 To 4
    Call ShowBall(i, outstr(i))
 Next i
 End If
End Sub

Private Sub Command3_Click()
Form8.Show
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Startp As Integer
 For i = 1 To 32
  Load Check2(i)
  If i > 16 Then
   Startp = 17
   Check2(i).Top = 400
  End If
  
  Check2(i).Caption = Format(Str(i + 1), "00")
  Check2(i).Left = (i - Startp) * 615
  Check2(i).Visible = True
  
  
 Next i
  
Call ShowFrame1
Call ShowFrame8
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form5
Unload Form2
End Sub

Private Sub Frame3_DblClick()
Frame3.Visible = False

Form7.Height = Form7.Height - Frame3.Height
End Sub

Private Sub Option3_Click(Index As Integer)
Dim i As Integer
Label4.Caption = ""
 For i = 0 To 27 Step 4
   If Option3(i + 1).Value = True Or Option3(i + 2).Value = True Or Option3(i + 3).Value = True Then
        Label4.Caption = Label4.Caption + Right(List3.List(i / 4), Len(List3.List(i / 4)) - 3) + vbCrLf
  End If
 Next i
 Call ShowTheSum

End Sub

