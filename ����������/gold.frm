VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBALL"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "gold.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12090
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9960
      TabIndex        =   141
      Text            =   "尾号分布"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "尾号分布"
      Height          =   735
      Left            =   9960
      TabIndex        =   140
      Top             =   0
      Width           =   1935
   End
   Begin VB.Frame Frame17 
      Caption         =   "遗漏检验"
      Height          =   1455
      Left            =   10320
      TabIndex        =   135
      Top             =   7080
      Width           =   1095
      Begin VB.CheckBox Check10 
         Caption         =   ">10"
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check9 
         Caption         =   "4-9"
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Top             =   600
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox Check8 
         Caption         =   "<3"
         Height          =   255
         Left            =   240
         TabIndex        =   136
         Top             =   240
         Value           =   1  'Checked
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set the Region"
      Height          =   3255
      Left            =   12360
      TabIndex        =   120
      Top             =   120
      Width           =   2655
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "gold.frx":4C4A
         Left            =   120
         List            =   "gold.frx":4C4C
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         TabIndex        =   123
         Top             =   600
         Width           =   1215
         Begin VB.OptionButton OptionReg1 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   124
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2535
         Left            =   1320
         TabIndex        =   121
         Top             =   600
         Width           =   1215
         Begin VB.OptionButton OptionReg2 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   122
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1440
         TabIndex        =   126
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "检验"
      Height          =   1095
      Left            =   14400
      TabIndex        =   117
      Top             =   6360
      Width           =   615
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox RegCheck 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "遗漏统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      TabIndex        =   116
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "选择组合"
      Height          =   735
      Left            =   10200
      TabIndex        =   115
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   64
      Top             =   7080
      Width           =   10335
      Begin VB.Frame Frame15 
         Caption         =   ">=10"
         Height          =   1335
         Left            =   9480
         TabIndex        =   102
         Top             =   0
         Width           =   735
         Begin VB.OptionButton Option6 
            Caption         =   "2"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   375
         End
         Begin VB.OptionButton Option6 
            Caption         =   "1"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   104
            Top             =   480
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option6 
            Caption         =   "0"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "4-9"
         Height          =   1215
         Left            =   8760
         TabIndex        =   97
         Top             =   0
         Width           =   735
         Begin VB.OptionButton Option5 
            Caption         =   "3"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   101
            Top             =   960
            Width           =   375
         End
         Begin VB.OptionButton Option5 
            Caption         =   "2"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   720
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option5 
            Caption         =   "1"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   99
            Top             =   480
            Width           =   375
         End
         Begin VB.OptionButton Option5 
            Caption         =   "0"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "<=3"
         Height          =   1335
         Left            =   8040
         TabIndex        =   92
         Top             =   0
         Width           =   735
         Begin VB.OptionButton Option4 
            Caption         =   "5"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   495
         End
         Begin VB.OptionButton Option4 
            Caption         =   "4"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   95
            Top             =   720
            Width           =   495
         End
         Begin VB.OptionButton Option4 
            Caption         =   "3"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   94
            Top             =   480
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option4 
            Caption         =   "2"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "大小奇偶 大数 奇数"
         Height          =   1215
         Left            =   2160
         TabIndex        =   83
         Top             =   0
         Width           =   2655
         Begin VB.CheckBox OddC 
            Height          =   255
            Left            =   2160
            TabIndex        =   91
            Top             =   720
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox BigC 
            Height          =   255
            Left            =   2160
            TabIndex        =   90
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check4 
            Caption         =   ">3"
            Height          =   255
            Left            =   1560
            TabIndex        =   89
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox Check3 
            Caption         =   ">3"
            Height          =   255
            Left            =   1560
            TabIndex        =   88
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox BigCheck 
            Caption         =   "Y"
            Height          =   255
            Left            =   1080
            TabIndex        =   87
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox OddCheck 
            Caption         =   "Y"
            Height          =   255
            Left            =   1080
            TabIndex        =   86
            Top             =   720
            Width           =   375
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "gold.frx":4C4E
            Left            =   120
            List            =   "gold.frx":4C67
            TabIndex        =   85
            Text            =   "大小比"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "gold.frx":4C8E
            Left            =   120
            List            =   "gold.frx":4CA7
            TabIndex        =   84
            Text            =   "奇偶比"
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "存在与虚无"
         Height          =   1215
         Left            =   4800
         TabIndex        =   74
         Top             =   0
         Width           =   3015
         Begin VB.TextBox Text5 
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   80
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   79
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   78
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   76
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   75
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "无"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   82
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "存在"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   81
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "遗漏"
         Height          =   1215
         Index           =   1
         Left            =   1560
         TabIndex        =   71
         Top             =   0
         Width           =   615
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   73
            Text            =   "0"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "和值"
         Height          =   1215
         Index           =   0
         Left            =   960
         TabIndex        =   68
         Top             =   0
         Width           =   615
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Text            =   "1"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   735
         Begin VB.Label Label3 
            Caption         =   "最大值"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   67
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "最小值"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame16 
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   8880
      Width           =   10215
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   33
         Left            =   9000
         TabIndex        =   58
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   32
         Left            =   8400
         TabIndex        =   57
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   31
         Left            =   7800
         TabIndex        =   56
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   30
         Left            =   7200
         TabIndex        =   55
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   29
         Left            =   6720
         TabIndex        =   54
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   28
         Left            =   6120
         TabIndex        =   53
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   27
         Left            =   5520
         TabIndex        =   52
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   26
         Left            =   4920
         TabIndex        =   51
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   25
         Left            =   4320
         TabIndex        =   50
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   24
         Left            =   3720
         TabIndex        =   49
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   23
         Left            =   3120
         TabIndex        =   48
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   47
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   21
         Left            =   1920
         TabIndex        =   46
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   20
         Left            =   1320
         TabIndex        =   45
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   19
         Left            =   720
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   17
         Left            =   9600
         TabIndex        =   42
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   16
         Left            =   9000
         TabIndex        =   41
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   15
         Left            =   8400
         TabIndex        =   40
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   14
         Left            =   7800
         TabIndex        =   39
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   13
         Left            =   7200
         TabIndex        =   38
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   12
         Left            =   6720
         TabIndex        =   37
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   11
         Left            =   6120
         TabIndex        =   36
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   10
         Left            =   5520
         TabIndex        =   35
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   34
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   33
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   32
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   31
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   30
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   29
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   28
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   27
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "01"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.Label LabelChoose 
         Caption         =   "0"
         Height          =   255
         Left            =   9720
         TabIndex        =   60
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.ListBox List4 
      Height          =   420
      Left            =   14640
      TabIndex        =   16
      Top             =   9240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   13200
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Frame Frame8 
         Caption         =   "本邻"
         Height          =   1215
         Index           =   0
         Left            =   600
         TabIndex        =   131
         Top             =   0
         Width           =   615
         Begin VB.OptionButton Option3 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   134
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   133
            Top             =   480
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   132
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "区间"
         Height          =   1215
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Width           =   615
         Begin VB.OptionButton Option2 
            Caption         =   "3"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "4"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   129
            Top             =   480
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "5"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   128
            Top             =   720
            Width           =   375
         End
      End
   End
   Begin VB.ListBox List3 
      Height          =   2760
      Left            =   13560
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   12360
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton Command12 
         Caption         =   "BlueReg"
         Height          =   855
         Index           =   2
         Left            =   2160
         TabIndex        =   114
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Red Star"
         Height          =   855
         Index           =   1
         Left            =   1560
         TabIndex        =   113
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "RedReg"
         Height          =   735
         Index           =   0
         Left            =   2160
         TabIndex        =   112
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "100遗漏"
         Height          =   735
         Left            =   1560
         TabIndex        =   111
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   735
         Left            =   840
         TabIndex        =   110
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "奇偶  大小"
         Height          =   735
         Left            =   0
         TabIndex        =   109
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "连号"
         Height          =   855
         Left            =   960
         TabIndex        =   108
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "上下连号"
         Height          =   735
         Left            =   2040
         TabIndex        =   63
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check7 
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   1320
         TabIndex        =   61
         Text            =   "Combo4"
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "11区分布"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         Caption         =   "4分区"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "3区分布"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   1440
         TabIndex        =   21
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg62"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   720
         TabIndex        =   20
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg61"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "分布遗漏"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "10期数据"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Appear"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   4440
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Lost"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   4440
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sum"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   4440
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "分布遗漏"
         Height          =   735
         Left            =   2040
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Get Num"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   5
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Sum Sort"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   3
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "篮球"
         Height          =   615
         Left            =   2280
         TabIndex        =   107
         Top             =   4920
         Width           =   255
      End
   End
   Begin VB.ListBox List2 
      Height          =   4200
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   11775
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   375
      Left            =   8040
      TabIndex        =   139
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "待命"
      Height          =   255
      Left            =   0
      TabIndex        =   106
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "选择数字"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   8520
      Width           =   7575
   End
   Begin VB.Label Label7 
      Caption         =   "000: 00"
      Height          =   375
      Left            =   12480
      TabIndex        =   15
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label6 
      Height          =   3975
      Left            =   13200
      TabIndex        =   13
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "000: 00"
      Height          =   375
      Left            =   13440
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SUMLINENUMS = 15
Const TAILLINES = 188

Dim TotalNum As Integer
Dim LoseR(1 To 888, 1 To 33) As Integer
Dim LoseB(1 To 888, 1 To 16) As Integer
Dim LoseRFlag(1 To 888, 1 To 33) As Integer
Dim LoseBFlag(1 To 888, 1 To 16) As Integer
Dim Data(1 To 888) As String
Dim CheckStr(0 To 5) As String
Dim LoseStr(33) As String
Const SumNum = 1000
Dim AllSum(1 To 33) As Integer
Dim SumRegion(1 To 33) As Integer
Dim NewSumRegion(1 To 33) As Integer
Dim NowRegion As String
Dim NowRegion1, NowRegion2 As String
Public CurrentNum As String
Dim SumMin, SumMax As Integer
Dim LostSumMin, LostSumMax As Integer
Dim List2_Width As Integer
Dim Picture_Y As Integer
Dim Num1(100) As Integer
Dim Num2(100) As Integer
Dim NowOddNums As Integer
Dim NowBigNums As Integer
Dim AddFrame16 As Boolean
Dim ChooseOp7 As Boolean
Dim ShowLostForm As Boolean
Dim TailMode As Boolean







Private Function CheckStrInit()
   Dim i As Integer
   CheckStr(0) = "000"
   CheckStr(1) = "001 010 100"
   CheckStr(2) = "002 011 020 101 110 200"
   CheckStr(3) = "003 012 021 030 102 111 120 201 210 300"
   CheckStr(4) = "013 022 031 103 112 121 130 202 211 220"
   CheckStr(5) = "023 032 113 122 131 203 212 221 230 302"
   Text1(0).Text = 70
   Text1(1).Text = 140
   Text1(2).Text = 16
   Text1(3).Text = 44
   For i = 1 To 33
    Check6(i).Caption = Format(Str(i))
   Next i
   
   
' if 1 then   001 010 100
' if 2 then   002 011 020 101 110 200
' if 3 then   003 012 021 030 102 111 120 201 210 300
' if 4 then   013 022 031 103 112 121 130 202 211 220
' if 5 then   023 032 113 122 131 203 212 221 230

End Function

Private Sub CheckOptionReg()




NowRegion = NowRegion1 + NowRegion2 + Mid(Combo1.List(Combo1.ListIndex), 3, 1)
NowRegion = Replace(NowRegion, " ", "")
If Not (NowRegion = "1111110") And Len(NowRegion) = 7 Then
    
    Label2.Caption = NowRegion
   ' Command3.Enabled = True
    Command4.Enabled = True
Else
      Label2.Caption = NowRegion
   ' Command3.Enabled = False
    Command4.Enabled = False
    
End If


End Sub

Private Sub BigC_Click()
If BigC.Value = 1 Then
  BigCheck.Enabled = True
  Check3.Enabled = True
  Combo3.Enabled = True
Else
 BigCheck.Enabled = False
  Check3.Enabled = False
  Combo3.Enabled = False
End If
End Sub

Private Sub BigCheck_Click()
If BigCheck.Value = 1 Then
    NowBigNums = 3
    Combo3.Enabled = True
    Combo3.ListIndex = 3
    Check3.Enabled = False
Else
    Check3.Enabled = True
    Combo3.Enabled = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check2.Caption = "YES"
Else
    Check2.Caption = "NO"
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
 Check3.Caption = "<3"
 
Else
 Check3.Caption = ">3"
 
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
 Check4.Caption = "<3"
Else
 Check4.Caption = ">3"
End If
End Sub

Private Sub Check6_Click(Index As Integer)
Dim i As Integer
Dim j As Integer
Label10.Caption = ""
For i = 1 To 33
 If Check6(i).Value = 1 Then
    Label10.Caption = Label10.Caption + Format(Str(i), "00")
    j = j + 1
 End If
Next i
LabelChoose.Caption = Format(Str(j))
End Sub

Private Sub Combo1_Click()
Dim i, CStrNum As Integer
Dim tempstr As String
Dim Check1Str As String
Label2.Caption = ""
NowRegion = ""
For i = 0 To 9
   ' Check1(i).Visible = False
    'Check1(i).Value = 0
    OptionReg1(i).Visible = False
    OptionReg2(i).Visible = False
    OptionReg1(i).Value = 0
    OptionReg2(i).Value = 0
Next i

tempstr = Combo1.List(Combo1.ListIndex)
 If Val(Mid(tempstr, 1, 1)) < 6 And Val(Mid(tempstr, 2, 1)) < 6 Then

 Check1Str = Replace(CheckStr(Val(Mid(tempstr, 1, 1))), " ", "")
 For i = 1 To Len(Check1Str) / 3
  OptionReg1(i - 1).Caption = Mid(Check1Str, 3 * i - 2, 3)
  OptionReg1(i - 1).Visible = True
 Next i

Check1Str = Replace(CheckStr(Val(Mid(tempstr, 2, 1))), " ", "")
For i = 1 To Len(Check1Str) / 3
'  Check1(9 + i).Caption = Mid(Check1Str, 3 * i - 2, 3)
 ' Check1(9 + i).Visible = True
 OptionReg2(i - 1).Caption = Mid(Check1Str, 3 * i - 2, 3)
  OptionReg2(i - 1).Visible = True

 Next i
End If
End Sub



Private Sub Combo2_Click()
NowOddNums = Combo2.ListIndex
End Sub



Private Sub Combo3_Click()

NowBigNums = Combo3.ListIndex
End Sub

Private Sub Command1_Click()
List2.Clear
Picture1.Visible = False
List2.Width = List2_Width
List4.Visible = False
    Call ShowRBLost
'    Call Show10Lost



End Sub
Private Sub Show10Lost()
Dim i, j, k  As Integer
Const QQNum = 2
Dim tempstr As String
Dim LineStr As String
Dim Lost3(5) As Integer
Dim Lost3Sum(5) As Integer
For i = 0 To List1.ListCount - 1
  tempstr = Right(Left(List1.List(i), 43), 17)
  tempstr = Replace(tempstr, " ", "")
   Call CheckLost3(Lost3(), tempstr)
   For j = 0 To 5
    Lost3Sum(j) = Lost3Sum(j) + Lost3(j)
   Next j
   
   k = k + 1
   If k = QQNum Then
    k = 0
    For j = 0 To 5
     LineStr = LineStr + Format(Str(Lost3Sum(j)), "00") + " |  "
     Lost3Sum(j) = 0
    Next j
    
    List2.AddItem LineStr
    LineStr = ""
   End If

Next i
    For j = 0 To 5
     LineStr = LineStr + Format(Str(Lost3Sum(j)), "00") + " |  "
     Lost3Sum(j) = 0
    Next j
    
    List2.AddItem LineStr
   
    List2.AddItem "**********************************"
    List2.AddItem "00 |  01 |  02 |  03 |04-09| >10 |" + "   " + Str(QQNum) + "期数据统计  " + Str(k) + "期数据剩下”"
     List2.ListIndex = List2.ListCount - 1

End Sub


Private Sub Command10_Click()
Dim i, j, k As Integer
Dim tempstr As String
Dim LineStr As String
Dim Data4(4) As Integer
List2.Clear
For i = 0 To List1.ListCount - 1
    tempstr = Right(List1.List(i), 4)
    For j = 1 To 4
        Data4(j) = Data4(j) + Val(Mid(tempstr, j, 1))
    Next j
    k = k + 1
    If k = 7 Then
        k = 0
        For j = 1 To 4
         LineStr = LineStr + Format(Str(Data4(j)), "00") + " "
            Data4(j) = 0
        Next j
       List2.AddItem LineStr
       LineStr = ""
    
    End If

 Next i
         For j = 1 To 4
         LineStr = LineStr + Format(Str(Data4(j)), "00") + " "
            Data4(j) = 0
        Next j
       List2.AddItem LineStr
    


End Sub

Private Sub Command11_Click()
List2.Clear
Dim i, j, k, L As Integer
Dim tempstr As String
Dim LineStr As String
For j = 1 To 33
If j = 6 Then
  j = 6
End If
For i = List1.ListCount - 1 To List1.ListCount - 80 Step -1
    tempstr = Right(Left(List1.List(i), 17), 12)
    For k = 1 To 6
     If Val(Mid(tempstr, 2 * k - 1, 2)) = j Then
         LineStr = Format(Str(L), "00") + " " + LineStr
        L = -1
        Exit For
     End If
    Next k
    L = L + 1
Next i
L = 0
List2.AddItem "R" + Format(Str(j), "00") + " " + LineStr
LineStr = ""
Next j
List2.AddItem "*****************************************************************************"
For j = 1 To 16
For i = List1.ListCount - 1 To List1.ListCount - 200 Step -1
    tempstr = Right(Left(List1.List(i), 19), 2)
        If Val(tempstr) = j Then
             LineStr = Format(Str(L), "00") + " " + LineStr
            L = -1
         End If

    L = L + 1
Next i
L = 0
List2.AddItem "B" + Format(Str(j), "00") + " " + LineStr
LineStr = ""
Next j
LineStr = ""
For i = 0 To 32
 tempstr = Replace(List2.List(i), " ", "")
 LineStr = LineStr + WriteLostTR(tempstr)
Next i

Call CreateHtmlFile(LineStr, "Result4.html", 104)









End Sub



Private Sub Command12_Click(Index As Integer)
Dim tempstr As String
Dim GetRegStr As String
Dim WebStr As String
Dim BlueStr As String
Dim i As Integer
If Index = 0 Then
    For i = List1.ListCount - 44 To List1.ListCount - 1
        tempstr = Left(List1.List(i), 17)
        WebStr = WebStr + WriteRTR(tempstr)
        Next i
        Call CreateHtmlFile(WebStr, "RedReg.html", 101)
        Shell "cmd.exe /c RedReg.html"
End If
                                 '      7   11  15  19      27
If Index = 1 Then               ''10010|2:4|3:3|321|2111100|420|3210
                                ' 10100|2111100|420|3210|2:4|3:3|
                                '
    For i = List1.ListCount - 44 To List1.ListCount - 1
        tempstr = Left(List1.List(i), 5) + Right(List1.List(i), 29)
        GetRegStr = Left(tempstr, 5) + "|" + Mid(tempstr, 19, 7) + Right(tempstr, 9) + Mid(tempstr, 6, 9)
        WebStr = WebStr + WriteRegTR(GetRegStr)
        Next i
        Call CreateHtmlFile(WebStr, "RedStar.html", 102)
        Shell "cmd.exe /c RedStar.html"
End If

If Index = 2 Then
For i = List1.ListCount - 44 To List1.ListCount - 1
 tempstr = Left(List1.List(i), 19)
 BlueStr = BlueStr + WriteBLueAll(tempstr)
Next i

Call CreateHtmlFile(BlueStr, "blue.html", 103)
 Shell "cmd.exe /c blue.html"

End If





End Sub

Private Sub Command13_Click()
 Dim i, j, k As Integer
 Dim TailAll As Integer
 Dim tempstr, LineStr, tailstr As String
 Dim Tail(9) As Integer
 Dim TailCount(9) As Integer
 Dim TailPre(9) As Integer
 Dim PreCount As Integer
 Dim TailReg(4) As Integer
 Dim TailRegStr As String
TailMode = True
List2.Clear
Picture1.Visible = False
List2.Width = List2_Width
List4.Visible = False
 
 
 For i = List1.ListCount - TAILLINES To List1.ListCount - 1
    tempstr = List1.List(i)
    For j = 1 To 6
     k = Val(Mid(tempstr, 4 + 2 * j, 2))
     Tail(k Mod 10) = Tail(k Mod 10) + 1
    Next j
    
      '尾号区间分布  0区间，1-3 ，4-6.7-9
     TailReg(0) = Tail(0)
     TailReg(1) = Tail(1) + Tail(2) + Tail(3)
     TailReg(2) = Tail(4) + Tail(5) + Tail(6)
     TailReg(3) = Tail(7) + Tail(8) + Tail(9)
     For j = 0 To 3
      TailRegStr = TailRegStr + " " + Format(TailReg(j), "0")
      TailReg(j) = 0
    Next j
    
    '尾号总数统计，及与上期相同数据统计
    
    For k = 0 To 9
     If Not (Tail(k) = 0) Then
      TailCount(k) = TailCount(k) + Tail(k)
      LineStr = LineStr + Str(k) + Str(Tail(k)) + " | "
      TailAll = TailAll + 1
      tailstr = tailstr + Format(Str(k), "0")
      If Not (TailPre(k) = 0) Then
        PreCount = PreCount + 1
      End If
      
     Else
      LineStr = LineStr + "    " + " | "
     End If
     
 
     
     
     TailPre(k) = Tail(k)
     Tail(k) = 0
    Next k
    List2.AddItem LineStr + Str(TailAll) + " | " + Str(PreCount) + " | " + Format(tailstr, "@@@@@@") + " | " + TailRegStr
    tailstr = ""
    LineStr = ""
    TailRegStr = ""
    TailAll = 0
    PreCount = 0
    
    
Next i

  For i = 0 To 9
   tailstr = tailstr + Format(Str(TailCount(i)), "@@@") + "  | "
   LineStr = LineStr + "   " + "  | "
   Next i
   List2.AddItem LineStr
   List2.AddItem tailstr

   List2.ListIndex = List2.ListCount - 1
   Text2.Text = ""
   Text2.SetFocus
End Sub

Private Sub Command2_Click(Index As Integer)
List2.Clear
List4_init
'Call Picture_Init
If Check5.Value = 1 Then
    Call ShowAddSum(Index)

    If Check1.Value = 1 Then
       Call Show10Appear(Index)
    Else
        Call ShowRegLost(Index)
    End If
End If



End Sub

Private Sub Check1Init()
Dim i As Integer
For i = 1 To 9
    Load OptionReg1(i)      'unload
    Load OptionReg2(i)
    OptionReg1(i).Top = OptionReg1(i - 1).Top + OptionReg1(0).Height
    OptionReg2(i).Top = OptionReg2(i - 1).Top + OptionReg2(0).Height
    OptionReg1(i).Left = OptionReg1(0).Left
    OptionReg2(i).Left = OptionReg2(0).Left
    
Next i


End Sub

Private Sub Command3_Click()
List2.Clear
Dim NumType(SumNum) As SumType
Dim OutStr(SumNum) As StrSum
Dim TheSumStr(SumNum) As String
Dim TheLostSumStr(SumNum) As String
Dim TheAppearStr(SumNum) As String
Dim TheReg3Str(SumNum) As String
Dim TheReg4Str(SumNum) As String
Call Picture_Init

Dim i As Integer
Dim tempstr As String
For i = 0 To List1.ListCount - 1
 tempstr = Replace(List1.List(i), " ", "")
' NumType(i).TheSum = Mid(tempstr, 28, 3)
' NumType(i).TheLostSum = Mid(tempstr, 45, 2)
' NumType(i).TheAppear = Replace(List3.List(i + 6), " ", "")
' NumType(i).TheReg4 = Right(tempstr, 4)
' NumType(i).TheReg3 = Left(Right(tempstr, 8), 3)
 TheSumStr(i) = Mid(tempstr, 21, 3)
 TheLostSumStr(i) = Mid(tempstr, 38, 2)

 TheReg3Str(i) = Left(Right(tempstr, 8), 3)
 TheReg4Str(i) = Right(tempstr, 4)
 Next i
 For i = 20 To List3.ListCount - 1
     TheAppearStr(i) = Replace(List3.List(i + 6), " ", "")
Next i
 
 If Option1(0).Value = True Then
     Call AddSum(TheSumStr(), OutStr())
    Call SumSort(OutStr())
    Call ShowSumSort(OutStr())
    Call ShowSumSortPic
End If
 If Option1(1).Value = True Then
     Call AddSum(TheLostSumStr(), OutStr())
    Call SumSort(OutStr())
    Call ShowSumSort(OutStr())
End If
 If Option1(2).Value = True Then
     Call AddSum(TheAppearStr(), OutStr())
    Call SumSort(OutStr())
    Call ShowSumSort(OutStr())
End If


End Sub
Private Sub ShowSumSortPic()
Dim tempstr As String
Dim i, j As Integer
Dim MMin1 As MaxMin
Dim MMin2 As MaxMin
Dim NClow As Integer

Picture1.DrawWidth = 2
For i = 0 To List2.ListCount - 1
    tempstr = List2.List(i)
    Num1(i) = Val(Left(tempstr, 3))
    Num2(i) = Val(Right(tempstr, 2))
  
Next i
MMin1 = GetMaxMin(Num1())
MMin2 = GetMaxMin(Num2())


Picture1.Scale (MMin1.min - 1, MMin2.min - 1)-(MMin1.max + 1, MMin2.max + 1)
Picture_Y = MMin2.max
For i = 0 To List2.ListCount - 1
tempstr = List2.List(i)
' Picture1.CurrentX = Val(Left(tempstr, 3))
' Picture1.CurrentY = MMin2.max - NClow
 'Picture1.Print Left(tempstr, 3)
 Picture1.Line (Val(Left(tempstr, 3)), MMin2.max)-(Val(Left(tempstr, 3)), MMin2.max - Val(Right(tempstr, 2))), vbGreen
 Picture1.Print Right(tempstr, 2)


Next i
Picture1.DrawWidth = 13
'Picture1.PSet (95, 8), vbRed
For i = 1 To 20
tempstr = Right(Left(List1.List(List1.ListCount - i), 24), 3)
    For j = 0 To List2.ListCount - 1
        If Val(tempstr) = Num1(j) Then
            Picture1.PSet (Num1(j), MMin2.max - Num2(j)), RGB((i \ 5) * 255, (i \ 10) * 255, (i \ 15) * 255)
            Picture1.Print Str(i)
            Exit For
        End If
        
    Next j
    Label6.Caption = Label6.Caption + Str(Num1(j))
    
Next i

End Sub

Private Sub List4_init()
List2.Width = List2_Width \ 5
List4.Left = List2.Left + List2.Width
List4.Top = List2.Top
List4.Height = List2.Height
List4.Width = List2_Width - List2.Width
List4.Visible = True
End Sub


Private Sub List4_Uinit()
List2.Width = List2_Width
'List2.Clear
List4.Visible = False
End Sub


Private Sub Picture_Init()
Picture1.Visible = True
Picture1.Cls
List2.Width = List2_Width \ 4
Picture1.Top = List2.Top + 30
Picture1.Left = List2.Left + List2.Width
Picture1.Width = List2_Width - List2.Width
Picture1.Height = List2.Height - 50
'Picture1.Scale (0, 0)-(100, 100)
Picture1.DrawWidth = 3
Label5.Top = Picture1.Top
Label5.Left = Picture1.Left + Picture1.Width
Label6.Top = Label5.Top + Label5.Height
Label6.Left = Label5.Left
Label6.Height = List2.Height - Label5.Height
Label5.Caption = ""
Label6.Caption = ""
End Sub
Private Sub ShowPicture1(Nums() As String)
Dim PScale As MaxMin
'PScale = GetMaxMin(Nums())



End Sub

Private Sub ShowSumSort(InSum() As StrSum)
Dim i As Integer
For i = 0 To SumNum
    
    If Not (InSum(i).StrChr = "") Then
        List2.AddItem InSum(i).StrChr + "  " + Format(Str(InSum(i).StrSum), "00")
    End If
Next i
End Sub
Private Function CheckAllData(IStr As String) As String
Dim i, j As Integer
Dim NumAttrib As DataAttrib
Dim NowLostSum, NowSum As Integer
Dim Small4, Small3 As Integer
Dim AppregStr As String
Dim NowlostStr As String
Dim AppNums As Integer
Dim LHFlag As Boolean       '连号标志
Dim YESStr As String        '包含数值
Dim NOTStr As String        '不包含数值
Dim ExistFlag As Boolean     '包含数值标志
Dim InLostStr As String     '包含遗漏次数
Dim OutLostStr As String    '不包含遗漏
Dim Reg6Str As String
Dim StrLostFlag As Boolean  '遗漏数值标志



Dim Lost03Num As Integer    '遗漏小于3
Dim Lost49Num  As Integer   '遗漏4-9
Dim Lost10Num As Integer    '遗漏大于等于10

Dim Lost03Flag As Boolean
Dim Lost49Flag As Boolean
Dim Lost10Flag As Boolean

Dim BigNums As Integer      '当前大数个数
Dim OddNums As Integer      '当前奇数个数


Dim AppYES As String   '包含区间
Dim AppNOT As String     '不包含区间    '
Dim AppFlag As Boolean '出现包含区间标志
Dim BeforeFlag As Boolean   '以前出现标记
Dim BigFlag As Boolean      '大数
Dim OddFlag As Boolean      '奇数
Dim AppNumFlag As Boolean  '分布区间次数标志


'***和值 范围  ，遗漏和值范围
SumMin = Val(Text1(0).Text)
SumMax = Val(Text1(1).Text)
LostSumMin = Val(Text1(2).Text)
LostSumMax = Val(Text1(3).Text)

'****依次包含数字 ，不包含数字  ，包含遗漏  ，包含区间
YESStr = Replace(Text3(0).Text, " ", "")
NOTStr = Replace(Text3(1).Text, " ", "")

InLostStr = Replace(Text4(0).Text, " ", "")
OutLostStr = Replace(Text4(1).Text, " ", "")

AppYES = Replace(Text5(0).Text, " ", "")
AppNOT = Replace(Text5(1).Text, " ", "")

'*******************************

'** 得到遗漏区间分布
For i = 0 To 2
 If Option4(i).Value = True Then
  Lost03Num = i + 2
 End If
 If Option5(i).Value = True Then
  Lost49Num = i
 End If
 If Option6(i).Value = True Then
  Lost10Num = i
 End If
 Next i
 If Option4(3).Value = True Then
  Lost03Num = 5
 End If
 If Option5(3).Value = True Then
    Lost49Num = 3
 End If
 
 '********************************888
 Call CheckAttrib(IStr, NumAttrib)          ' 得到 当前数字的大小 奇数 偶数特性
        
         NowSum = CheckNowSum(IStr)                        '得到和值大小
        NowLostSum = CheckNowLost(IStr, NowlostStr)        '函数返回 遗漏次数和值   。。 NowLoststr 得到 遗漏字符
        
      '**********  遗漏为0-3
        If Lost03Num = CheckSmall(IStr, 0, 3) Then
            Lost03Flag = True
        Else
            Lost03Flag = False
        End If
        '************************* 遗漏4-9
        If Lost49Num = CheckSmall(IStr, 4, 9) Then
            Lost49Flag = True
        Else
            Lost49Flag = False
        End If
        '***********************遗漏次数10 -50
        If Lost10Num = CheckSmall(IStr, 10, 50) Then
            Lost10Flag = True
        Else
            Lost10Flag = False
        End If
        
        '********************************
               
       AppNumFlag = True
    LHFlag = True
     BeforeFlag = True
     
     If Check8.Value = 0 Then
        Lost03Flag = True
     End If
     If Check9.Value = 0 Then
        Lost49Flag = True
     End If
     If Check10.Value = 0 Then
        Lost10Flag = True
     End If
     
     
    
   '**********************************************************
       '连号统计
     '   LHFlag = False
     '   For j = 0 To 2
     '   If Option3(j).Value And (CheckLH(IStr) = j) Then
     '       LHFlag = True
     '       Exit For
     '   End If
     '   Next j
       '检验连号的次数
    '***************************************************
       '检查之前是否出现
      
'       If Check2.Value = 1 Then
'         BeforeFlag = CheckBefore(IStr)
 '      Else
  '        BeforeFlag = True
  '     End If
     '***************************
      BigNums = CheckBSNum(IStr)
      If BigCheck.Value = 1 Then              '大小判断
        If NowBigNums = BigNums Then
            BigFlag = True
        Else
            BigFlag = False
        End If
    Else
        If Check3.Value = 0 Then
            If BigNums > 3 Then
              BigFlag = True
            Else
              BigFlag = False
            End If
        Else
            
             If BigNums < 3 Then
              BigFlag = True
            Else
              BigFlag = False
            End If
       
        End If
         'If BigNums >= 1 And BigNums <= 5 Then
          '    BigFlag = True
         'Else
         '   BigFlag = False
        'End If
    End If

    
    
    '********************************
       
       
    '
    OddNums = CheckOENum(IStr)
    If OddCheck.Value = 1 Then              '奇偶数判断
        If OddNums = NowOddNums Then
            OddFlag = True
        Else
            OddFlag = False
        End If
    Else
           If Check4.Value = 0 Then
            If OddNums > 3 Then
              OddFlag = True
            Else
              OddFlag = False
            End If
        Else
            
             If OddNums < 3 Then
              OddFlag = True
            Else
              OddFlag = False
            End If
       
        End If
    End If
      If BigC.Value = 0 Then
        BigFlag = True
       End If
       If OddC.Value = 0 Then
        OddFlag = True
      End If
   
        
        
        Reg6Str = CheckReg6(IStr)
       ExistFlag = CheckExist(IStr, YESStr, NOTStr)       '检查 存在的数字及不存在的数字
       StrLostFlag = CheckLostFlag(IStr, InLostStr, OutLostStr)   '检查存在的遗漏及不存在的遗漏
       AppFlag = CheckAppFlag(IStr, AppYES, AppNOT)            '检查出现的区间和不出现的区间
       
    If OddFlag And BigFlag _
        And NowLostSum >= LostSumMin And NowLostSum <= LostSumMax _
        And NowSum >= SumMin And NowSum <= SumMax _
        And Lost03Flag And Lost49Flag And Lost10Flag _
        And LHFlag And ExistFlag And StrLostFlag And AppFlag And BeforeFlag And AppNumFlag Then
        
      ' List2.AddItem Istr + " " + NumAttrib.BSAttrib + " " + NumAttrib.OEAttrib + " " + NowlostStr + " " + Str(NowLostSum) + " " + Format(Str(NowSum), "000") + " " + AppregStr
        CheckAllData = NumAttrib.BSAttrib + " " + NumAttrib.OEAttrib + " " + NowlostStr + " " + Str(NowLostSum) + " " + Format(Str(NowSum), "000") + " " + AppregStr + "  " + Reg6Str
    Else
        CheckAllData = ""
     End If
        


End Function



Private Sub Command4_Click()
Dim i As Integer
Dim LineStr As String
List2.Clear

Label9.Caption = "计算中"
Picture1.Visible = False
List2.Width = List2_Width
List4.Visible = False
If RegCheck.Value = 1 Then
Dim OutStr(10000) As String
'NowRegion 7区区间分布排列 0121200  -- 330
Call GetAllCombin(NowRegion, OutStr())
For i = 0 To 10000
    If Not (OutStr(i) = "") Then
         
     LineStr = CheckAllData(OutStr(i))
      If Len(LineStr) > 0 Then
        List2.AddItem OutStr(i) + " " + LineStr
      End If
       
    End If
Next i

End If



Label9.Caption = "完成"
Label1.Caption = "Total = " + Str(List2.ListCount)
End Sub

Private Function CheckBefore(IStr As String) As Boolean
Dim i As Integer
Dim tempstr As String
For i = 0 To List1.ListCount - 1
    tempstr = Right(Left(List1.List(i), 17), 12)
    If IStr = tempstr Then
        CheckBefore = False
        Exit Function
    End If
Next i
    CheckBefore = True



End Function
Private Function CheckLH(IStr As String) As Integer
Dim i, LH As Integer
For i = 2 To 6
  If (Val(Mid(IStr, 2 * i - 1, 2)) - Val(Mid(IStr, 2 * (i - 1) - 1, 2))) = 1 Then
    LH = LH + 1
  End If
Next i
CheckLH = LH

End Function
Private Function CheckAppFlag(IStr As String, YESStr As String, NOTStr As String) As Boolean
Dim TheAppReg(1 To 6) As Integer
Dim i, j As Integer
Dim Flag1 As Boolean
Dim Flag2 As Boolean
Dim AppFNum As Integer
For i = 1 To 6
   TheAppReg(i) = NewSumRegion(Val(Mid(IStr, 2 * i - 1, 2)))
Next i

For i = 1 To Len(YESStr)
    For j = 1 To 6
        If Val(Mid(YESStr, i, 1)) = TheAppReg(j) Then
            AppFNum = AppFNum + 1
            Exit For
        End If
    Next j
Next i
If Len(YESStr) = AppFNum Then
    Flag1 = True
Else
    Flag1 = False
End If

Flag2 = True
For i = 1 To Len(NOTStr)
    For j = 1 To 6
        If Val(Mid(NOTStr, i, 1)) = TheAppReg(j) Then
            Flag2 = False
            Exit For
        End If
    Next j
    If Flag2 = False Then
        Exit For
    End If
    
Next i

If Len(YESStr) = 0 Then
 Flag1 = True
End If
If Len(NOTStr) = 0 Then
    Flag2 = True
End If

If Flag1 And Flag2 Then
    CheckAppFlag = True
Else
    CheckAppFlag = False
End If




End Function
Private Function CheckLostFlag(IStr As String, YESStr As String, NOTStr As String) As Boolean
    Dim i, j As Integer
  Dim Flag1, Flag2 As Boolean
    Dim InLostNum As Integer

For i = 1 To Len(YESStr) \ 2
    For j = 1 To 6
         If LoseR(TotalNum, Val(Mid(IStr, 2 * j - 1, 2))) = Val(Mid(YESStr, 2 * i - 1, 2)) Then
            InLostNum = InLostNum + 1
            Exit For
         End If
    Next j
Next i
    If InLostNum = Len(YESStr) \ 2 Then
        Flag1 = True
    Else
        Flag1 = False
    End If
Flag2 = True
For i = 1 To Len(NOTStr) \ 2
    For j = 1 To 6
         If LoseR(TotalNum, Val(Mid(IStr, 2 * j - 1, 2))) = Val(Mid(NOTStr, 2 * i - 1, 2)) Then
            Flag2 = False
            Exit For
         End If
    Next j
    If Flag2 = False Then
        Exit For
    End If
Next i

If Len(YESStr) = 0 Then
 Flag1 = True
End If
If Len(NOTStr) = 0 Then
    Flag2 = True
End If

If Flag1 And Flag2 Then
    CheckLostFlag = True
Else
    CheckLostFlag = False
End If



End Function
Private Function CheckExist(IStr As String, YESStr As String, NOTStr As String) As Boolean
Dim i, j As Integer
Dim Flag1, Flag2 As Boolean

Dim ExistNum As Integer

If Len(YESStr) = 0 And Len(NOTStr) = 0 Then
    CheckExist = True
    Exit Function
End If

For i = 1 To Len(YESStr) \ 2
    For j = 1 To 6
        If Mid(IStr, 2 * j - 1, 2) = Mid(YESStr, 2 * i - 1, 2) Then
         ExistNum = ExistNum + 1
         Exit For
        End If
    Next j
Next i
If ExistNum = (Len(YESStr) \ 2) Then
    Flag1 = True
Else
    Flag1 = False
End If



Flag2 = True
For i = 1 To Len(NOTStr) \ 2
    For j = 1 To 6
        If Mid(IStr, 2 * j - 1, 2) = Mid(NOTStr, 2 * i - 1, 2) Then
            Flag2 = False
        End If
    Next j
Next i
If Flag1 And Flag2 Then
    CheckExist = True
Else
    CheckExist = False
End If






End Function

Private Function CheckAppReg(IStr As String, OStr As String) As Integer
  Dim i, Sum As Integer
  Dim NumSum(8) As Integer
  OStr = ""
 For i = 1 To 6
    NumSum(NewSumRegion(Val(Mid(IStr, 2 * i - 1, 2)))) = 1
    OStr = OStr + Format(Str(NewSumRegion(Val(Mid(IStr, 2 * i - 1, 2)))), "0") + " "
Next i
  For i = 0 To 8
   If NumSum(i) = 1 Then
    CheckAppReg = CheckAppReg + 1
  End If
 Next i
 
  
  
  'NewSumRegion (i)
End Function
Private Function CheckNowSum(IStr As String) As Integer
    Dim i, Sum As Integer
 For i = 1 To 6
    Sum = Sum + Val(Mid(IStr, 2 * i - 1, 2))
Next i
CheckNowSum = Sum
 

End Function
Private Function CheckNowLost(IStr As String, LostStr As String) As Integer
 Dim i, Sum As Integer
 LostStr = ""
 For i = 1 To 6
    LostStr = LostStr + Format(Str(LoseR(TotalNum, Val(Mid(IStr, 2 * i - 1, 2)))), "00") + " "
    Sum = Sum + LoseR(TotalNum, Val(Mid(IStr, 2 * i - 1, 2)))
Next i
 CheckNowLost = Sum

End Function

Private Function CheckSmall(IStr As String, RS As Integer, RB As Integer) As Integer
Dim i, Nums As Integer
Dim NowLost As Integer
For i = 1 To 6
    NowLost = LoseR(TotalNum, Val(Mid(IStr, 2 * i - 1, 2)))
    If NowLost >= RS And NowLost <= RB Then
        Nums = Nums + 1
    End If
Next i
CheckSmall = Nums
End Function
Private Sub ShowList3()
Dim i, j As Integer
Dim tempstr As String
Dim LineStr As String
Dim SumRegionStr As String
Dim NumRegion(1 To 33) As Integer
For i = 0 To List1.ListCount - 1
    tempstr = Right(Left(List1.List(i), 17), 12)
 For j = 1 To 6
    NumRegion(Val(Mid(tempstr, 2 * j - 1, 2))) = NumRegion(Val(Mid(tempstr, 2 * j - 1, 2))) + 1
    LineStr = LineStr + Str(NumRegion(Val(Mid(tempstr, 2 * j - 1, 2))))
 Next j
 If i < 8 Then
    List3.AddItem LineStr
    LineStr = ""
Else
    LineStr = ""
    Call GetSumRegion(NumRegion(), SumRegion())        ''  计算区间分布

    For j = 1 To 6
        LineStr = LineStr + Str(SumRegion(Val(Mid(tempstr, 2 * j - 1, 2))))
    Next j
    List3.AddItem LineStr
 End If
 Next i
 List3.ListIndex = List3.ListCount - 1


End Sub

Private Sub ShowSumRegion(SumRegion() As Integer)
    Dim i, j As Integer
    Dim LineStr As String
    
For i = 1 To SumRegNums
        LineStr = Str(i) + "|"
        For j = 1 To 33
            If SumRegion(j) = i Then
              LineStr = LineStr + Format(Str(j), "00") + " "
            End If
        Next j
        List3.AddItem LineStr
        LineStr = ""
Next i
End Sub
Private Sub ShowCommand5()
'Form2.Show
List3.Clear
Dim tempstr As String
Dim LineStr As String
Dim Str1, Str2 As String
Dim i, j As Integer
Dim once As Boolean
Call ShowSumRegion(SumRegion())

For i = List1.ListCount - 1 To 0 Step -1
    tempstr = Right(Left(List1.List(i), 17), 12)
    For j = 1 To 6
        AllSum(Val(Mid(tempstr, 2 * j - 1, 2))) = AllSum(Val(Mid(tempstr, 2 * j - 1, 2))) - 1
    Next j
        Call GetSumRegion(AllSum(), SumRegion())
      '  Call ShowSumRegion(SumRegion())                     ''显示区间分布
    
    For j = 1 To 6
        LineStr = LineStr + Str(SumRegion(Val(Mid(tempstr, 2 * j - 1, 2))))
    Next j
    List3.AddItem LineStr
    LineStr = ""
Next i
' For i = 1 To 33
 '  Str1 = Str1 + Str(SumRegion(i))
 '  Str2 = Str2 + Str(AllSum(i))
 ' Next i
 ' List2.AddItem Str1
  'List2.AddItem Str2
   ' SumRegion(1) = SumRegion(2) + SumRegion(3) + AllSum(1) + AllSum(2) + AllSum(3) + AllSum(4)
End Sub

Private Sub Command5_Click()
Dim i, j, k, L As Integer
Dim PrePos As Integer
Dim tempstr As String
Dim LineStr As String
Dim LineStr2 As String
Dim Nums As Integer
Dim AppStr(12) As String
Dim Sum10(7) As Integer
Dim App6(6) As Integer
AppStr(0) = "330"
AppStr(1) = "240"
AppStr(2) = "420"
AppStr(3) = "321"
AppStr(4) = "231"
AppStr(5) = "141"
AppStr(6) = "411"
AppStr(7) = "510"
AppStr(8) = "150"
AppStr(9) = "222"
AppStr(10) = "312"
AppStr(11) = "132"
List2.Clear
Picture1.Visible = False
List2.Width = List2_Width
List4.Visible = False

If Check1.Value = 1 Then
    For i = 0 To List1.ListCount - 1
      k = k + 1
      tempstr = Left(Right(List1.List(i), 16), 7)
      For j = 1 To 7
        Sum10(j) = Sum10(j) + Val(Mid(tempstr, j, 1))
      Next j
      tempstr = Replace(List3.List(i), " ", "")
      For j = 1 To 6
        App6(Val(Mid(tempstr, j, 1))) = App6(Val(Mid(tempstr, j, 1))) + 1
        
      Next j
      
     If k = 10 Then
        For j = 1 To 7
            LineStr = LineStr + " | " + Format(Str(Sum10(j)), "00") + " "
            Sum10(j) = 0
        Next j
        For j = 1 To 6
            LineStr2 = LineStr2 + " | " + Format(Str(App6(j)), "00") + " "
            App6(j) = 0
        Next j
        
        
        List2.AddItem LineStr + "    |*|    " + LineStr2
        k = 0
        LineStr = ""
        LineStr2 = ""
     End If
    Next i

    
    
    
    
    For j = 1 To 7
            LineStr = LineStr + " | " + Format(Str(Sum10(j)), "00") + " "
    Next j
    
      For j = 1 To 6
            LineStr2 = LineStr2 + " | " + Format(Str(App6(j)), "00") + " "
            App6(j) = 0
        Next j
    
    
    
    
    
    
    List2.AddItem LineStr + "    |*|    " + LineStr2
    List2.AddItem Str(k) + "  期数据剩下"
    List2.ListIndex = List2.ListCount - 1
    Exit Sub
End If



Dim TheReg3Str(SumNum) As String
Nums = List1.ListCount - 1


For i = 0 To Nums
    tempstr = Replace(List1.List(i), " ", "")
    TheReg3Str(i) = Left(Right(tempstr, 8), 3)
Next i
'List2.AddItem GetLost(TheReg3Str(), "330")

For i = Nums To Nums - 100 Step -1
 For j = i - 1 To 0 Step -1
   If TheReg3Str(j) = TheReg3Str(i) Then
        LineStr = LineStr + TheReg3Str(i) + "-" + Format(Str(i - j - 1), "@@@") + "  "
        Exit For
    End If
Next j
If Len(LineStr) > 80 Then
    List2.AddItem LineStr
    LineStr = ""
End If
Next i

List2.AddItem "          "
For j = 0 To 11
   LineStr2 = AppStr(j) + "- "
   k = 0
    For i = Nums To 1 Step -1
         k = k + 1
        If TheReg3Str(i) = AppStr(j) Then
            LineStr2 = LineStr2 + Format(Str(k - 1), "00") + " "
            k = 0
        End If
       
     Next i
     List2.AddItem LineStr2
        LineStr2 = ""
Next j
'Open "C:\123.txt " For Append As #1
    ' Write #1, "cc"
'        Print #1, LineStr
'        Close #1


End Sub

Private Sub Command6_Click()
Dim IStr As String
Dim i As Integer
Dim tempstr As String
Dim LineStr As String
Dim LineStr2 As String
Dim LClen As Integer
Dim Allm As Integer
Call List4_Uinit
List2.Clear
LClen = Len(Label10.Caption) \ 2


If LClen >= 6 Then
    tempstr = Replace(Label10.Caption, " ", "")
    IStr = "111111"
    For i = 7 To Len(tempstr) \ 2
        IStr = IStr + "0"
    Next i
    Allm = CCInOut(LClen, 6)
  For i = 1 To Allm

   LineStr = CombinNum(tempstr, IStr)
 '  Call CheckAllData(LineStr)
       LineStr2 = CheckAllData(LineStr)
      If Len(LineStr2) > 0 Then
        List2.AddItem LineStr + " " + LineStr2
      End If
   
   IStr = ChangeStr(IStr)
 '  List2.AddItem LineStr
 Next i



End If



Label1.Caption = "Total= " + Str(List2.ListCount)

End Sub

Private Sub Command7_Click()
Dim i As Integer
Dim TempstrB As String
Dim tempstr As String
Dim LHnow, LHBefore As Integer
Dim THnow As Integer
List2.Clear
Call List4_Uinit
For i = List1.ListCount - 100 To List1.ListCount - 1
 tempstr = Right(Left(List1.List(i), 17), 12)
 TempstrB = Right(Left(List1.List(i - 1), 17), 12)

LHnow = CheckLH(tempstr)
LHBefore = CheckLHB(TempstrB, tempstr)
THnow = CheckTHB(TempstrB, tempstr)
List2.AddItem tempstr + "  | " + Str(LHnow) + "  |  " + Str(LHBefore) + "   |  " + Str(THnow) + "   |  "
Next i
List2.AddItem "             本期相邻|上期相邻|上期相同 "
List2.ListIndex = List2.ListCount - 1
'020309242627 081013141623
'Call ShowNowCheck(List1.ListIndex)


End Sub
Public Function CheckLHB(TempstrB As String, TempstrN As String) As Integer
Dim i, j As Integer
Dim LHCount As Integer
For i = 1 To 6
 For j = 1 To 6
  If Abs(Val(Mid(TempstrN, 2 * i - 1, 2)) - Val(Mid(TempstrB, 2 * j - 1, 2))) = 1 Then
     LHCount = LHCount + 1
     Exit For
  End If
  Next j
Next i
 CheckLHB = LHCount
End Function
Public Function CheckTHB(TempstrB As String, TempstrN As String) As Integer
Dim i, j As Integer
Dim THCount As Integer
For i = 1 To 6
 For j = 1 To 6
  If Abs(Val(Mid(TempstrN, 2 * i - 1, 2)) - Val(Mid(TempstrB, 2 * j - 1, 2))) = 0 Then
     THCount = THCount + 1
     Exit For
  End If
  Next j
Next i
 CheckTHB = THCount
End Function


Private Sub Command8_Click()
Dim i, j As Integer
Dim LineStr As String
Dim LHSum(32) As StrSum
Dim NumArr(1000) As String
Dim OutStr(20) As String
List2.Clear
For i = 0 To List1.ListCount - 1
    NumArr(i) = Right(Left(List1.List(i), 17), 12)
Next i
For i = 1 To 32
    LHSum(i).StrChr = Format(Str(i), "00") + Format(Str(i + 1), "00")
    LHSum(i).StrSum = FindLHSum(NumArr(), LHSum(i).StrChr)
    LineStr = LineStr + LHSum(i).StrChr + "  " + Format(Str(LHSum(i).StrSum), "00") + " |"
    j = j + 1
    If j = 8 Then
        j = 0
        List2.AddItem LineStr
        LineStr = ""
     End If

Next i
    List2.AddItem LineStr
    LineStr = ""
    j = 0
    List2.AddItem "*****************************************"
    Call SumSort(LHSum())
For i = 1 To 32
    LineStr = LineStr + LHSum(i).StrChr + "  " + Format(Str(LHSum(i).StrSum), "00") + " |"
    j = j + 1
    If j = 8 Then
        j = 0
        List2.AddItem LineStr
        LineStr = ""
     End If
Next i
    
    List2.AddItem LineStr
   ' Call ShowLast10(List1.ListCount - 1)
    For i = 0 To 9
 NumArr(i) = Right(Left(List1.List(List1.ListCount - 1 - i), 17), 12)
Next i

Call TheMainValue(NumArr(), OutStr())


For i = 0 To 4
 If Not OutStr(i) = "" Then
     List2.AddItem "5期内出现" + Str(i) + "次   " + OutStr(i)
 End If
Next i
List2.AddItem "************************************"
For i = 5 To 20
 If Not OutStr(i) = "" Then
     List2.AddItem "10期内出现" + Str(i - 5) + "次  " + OutStr(i)
 End If
Next i


End Sub
Private Sub ShowLast10(Lindex As Integer)
Dim i As Integer
Dim NumArr(10) As String
Dim OutStr(20) As String
For i = 0 To 9
 NumArr(i) = Right(Left(List1.List(Lindex - i), 17), 12)
Next i

Call TheMainValue(NumArr(), OutStr())


For i = 0 To 4
 If Not OutStr(i) = "" Then
     List4.AddItem "5期内出现" + Str(i) + "次   " + OutStr(i)
 End If
Next i
List2.AddItem "************************************"
For i = 5 To 20
 If Not OutStr(i) = "" Then
     List4.AddItem "10期内出现" + Str(i - 5) + "次  " + OutStr(i)
 End If
Next i
End Sub

Private Sub Command9_Click()
Dim i, j, k As Integer
Call List4_Uinit
List2.Clear
Dim OddNums, BigNums As Integer
Dim LineStr1, LineStr2 As String
Dim tempstr As String
For i = 0 To List1.ListCount - 1
    tempstr = Right(Left(List1.List(i), 17), 12)
    OddNums = OddNums + CheckOENum(tempstr)
    BigNums = BigNums + CheckBSNum(tempstr)
    j = j + 1
    If j = 5 Then
        LineStr1 = "奇偶比 " + Str(OddNums) + "  : " + Str(j * 6 - OddNums)
        LineStr2 = "大小比 " + Str(BigNums) + "  : " + Str(j * 6 - BigNums)
        List2.AddItem LineStr1 + "  |  " + LineStr2
        LineStr1 = ""
        LineStr2 = ""
        j = 0
        OddNums = 0
        BigNums = 0
    End If
 Next i
        LineStr1 = "奇偶比 " + Str(OddNums) + "  : " + Str(j * 6 - OddNums)
        LineStr2 = "大小比 " + Str(BigNums) + "  : " + Str(j * 6 - BigNums)
        List2.AddItem LineStr1 + "  |  " + LineStr2


End Sub

Private Sub Form_Load()
'Dim TheAllSum(1 To 33) As StrSum
Dim i As Integer
Dim tempstr As String
List2_Width = List2.Width
Open App.Path & "\golddata.txt" For Input As #1
While Not EOF(1)
i = i + 1
Line Input #1, Data(i)
Data(i) = Replace(Data(i), " ", "")
If Len(Data(i)) > 0 Then
    Call AddAllSum(Left(Right(Data(i), 14), 12), AllSum())
End If


Wend

Close #1
TotalNum = i

CurrentNum = Left(Data(TotalNum), 5)
Call Check1Init
Call CheckStrInit
'Call GetSumRegion(AllSum(), SumRegion())        ''  计算区间分布
'Call ShowSumRegion(SumRegion())                     ''显示区间分布




Call CheckLost
Call ShowList1
Command2_Click (0)
Call Combo1Init
For i = 1 To 33
    List2.AddItem Str(i) + " " + Str(AllSum(i))
  '  TheAllSum(i).StrChr = Str(i)
 '   TheAllSum(i).StrSum = AllSum(i)
Next i
List2.AddItem tempstr
'tempstr = Right(Left(List1.List(List1.ListCount - 1), 17), 12)
'Dim nums As Integer
'Dim maxstr As String
'Call ShowSumRegion(SumRegion())
'Call CheckAppearRegions(tempstr, SumRegion(), nums, maxstr)    'check the appears
'Label1.Caption = Str(nums) + " " + maxstr

Label1.Caption = Label1.Caption + Str(TotalNum)
'Call ShowList2
'Call ShowCommand5
Call ShowList3
For i = 1 To 33
    NewSumRegion(i) = SumRegion(i)
 
Next i
'Command5_Click
Call Command1_Click

Text2.Text = ""
End Sub
Private Sub CheckLost()
   Dim i, j As Integer
   Dim tempstr As String
   Dim LineStr As String
   
   For i = 1 To TotalNum
    tempstr = Right(Replace(Data(i), " ", ""), 14)
    For j = 1 To 6
        LoseRFlag(i, Val(Mid(tempstr, 2 * j - 1, 2))) = 1
    Next j
        LoseBFlag(i, Val(Right(tempstr, 2))) = 1
   Next i
   For j = 1 To 33
    If LoseRFlag(1, j) = 1 Then
        LoseR(1, j) = 0
    Else
        LoseR(1, j) = 1     'Lose One
    End If
   Next j
   For j = 1 To 16
    If LoseBFlag(1, j) = 1 Then
        LoseB(1, j) = 0
    Else
        LoseB(1, j) = 1
     End If
    Next j
   
  For i = 2 To TotalNum
    For j = 1 To 33
        If LoseRFlag(i, j) = 1 Then
            LoseR(i, j) = 0
        Else
            LoseR(i, j) = LoseR(i - 1, j) + 1
        End If
    Next j
    For j = 1 To 16
        If LoseBFlag(i, j) = 1 Then
            LoseB(i, j) = 0
        Else
            LoseB(i, j) = LoseB(i - 1, j) + 1
        End If
   
    
    Next j
   
  Next i

End Sub
Private Sub ShowList1()
 Dim tempstr As String
 Dim LineStr, BlueStr, BSStr, OEStr, RegionStr1, RegionStr2 As String
 Dim i, j As Integer
 Dim TempLostSum, TempSum As Integer
For i = 2 To TotalNum
  tempstr = Right(Replace(Data(i), " ", ""), 14)
   For j = 1 To 6
    TempSum = TempSum + Val(Mid(tempstr, 2 * j - 1, 2))
   LineStr = LineStr + Format(LoseR(i - 1, Val(Mid(tempstr, 2 * j - 1, 2))), "00") + " "
   TempLostSum = TempLostSum + LoseR(i - 1, Val(Mid(tempstr, 2 * j - 1, 2)))
   Next j
   LineStr = LineStr + "= " + Format(Str(TempLostSum), "00")
    BlueStr = Format(Str(LoseB(i - 1, Val(Right(tempstr, 2)))), "00")
    BSStr = CheckBS(tempstr)
    OEStr = CheckOE(tempstr)
    RegionStr1 = CheckR11(tempstr) + "|" + CheckR6(tempstr) + "|" + CheckR3(tempstr) + "|" + CheckR4(tempstr)
    List1.AddItem RTrim(Data(i)) + "= " + Format(Str(TempSum), "000") + " |" + LineStr + " *" + BlueStr + "|" + BSStr + "|" + OEStr + "|" + RegionStr1
    LineStr = ""
    TempLostSum = 0
    TempSum = 0

 Next i
 List1.ListIndex = List1.ListCount - 1
    
End Sub

Private Function ShowRBLost()
    Dim i, j, k, L As Integer
    Dim LineStrR As String
    Dim LineStrB As String
    Dim LineStr(33) As String

  For j = 1 To 33
    If LoseR(TotalNum, j) = 0 Then
        LineStrR = Format(Str(LoseR(TotalNum - 1, j)), "00")
    Else
    
    LineStrR = Format(Str(LoseR(TotalNum, j)), "00")
    End If
       LoseStr(j) = LineStrR
    
   For i = TotalNum - 1 To 1 Step -1
    If LoseR(i, j) = 0 Then
        LineStrR = Format(Str(LoseR(i - 1, j)), "00") + "|" + LineStrR
        k = k + 1
    End If
    If k > 10 Then
        k = 0
        Exit For
    End If
    
        
   Next i
       ' List2.AddItem Format(Str(j), "00") + "|" + LineStrR
       LineStr(j) = "R:" + Format(Str(j), "00") + "|" + LineStrR + "|"
        LineStrR = ""
  
   Next j
    'List2.AddItem "------------------------------------"

   For j = 1 To 16
   
       LineStrB = Format(Str(LoseB(TotalNum, j)), "00")
    For i = TotalNum To 1 Step -1
        If LoseB(i, j) = 0 Then
            LineStrB = Format(Str(LoseB(i - 1, j)), "00") + "|" + LineStrB
            k = k + 1
        End If
    If k > 10 Then
        k = 0
        Exit For
    End If
        
    Next i
       ' List2.AddItem Format(Str(j), "00") + "|" + LineStrB
        LineStr(j) = LineStr(j) + "   B:" + Format(Str(j), "00") + "|" + LineStrB
        LineStrB = ""
   Next j
   For i = 1 To 33
    List2.AddItem LineStr(i)
    Next i
    

   
End Function

Private Sub ShowAddSum(cmdindex As Integer)
   Dim TSum As Integer
    TSum = 300          ''leaf
   Dim Temp3str(SumNum) As String
   Dim Temp4str(SumNum) As String
   Dim Temp6Str1(SumNum) As String
   Dim Temp6Str2(SumNum) As String
   Dim Temp11Str(SumNum) As String
   
   Dim Temp3d(SumNum) As StrSum
   Dim Temp4d(SumNum) As StrSum
   Dim Temp6d1(SumNum) As StrSum
   Dim Temp6d2(SumNum) As StrSum
   Dim Temp11d(SumNum) As StrSum
   Dim tempstr As String
 '   If Not (Text2.Text = "") And Val(Text2.Text) <= 100 Then
  '      TSum = Val(Text2.Text)
   ' Else
    '    TSum = SumNum
    'End If

    
    Dim i As Integer
    For i = TSum To 1 Step -1
        tempstr = Right(Replace(List1.List(List1.ListCount - i), " ", ""), 20)
        Temp6Str1(i) = Left(Right(tempstr, 16), 3)
        Temp6Str2(i) = Left(Right(tempstr, 13), 3)
        Temp3str(i) = Left(Right(tempstr, 8), 3)
        Temp4str(i) = Left(Right(tempstr, 4), 4)
        Temp11Str(i) = Left(tempstr, 3)
    Next i
   ' List2.AddItem GetLost(Temp3str(), "330")
 
    


 

    
    If (cmdindex = 0) Then
       Call AddSum(Temp3str(), Temp3d())
    Call SumSort(Temp3d())
    For i = 0 To SumNum
  
        tempstr = ""
        If Not (Temp3d(i).StrChr = "") Then
            List2.AddItem Temp3d(i).StrChr + "  " + Format(Str(Temp3d(i).StrSum), "00") + "  "
         End If
     Next i
'     List2.AddItem "Region3"
    End If
    If (cmdindex = 1) Then
    Call AddSum(Temp4str(), Temp4d())
    Call SumSort(Temp4d())
    For i = 1 To SumNum
        tempstr = ""
  
        If Not (Temp4d(i).StrChr = "") Then
            List2.AddItem Temp4d(i).StrChr + "  " + Format(Str(Temp4d(i).StrSum), "00") + "  "
         End If
     Next i
 '    List2.AddItem "Region4"
    End If
    If (cmdindex = 2) Then      ''6-1
        Call AddSum(Temp6Str1(), Temp6d1())
    Call SumSort(Temp6d1())
    For i = 1 To SumNum
        tempstr = ""
        If Not (Temp6d1(i).StrChr = "") Then
            List2.AddItem Temp6d1(i).StrChr + "  " + Format(Str(Temp6d1(i).StrSum), "00") + "  "
         End If
     Next i
  '   List2.AddItem "Region61"
    End If
        If (cmdindex = 3) Then  '6-2
        Call AddSum(Temp6Str2(), Temp6d2())
    Call SumSort(Temp6d2())
    For i = 1 To SumNum
        tempstr = ""
        If Not (Temp6d2(i).StrChr = "") Then
            List2.AddItem Temp6d2(i).StrChr + "  " + Format(Str(Temp6d2(i).StrSum), "00") + "  "
         End If
     Next i
   '  List2.AddItem "Region62"
    End If
    If (cmdindex = 4) Then
        Call AddSum(Temp11Str(), Temp11d())
     Call SumSort(Temp11d())
    For i = 1 To SumNum
        tempstr = ""
        If Not (Temp11d(i).StrChr = "") Then
            List2.AddItem Temp11d(i).StrChr + "  " + Format(Str(Temp11d(i).StrSum), "00") + "  "
         End If
     Next i
    ' List2.AddItem "Region11"
    
    End If
End Sub

Private Sub Frame12_DblClick()
Call CheckAllData("081112141822")
End Sub

Private Sub Combo1Init()
Dim i As Integer
For i = 0 To List2.ListCount - 2
    Combo1.AddItem List2.List(i)
Next i
 Combo1.ListIndex = 0
End Sub

Private Sub Label1_Click()
Dim i As Integer
Open "List1.txt" For Append As #1
    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
    Next i
Close #1

Open "List2.txt" For Append As #1
    For i = 0 To List2.ListCount - 1
        Print #1, List2.List(i)
    Next i
Close #1

End Sub

Private Sub Label8_Click()
Form3.Show
End Sub


Private Sub CheckLost3(Lost3() As Integer, tempstr As String)

Dim i, j As Integer
For i = LBound(Lost3) To UBound(Lost3)
    Lost3(i) = 0
Next i

For j = 1 To 6
  If Val(Mid(tempstr, 2 * j - 1, 2)) < 4 Then
      Lost3(Val(Mid(tempstr, 2 * j - 1, 2))) = Lost3(Val(Mid(tempstr, 2 * j - 1, 2))) + 1
  End If
   If Val(Mid(tempstr, 2 * j - 1, 2)) > 3 And Val(Mid(tempstr, 2 * j - 1, 2)) < 10 Then
    Lost3(4) = Lost3(4) + 1
   End If
If Val(Mid(tempstr, 2 * j - 1, 2)) > 9 Then
    Lost3(5) = Lost3(5) + 1
   End If
Next j


End Sub

Private Sub ShowLost3(Lost3() As Integer)
List2.AddItem "********************************"
List2.AddItem "遗漏0次    " + Str(Lost3(0))
List2.AddItem "遗漏1次    " + Str(Lost3(1))
 List2.AddItem "遗漏2次    " + Str(Lost3(2))
 List2.AddItem "遗漏3次    " + Str(Lost3(3))
 List2.AddItem "遗漏4-9    " + Str(Lost3(4))
  List2.AddItem "遗漏>10    " + Str(Lost3(5))

End Sub
Private Function Check2Max(InNum() As Integer) As MaxMin
    Dim max, min As Integer
   Dim Maxn, Minn As Integer
   Dim i As Integer
    max = InNum(0)
    
    min = InNum(0)
    For i = 1 To 3
     If max < InNum(i) Then
      max = InNum(i)
      Maxn = i
     End If
     If min > InNum(i) Then
      min = InNum(i)
      Minn = i
     End If
   Next i
    Check2Max.max = Maxn
    Check2Max.min = Minn


End Function
Private Sub Check2Sum(C2index As Integer, BSum2() As Integer)
Dim tempstr As String
Dim i, j As Integer
Dim Lost3(5) As Integer
For i = LBound(BSum2) To UBound(BSum2)
       BSum2(i) = 0
Next i
tempstr = Right(Left(List1.List(C2index - 2), 43), 17)
  tempstr = Replace(tempstr, " ", "")
  Call CheckLost3(Lost3(), tempstr)
   For i = 0 To 3
   BSum2(i) = Lost3(i)
    Next i
tempstr = Right(Left(List1.List(C2index - 1), 43), 17)
  tempstr = Replace(tempstr, " ", "")
  Call CheckLost3(Lost3(), tempstr)
   For i = 0 To 3
   BSum2(i) = BSum2(i) + Lost3(i)
Next i



End Sub

Private Sub List4ShowLost3(L1index As Integer)
Dim i, j, k As Integer
Dim tempstr As String
Dim LineStr As String
Dim LineStr2 As String
Dim Lost3(5) As Integer
Dim S3, S10 As Integer
Dim BSum2(3) As Integer
Dim LineSum2 As String
Dim Lost3Sum5(5), Lost3Sum10(5) As Integer



 For i = 6 To 0 Step -1
    tempstr = Right(Left(List1.List(L1index - i), 43), 17)
    Call Check2Sum(L1index - i, BSum2())
LineStr = tempstr + "  |  "
  tempstr = Replace(tempstr, " ", "")
  Call CheckLost3(Lost3(), tempstr)
 For j = 0 To 3
    LineStr = LineStr + Format(Str(Lost3(j)), "00") + " "
    S3 = S3 + Lost3(j)
 Next j
    
    LineSum2 = Format(Str(BSum2(0)), "0") + ":" + Format(Str(BSum2(1)), "0") + ":" + Format(Str(BSum2(2)), "0") + ":" + Format(Str(BSum2(3)), "0")
 
    LineStr = LineStr + " * " + Format(Str(Lost3(4)), "00") + " * " + Format(Str(Lost3(5)), "00") + "  *   " + Format(Str(S3), "0") + ":" + Format(Str(Lost3(4)), "0") + ":" + Format(Str(Lost3(5)), "0")
    S3 = 0
    
    '''leaf
    
    If Check2Max(BSum2).max = Check2Max(Lost3).max Then
      LineSum2 = LineSum2 + "  爆发“"
    End If
    If Check2Max(BSum2).min = Check2Max(Lost3).max Then
      LineSum2 = LineSum2 + "  回补"
    End If
    
 
 
 
 List4.AddItem LineStr + "   " + LineSum2
 LineStr = ""
 LineSum2 = ""
Next i



''*********************
List4.AddItem "*****************************"

For i = 10 To 0 Step -1
    tempstr = Right(Left(List1.List(L1index - i), 43), 17)
    LineStr = tempstr + "  |  "
    tempstr = Replace(tempstr, " ", "")
    Call CheckLost3(Lost3(), tempstr)
    For j = 0 To 5
        Lost3Sum10(j) = Lost3Sum10(j) + Lost3(j)
    
     If i < 5 Then
       Lost3Sum5(j) = Lost3Sum5(j) + Lost3(j)
    End If
    Next j
Next i
LineStr = "  5期总和             "
LineStr2 = " 10期总和             "

  For j = 0 To 3
    LineStr = LineStr + Format(Str(Lost3Sum5(j)), "00") + " "
    S3 = S3 + Lost3Sum5(j)
    
       LineStr2 = LineStr2 + Format(Str(Lost3Sum10(j)), "00") + " "
    S10 = S10 + Lost3Sum10(j)
 Next j
LineStr = LineStr + " * " + Format(Str(Lost3Sum5(4)), "00") + " * " + Format(Str(Lost3Sum5(5)), "00") + "  *   " + Format(Str(S3), "0") + ":" + Format(Str(Lost3Sum5(4)), "0") + ":" + Format(Str(Lost3Sum5(5)), "0")
List4.AddItem LineStr

LineStr2 = LineStr2 + " * " + Format(Str(Lost3Sum10(4)), "00") + " * " + Format(Str(Lost3Sum10(5)), "00") + "  *   " + Format(Str(S10), "0") + ":" + Format(Str(Lost3Sum10(4)), "0") + ":" + Format(Str(Lost3Sum10(5)), "0")
List4.AddItem LineStr2





End Sub

Private Sub List1_Click()
Dim i, j As Integer
Dim List1Str(6) As String
Dim LineStr As String
Dim NowLost As String
Dim tempstr As String
Dim Lost3(5) As Integer
Dim LostSum3(5) As Integer
List2.Clear
List4.Clear
Call List4_init
If List1.ListIndex > 7 Then
 For i = 6 To 0 Step -1
  List1Str(i) = Right(Left(List1.List(List1.ListIndex - i), 17), 12)
  For j = 1 To 6
    LineStr = LineStr + Mid(List1Str(i), 2 * j - 1, 2) + " "
  Next j
  List2.AddItem LineStr
  LineStr = ""
Next i
End If


If List1.ListIndex > 5 Then

 For i = 2 To 1 Step -1
  tempstr = Right(Left(List1.List(List1.ListIndex - i), 43), 17)
  tempstr = Replace(tempstr, " ", "")
  Call CheckLost3(Lost3(), tempstr)
  For j = 0 To 5
   LostSum3(j) = LostSum3(j) + Lost3(j)
  Next j

Next i
Call ShowLost3(LostSum3())
End If

 List2.AddItem "****************************"

NowLost = Right(Left(List1.List(List1.ListIndex - i), 43), 17)
List2.AddItem NowLost

  For j = 0 To 5
   LostSum3(j) = 0
   Next j

'Call CheckLost3(Lost3(), Replace(NowLost, " ", ""))
For i = 1 To 0 Step -1
  tempstr = Right(Left(List1.List(List1.ListIndex - i), 43), 17)
  tempstr = Replace(tempstr, " ", "")
  Call CheckLost3(Lost3(), tempstr)
  For j = 0 To 5
   LostSum3(j) = LostSum3(j) + Lost3(j)
  Next j

Next i
Call ShowLost3(LostSum3())

List4ShowLost3 (List1.ListIndex)
List4.AddItem " "
Call ShowNowCheck(List1.ListIndex)
Call ShowLast10(List1.ListIndex)

End Sub
Private Sub ShowNowCheck(Index As Integer)
Dim tempstr(4) As String
Dim LNum(33) As Integer
Dim LineStr As String
Dim LineStrL As String
Dim LL0, LL1, LL2, LL3 As String
Dim i, j As Integer
Dim Tempnum As Integer
For i = 0 To 3
    tempstr(i) = Right(Left(List1.List(Index - i), 17), 12)
Next i
For i = 0 To 3
    For j = 1 To 6
        LNum(Val(Mid(tempstr(i), 2 * j - 1, 2))) = LNum(Val(Mid(tempstr(i), 2 * j - 1, 2))) + 1
    Next j
Next i

For i = 1 To 6
    For j = 1 To 6
        If Mid(tempstr(0), 2 * i - 1, 2) = Mid(tempstr(1), 2 * j - 1, 2) Then
            LNum(Val(Mid(tempstr(0), 2 * i - 1, 2))) = 5
         End If
    Next j
Next i



 LineStr = "3期内数据: "
 LineStrL = "3期外数据: "

For i = 1 To 33
 If LNum(i) > 0 And LNum(i) < 5 Then
    LineStr = LineStr + Format(Str(i), "00") + " "
 Else
    LineStrL = LineStrL + Format(Str(i), "00") + " "
 End If
 LNum(i) = 0
Next i
List4.AddItem LineStr
List4.AddItem LineStrL
For i = 1 To 6
    Tempnum = Val(Mid(tempstr(0), 2 * i - 1, 2))
    If Tempnum = 1 Then
        LNum(2) = 1
    ElseIf Tempnum = 33 Then
        LNum(32) = 1
    Else
        LNum(Tempnum - 1) = 1
        LNum(Tempnum + 1) = 1
       For j = 1 To 6
        If Abs(Val(Mid(tempstr(1), 2 * j - 1, 2)) - Tempnum) = 1 Then
            LNum((Val(Mid(tempstr(1), 2 * j - 1, 2)))) = 2
        End If
        If Abs(Val(Mid(tempstr(2), 2 * j - 1, 2)) - Tempnum) = 1 Then
            LNum((Val(Mid(tempstr(2), 2 * j - 1, 2)))) = 3
        End If
        If Abs(Val(Mid(tempstr(3), 2 * j - 1, 2)) - Tempnum) = 1 Then
            LNum((Val(Mid(tempstr(3), 2 * j - 1, 2)))) = 4
        End If
       Next j
    
    
    
    End If
Next i
    
    LineStr = "相邻号码0:  "
    LL1 = "相邻号码1:  "
    LL2 = "相邻号码2:  "
    LL3 = "相邻号码3:  "
 For i = 1 To 33
 If LNum(i) = 1 Then
    LineStr = LineStr + Format(Str(i), "00") + " "
 ElseIf LNum(i) = 2 Then
    LL1 = LL1 + Format(Str(i), "00") + " "
 ElseIf LNum(i) = 3 Then
    LL2 = LL2 + Format(Str(i), "00") + " "
ElseIf LNum(i) = 4 Then
    LL3 = LL3 + Format(Str(i), "00") + " "
End If
 
 
 LNum(i) = 0
Next i
    List4.AddItem LineStr
    List4.AddItem LL1
    List4.AddItem LL2
    List4.AddItem LL3
    
List4.AddItem "****************************************"
'''显示其他遗漏
LL1 = ""
LL2 = ""
LL3 = ""
LineStr = ""
LineStrL = ""
LineStrL = ""
For i = 1 To 33
 If LoseR(TotalNum, i) = 0 Then
    LL0 = LL0 + Format(Str(i), "00") + " "
ElseIf LoseR(TotalNum, i) = 1 Then
    LL1 = LL1 + Format(Str(i), "00") + " "
ElseIf LoseR(TotalNum, i) = 2 Then
    LL2 = LL2 + Format(Str(i), "00") + " "
ElseIf LoseR(TotalNum, i) = 3 Then
    LL3 = LL3 + Format(Str(i), "00") + " "
ElseIf LoseR(TotalNum, i) > 3 And LoseR(TotalNum, i) < 10 Then
    LineStr = LineStr + Format(Str(i), "00") + " "
ElseIf LoseR(TotalNum, i) > 9 Then
    LineStrL = LineStrL + Format(Str(i), "00") + " "
End If

Next i

List4.AddItem "遗漏0次： " + LL0
List4.AddItem "遗漏1次： " + LL1
List4.AddItem "遗漏2次： " + LL2
List4.AddItem "遗漏3次： " + LL3
List4.AddItem "遗漏4-9： " + LineStr
List4.AddItem "遗漏10次： " + LineStrL



    
    



End Sub


Private Sub List2_Click()
If ChooseOp7 = True Then
 Combo4.ListIndex = List2.ListIndex
End If
'List1.ListIndex = List1.ListCount - 1 - (TAILLINES + 2 - List2.ListIndex)


End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
End Sub

Private Sub OddC_Click()
If OddC.Value = 1 Then
 OddCheck.Enabled = True
 Check4.Enabled = True
 Combo2.Enabled = True
Else
 OddCheck.Enabled = False
 Check4.Enabled = False
 Combo2.Enabled = False
End If

End Sub

Private Sub OddCheck_Click()
If OddCheck.Value = 1 Then
    NowOddNums = 3
    Combo2.Enabled = True
    Combo2.ListIndex = 3
    Check4.Enabled = False
Else
    Combo2.Enabled = False
    
    Check4.Enabled = True
End If
End Sub

Private Sub Option4_Click(Index As Integer)
Dim i As Integer
Dim TempData As Integer
For i = 0 To 3
If Option5(i).Value = True Then
 TempData = i
 Exit For
End If
Next i

If (4 - TempData - Index) >= 0 And (4 - TempData - Index) <= 2 Then
    Option6(6 - TempData - Index - 2).Value = True
End If
End Sub

Private Sub Option5_Click(Index As Integer)
Dim i As Integer
Dim TempData As Integer
For i = 0 To 3
If Option6(i).Value = True Then
    TempData = i
    Exit For
End If
Next i
If (4 - TempData - Index) >= 0 And (4 - TempData - Index) <= 2 Then
    Option6(6 - TempData - Index - 2).Value = True
End If

End Sub

Private Sub Option7_Click(Index As Integer)
ChooseOp7 = True
Dim i As Integer
Call Command2_Click(Index)
Combo4.Clear
For i = 0 To List2.ListCount - 1
    Combo4.AddItem List2.List(i)
Next i

End Sub

Private Sub OptionReg1_Click(Index As Integer)
NowRegion1 = OptionReg1(Index).Caption
Call CheckOptionReg
End Sub

Private Sub OptionReg2_Click(Index As Integer)
NowRegion2 = OptionReg2(Index).Caption
Call CheckOptionReg
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Label5.Caption = Str(CInt(X)) + " -" + Str(Picture_Y - CInt(Y))
For i = 1 To 100
 If Num1(i) = CInt(X) Then
    Label7.Caption = Str(CInt(X)) + " -" + Str(Num2(i))
 Exit For
 End If
Next i

End Sub

Private Sub RegCheck_Click()
If RegCheck.Value = 1 Then
 Frame1.Enabled = True
 Label2.Caption = "详细区间"
Else
 Frame1.Enabled = False
 Label2.Caption = "忽略区间"
End If

End Sub

Private Sub ShowRegLost(cmdindex As Integer)
Dim i  As Integer
Dim j, k As Integer
Dim NowLost As Integer
Dim tempstr As String
Dim RegNum(300) As String
Dim OutStr(300) As Integer
Dim Tempnum As Integer
Dim LineStr As String
List4.Clear
' 3区间分布
If cmdindex = 0 Then
For i = 1 To 300
 RegNum(i) = Left(Right(List1.List(List1.ListCount - i), 8), 3)
Next i

For i = 0 To List2.ListCount - 2
 Tempnum = FindLost(RegNum(), Left(List2.List(i), 3), OutStr())
 LineStr = Left(List2.List(i), 3) + "- "
  For j = 0 To Tempnum
   LineStr = LineStr + Format(Str(OutStr(j)), "00") + " "
 Next j
  List4.AddItem LineStr
  LineStr = ""
  
Next i

End If

' 4区间分布
If cmdindex = 1 Then
For i = 1 To 300
 RegNum(i) = Right(List1.List(List1.ListCount - i), 4)
 Next i

For i = 0 To List2.ListCount - 2
 Tempnum = FindLost(RegNum(), Left(List2.List(i), 4), OutStr())
 LineStr = Left(List2.List(i), 4) + "- "
  For j = 0 To Tempnum
   LineStr = LineStr + Format(Str(OutStr(j)), "00") + " "
 Next j
  List4.AddItem LineStr
  LineStr = ""
 
Next i
End If

'6-1区间分布
If cmdindex = 2 Then
For i = 1 To 300
 RegNum(i) = Left(Right(List1.List(List1.ListCount - i), 16), 3)
 Next i

For i = 0 To List2.ListCount - 2
 Tempnum = FindLost(RegNum(), Left(List2.List(i), 3), OutStr())
 LineStr = Left(List2.List(i), 3) + "- "
  For j = 0 To Tempnum
   LineStr = LineStr + Format(Str(OutStr(j)), "00") + " "
 Next j
  List4.AddItem LineStr
  LineStr = ""
  
Next i

End If


'6-2区间分布
If cmdindex = 3 Then
For i = 1 To 300
 RegNum(i) = Left(Right(List1.List(List1.ListCount - i), 13), 3)
 Next i

For i = 0 To List2.ListCount - 2
 Tempnum = FindLost(RegNum(), Left(List2.List(i), 3), OutStr())
 LineStr = Left(List2.List(i), 3) + "- "
  For j = 0 To Tempnum
   LineStr = LineStr + Format(Str(OutStr(j)), "00") + " "
 Next j
  List4.AddItem LineStr
  LineStr = ""
Next i

End If

If cmdindex = 4 Then        '' 11区域间隔分布
For i = 1 To 300
 RegNum(i) = Left(Right(List1.List(List1.ListCount - i), 20), 3)
 Next i

For i = 0 To List2.ListCount - 2
 Tempnum = FindLost(RegNum(), Left(List2.List(i), 3), OutStr())
 LineStr = Left(List2.List(i), 3) + "- "
  For j = 0 To Tempnum
   LineStr = LineStr + Format(Str(OutStr(j)), "00") + " "
 Next j
  List4.AddItem LineStr
  LineStr = ""
  
Next i

End If
LineStr = ""
For i = 1 To 100
   NowLost = FindNowLost(RegNum(), i)
  LineStr = LineStr + RegNum(i) + "-" + Format(Str(NowLost), "00") + " "
  k = k + 1
  If k = 10 Then
   List4.AddItem LineStr
   LineStr = ""
   k = 0
  End If
Next i

End Sub
Private Sub Show10Appear(cmdindex As Integer)
Dim i, j, k As Integer
Dim tempstr As String
Dim LineStr As String
Dim AppNum10(7) As Integer
Dim Interval As Integer
Interval = 12
List4.Clear
If cmdindex = 0 Then
For i = 0 To List1.ListCount - 1
 tempstr = Left(Right(List1.List(i), 8), 3)
  AppNum10(1) = AppNum10(1) + Val(Mid(tempstr, 1, 1))
  AppNum10(2) = AppNum10(2) + Val(Mid(tempstr, 2, 1))
  AppNum10(3) = AppNum10(3) + Val(Mid(tempstr, 3, 1))
  k = k + 1
  If k = Interval Then
    LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
    List4.AddItem LineStr
    LineStr = ""
    AppNum10(1) = 0
    AppNum10(2) = 0
    AppNum10(3) = 0
    k = 0
  End If
 Next i
  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
 List4.AddItem LineStr
 End If
'''**********************************************************************************
If cmdindex = 1 Then
For i = 0 To List1.ListCount - 1
 tempstr = Right(List1.List(i), 4)
  AppNum10(1) = AppNum10(1) + Val(Mid(tempstr, 1, 1))
  AppNum10(2) = AppNum10(2) + Val(Mid(tempstr, 2, 1))
  AppNum10(3) = AppNum10(3) + Val(Mid(tempstr, 3, 1))
 AppNum10(4) = AppNum10(4) + Val(Mid(tempstr, 4, 1))
  k = k + 1
  If k = Interval Then
    LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  " + Format(Str(AppNum10(4)), "00")
    List4.AddItem LineStr
    LineStr = ""
    AppNum10(1) = 0
    AppNum10(2) = 0
    AppNum10(3) = 0
    AppNum10(4) = 0
    
    k = 0
  End If
 Next i
  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  " + Format(Str(AppNum10(4)), "00")
 List4.AddItem LineStr
 End If
''*******************************************************************************************
If cmdindex = 2 Then
For i = 0 To List1.ListCount - 1
 tempstr = Left(Right(List1.List(i), 16), 7)
  For j = 1 To 7
    AppNum10(j) = AppNum10(j) + Val(Mid(tempstr, j, 1))
    Next j
'  AppNum10(1) = AppNum10(1) + Val(Mid(Tempstr, 1, 1))
 ' AppNum10(2) = AppNum10(2) + Val(Mid(Tempstr, 2, 1))
 ' AppNum10(3) = AppNum10(3) + Val(Mid(Tempstr, 3, 1))
  k = k + 1
  If k = Interval Then
  '  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
    For j = 1 To 7
     LineStr = LineStr + Format(Str(AppNum10(j)), "00") + " |  "
        AppNum10(j) = 0
     Next j
     
    List4.AddItem LineStr
    LineStr = ""
    k = 0
  End If
 Next i
'  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
  For j = 1 To 7
     LineStr = LineStr + Format(Str(AppNum10(j)), "00") + " |  "
        
     Next j

 List4.AddItem LineStr
 End If
''****************************
If cmdindex = 3 Then
For i = 0 To List1.ListCount - 1
 tempstr = Left(Right(List1.List(i), 13), 4)
  AppNum10(1) = AppNum10(1) + Val(Mid(tempstr, 1, 1))
  AppNum10(2) = AppNum10(2) + Val(Mid(tempstr, 2, 1))
  AppNum10(3) = AppNum10(3) + Val(Mid(tempstr, 3, 1))
 AppNum10(4) = AppNum10(4) + Val(Mid(tempstr, 4, 1))
  k = k + 1
  If k = Interval Then
    LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  " + Format(Str(AppNum10(4)), "00")
    List4.AddItem LineStr
    LineStr = ""
    AppNum10(1) = 0
    AppNum10(2) = 0
    AppNum10(3) = 0
    AppNum10(4) = 0
    
    k = 0
  End If
 Next i
  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  " + Format(Str(AppNum10(4)), "00")
 List4.AddItem LineStr
 End If
 
 '*********************************


If cmdindex = 4 Then
For i = 0 To List1.ListCount - 1
 tempstr = Left(Right(List1.List(i), 20), 3)
  AppNum10(1) = AppNum10(1) + Val(Mid(tempstr, 1, 1))
  AppNum10(2) = AppNum10(2) + Val(Mid(tempstr, 2, 1))
  AppNum10(3) = AppNum10(3) + Val(Mid(tempstr, 3, 1))
  k = k + 1
  If k = Interval Then
    LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
    List4.AddItem LineStr
    LineStr = ""
    AppNum10(1) = 0
    AppNum10(2) = 0
    AppNum10(3) = 0
    k = 0
  End If
 Next i
  LineStr = LineStr + Format(Str(AppNum10(1)), "00") + " |  " + Format(Str(AppNum10(2)), "00") + " |  " + Format(Str(AppNum10(3)), "00") + " |  "
 List4.AddItem LineStr
 End If
 List4.AddItem "剩余  " + Str(k) + "   期数据"
 List4.ListIndex = List4.ListCount - 1
End Sub

Private Sub Text2_Change()
Dim tempstr As String
Dim i, TailValue As Integer
Dim TailFalg(9) As Integer
For i = 1 To 33
 Check6(i).Value = 0
Next i
tempstr = Text2.Text
tempstr = Replace(tempstr, " ", "")
For i = 1 To Len(tempstr)
 ' Tail(Val(Mid(Tempstr, i, 1))) = 1
    TailValue = Val(Mid(tempstr, i, 1))
    If TailValue < 4 And TailValue > 0 Then
     Check6(TailValue).Value = 1
     Check6(10 + TailValue).Value = 1
     Check6(20 + TailValue).Value = 1
     Check6(30 + TailValue).Value = 1
    End If
    If TailValue > 3 Then
     Check6(TailValue).Value = 1
     Check6(10 + TailValue).Value = 1
     Check6(20 + TailValue).Value = 1
    End If
     If TailValue = 0 Then
     Check6(10 + TailValue).Value = 1
     Check6(20 + TailValue).Value = 1
     Check6(30 + TailValue).Value = 1
    End If
    
Next i

End Sub
