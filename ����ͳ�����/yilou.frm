VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lottery"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "yilou.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   15270
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C000C0&
      Caption         =   "随机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "滤号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   63
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "选号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   61
      Top             =   3720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   0
      Top             =   3720
   End
   Begin VB.ListBox baifenbi 
      Height          =   480
      Left            =   2880
      TabIndex        =   58
      Top             =   8400
      Width           =   9615
   End
   Begin VB.Frame Frame7 
      Caption         =   "遗漏和值"
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
      Left            =   12600
      TabIndex        =   51
      Top             =   9480
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "下期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   2
      Left            =   12600
      TabIndex        =   48
      Top             =   8040
      Width           =   3135
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "上期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   12600
      TabIndex        =   45
      Top             =   6600
      Width           =   3135
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "选期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   9360
      TabIndex        =   42
      Top             =   6600
      Width           =   3255
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "88 88 88 88 88 88"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.ListBox NowList 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2220
      Left            =   12600
      TabIndex        =   41
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      Caption         =   "总计统计期数"
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
      Left            =   0
      TabIndex        =   34
      Top             =   8400
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox labelList 
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
      ForeColor       =   &H000080FF&
      Height          =   780
      Left            =   2880
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3480
      Width           =   12735
   End
   Begin VB.ListBox yllist 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   9360
      TabIndex        =   32
      Top             =   4320
      Width           =   3255
   End
   Begin VB.ListBox mainlist2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   2880
      TabIndex        =   30
      Top             =   4320
      Width           =   6495
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   25
      Top             =   9000
      Width           =   12375
      Begin VB.CommandButton Command11 
         Caption         =   "伴侣数字"
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
         Left            =   7560
         TabIndex        =   62
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
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
         Height          =   735
         Left            =   6360
         TabIndex        =   60
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0000C000&
         Caption         =   "数据更新"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "连号统计"
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
         Left            =   8640
         TabIndex        =   57
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000C0C0&
         Caption         =   "出现次数分区统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9720
         MaskColor       =   &H00008080&
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "和值，区间，奇偶"
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
         Left            =   4920
         TabIndex        =   55
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "RedBall"
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
         Left            =   11640
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "详细区间分布"
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
         Left            =   3720
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame Frame4 
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
         TabIndex        =   27
         Top             =   120
         Width           =   2655
         Begin VB.CommandButton ColdButton 
            Caption         =   "热冷数字"
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
            Height          =   615
            Left            =   1320
            TabIndex        =   29
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "遗漏统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "数据组合"
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
         Left            =   2640
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   9
      Top             =   5040
      Width           =   2775
      Begin VB.ListBox ResultList 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1200
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2160
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "遗漏统计期数"
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
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   7320
      Width           =   6495
      Begin VB.OptionButton Option1 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ListBox AllData 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.ListBox MainList 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label LessLabel 
      Caption         =   "小于10次统计："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   8040
      Width           =   9735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public TheNumStr As String
Dim qishu As Integer
Dim haoma As Integer
Dim yilou As Integer
Dim TotalNum As Integer
Dim Less10Times(9) As Integer
Dim Less5Times(5) As Integer

Dim AppearTimes(1 To 33) As Integer
Dim SaveData(1 To 1000, 1 To 33) As Integer
Dim LoseNum(1 To 1000, 1 To 33) As Integer
Dim TxtUpdateFlag As Boolean


'Dim RedNum(1 To 1000, 1 To 16) As Integer

Private Function SaveTotaldata()
Dim i, j As Integer
Dim TempStr As String
Dim TData(7) As String
Dim tempsave As Integer
Dim AppearSum As Integer
Dim AppearCount(1, 1 To 33) As Integer
Dim MaxAppear, MaxNum As Integer
Dim MaxNumStr, MaxAppearStr As String
For i = 1 To TotalNum
    For j = 1 To 33
        SaveData(i, j) = 0
    Next j
     
Next i


Dim Linestr, AppearStr As String
For i = 1 To TotalNum
    TempStr = AllData.List(AllData.ListCount - TotalNum - 1 + i)
'For i = 80 To 1 Step -1

    Call GetNum(TempStr, TData())

    For j = 1 To 6      ''1 to 6
 
   ' SaveData(i, j) = tdata(j)
    'SaveData(i, Val(tdata(j))) = SaveData(i, Val(tdata(j))) + 1
    SaveData(i, Val(TData(j))) = 1
    AppearTimes(Val(TData(j))) = AppearTimes(Val(TData(j))) + 1
    Next j
   
    
Next i
 '***********Leaf  remove add
 ' For j = 1 To 6
  ' AppearTimes(Val(TData(j))) = AppearTimes(Val(TData(j))) - 1     ' Don't add the last times
  'Next j


For j = 1 To 33
       If SaveData(1, j) = 1 Then
            LoseNum(1, j) = 0
        Else
            LoseNum(1, j) = 1
        End If
    For i = 2 To TotalNum
        If SaveData(i, j) = 1 Then
            LoseNum(i, j) = 0
        Else
            LoseNum(i, j) = LoseNum(i - 1, j) + 1
        End If
     Next i
Next j
 For i = 1 To TotalNum
    For j = 1 To 33
        'Text2.Text = Text2.Text + " " + Str(SaveData(i, j))
       ' Text2.Text = Text2.Text + " " + Format(Str(LoseNum(i, j)), "00")
       Linestr = Linestr + Format(Str(LoseNum(i, j)), "00") + "| "
       
     Next j
        MainList.AddItem Format(Str(i), "000") + ": " + Linestr
        Linestr = ""
        'Text2.Text = Text2.Text + vbCrLf
Next i
    For i = 1 To 33
       ' AppearSum = AppearSum + AppearTimes(i)
        AppearStr = AppearStr + Format(AppearTimes(i), "00") + "| "
        AppearCount(0, i) = i
        AppearCount(1, i) = AppearTimes(i)
        AppearTimes(i) = 0
        
     Next i
     'leaf 排序
     labelList.AddItem "     " + AppearStr '+ Str(AppearSum)
    AppearSum = 0
   
     For j = 1 To 32
     For i = 1 To 33 - j
       
        If AppearCount(1, i) < AppearCount(1, i + 1) Then
            MaxAppear = AppearCount(1, i + 1)
            MaxNum = AppearCount(0, i + 1)
            AppearCount(1, i + 1) = AppearCount(1, i)
            AppearCount(0, i + 1) = AppearCount(0, i)
            AppearCount(1, i) = MaxAppear
            AppearCount(0, i) = MaxNum
        End If
     Next i
     Next j
    
    For i = 1 To 33
        MaxAppearStr = MaxAppearStr + Format(Str(AppearCount(1, i)), "00") + "| "
        MaxNumStr = MaxNumStr + Format(Str(AppearCount(0, i)), "00") + "| "
    
    Next i
    
    labelList.AddItem "     " + MaxNumStr
    labelList.AddItem "     " + MaxAppearStr
   
    
   ' MainList.AddItem "01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33"
    
   ' Text2.SelStart = Len(Text2.Text) - 10
End Function
Private Function YLCount()
Dim j As Integer
Dim TempStr(9) As String
Dim tempmax As String
For j = 1 To 33
    'Select Case LoseNum(80, j)
   ' Case 0
   TopLose(j) = LoseNum(TotalNum, j)
    If LoseNum(TotalNum, j) < 10 Then
        TempStr(LoseNum(TotalNum, j)) = TempStr(LoseNum(TotalNum, j)) + Format(Str(j), "00") + " "
    End If
    If LoseNum(TotalNum, j) >= 10 Then
        tempmax = tempmax + Format(Str(j), "00") + " "
    End If
 Next j
 
 
For j = 0 To 9
    yllist.AddItem Format(Str(j), "00") + ": " + TempStr(j)
Next j
    yllist.AddItem "10: " + tempmax
  

End Function
Private Function ShowBaiFenBi()
Dim i, j As Integer
Dim Tstr0, Tstr1, Tstr2, Tstr3 As String
Dim TempStr(2, 1 To 33) As String
Dim TempFlag1(1 To 33), TempFlag2(1 To 33) As Integer
For i = 1 To 33
  ' TempStr(0, i) = "☆"
     TempStr(0, i) = "  "
    TempStr(1, i) = "  "
    TempStr(2, i) = "  "
    
Next i
For j = 1 To 33

    For i = TotalNum To TotalNum - 9 Step -1
        'If LoseNum(i, j) < 5 Then
            'TempStr(0, j) = "★ "
         If LoseNum(i, j) = 0 And LoseNum(TotalNum, j) < 3 Then
                TempFlag1(j) = TempFlag1(j) + 1
           End If
           
          If LoseNum(i, j) = 0 And LoseNum(TotalNum, j) < 5 Then ' then And LoseNum(TotalNum, j) > 3 Then
                TempFlag2(j) = TempFlag2(j) + 1
           End If
     Next i
    If TempFlag1(j) >= 1 Then
        TempStr(0, j) = "* "
    End If
    If TempFlag2(j) >= 1 Then
'        TempStr(1, j) = "★ "
        TempStr(1, j) = "* "
    End If
    If TempFlag2(j) >= 2 Then
'        TempStr(1, j) = "★ "
        TempStr(2, j) = "* "
    End If
    
    
    
    

Next j

For i = 1 To 33
    Tstr0 = Tstr0 + Format(Str(i), "00") + "|"
    Tstr1 = Tstr1 + TempStr(0, i) + "|"
    Tstr2 = Tstr2 + TempStr(1, i) + "|"
    Tstr3 = Tstr3 + TempStr(2, i) + "|"
Next i
baifenbi.AddItem Tstr0
baifenbi.AddItem Tstr1
baifenbi.AddItem Tstr2
baifenbi.AddItem Tstr3
End Function



Private Sub AllData_Click()
If MainList.ListCount > 60 And (AllData.ListCount - 1 - AllData.ListIndex) < TotalNum Then
    MainList.ListIndex = MainList.ListCount - 1 - (AllData.ListCount - 1 - AllData.ListIndex)
End If
End Sub

Private Sub AllData_DblClick()
If AllData.ListIndex > 50 Then
    AllData.ListIndex = AllData.ListIndex - 50
    Call AllData_Click
End If
End Sub

Private Sub ColdButton_Click()

mainlist2.Clear
LessLabel.Caption = "小于10次统计:"
Dim i, j As Integer
Dim Less10, LossSum As Integer
Dim Less10sum, Losssumsum As Integer
Dim ySum As Integer
Dim TempStr1, TempStr2 As String
'For i = 0 To 5
 '   If Option1(i).Value = True Then
  '        Ttnum = i + 5
   ' End If
'Next i
  For i = TotalNum - qishu + 1 To TotalNum
    For j = 1 To 33
        If SaveData(i, j) = 1 Then
            TempStr1 = TempStr1 + Format(Str(j), "00") + " "
            
            TempStr2 = TempStr2 + Format(Str(LoseNum(i - 1, j)), "00") + " "
            LossSum = LossSum + LoseNum(i - 1, j)
            If LoseNum(i - 1, j) < 10 Then
                Less10 = Less10 + 1
                Less10Times(LoseNum(i - 1, j)) = Less10Times(LoseNum(i - 1, j)) + 1
            End If
            If LoseNum(i - 1, j) < 6 Then
                Less5Times(LoseNum(i - 1, j)) = Less5Times(LoseNum(i - 1, j)) + 1
           End If
                        
            
        End If
    Next j
    mainlist2.AddItem TempStr1 + " | " + TempStr2 + "|" + Format(Str(Less10), "00") + "|" + Format(Str(LossSum), "00") + "|" + Str(Format((LossSum / 6), "0.00"))
     Less10sum = Less10sum + Less10
            Losssumsum = Losssumsum + LossSum
    TempStr1 = ""
    TempStr2 = ""
    LossSum = 0
    Less10 = 0
        
Next i
    
    mainlist2.AddItem "总计" + Str(qishu) + "期 小于10次平均值为 " + Str(Format(Less10sum / (qishu + 1), "00.00")) + "和值平准值为" + Str(Format((Losssumsum / (qishu + 1)), "00.00"))
  
    Less10sum = 0
    Losssumsum = 0
For i = 0 To 9
    
    LessLabel.Caption = LessLabel.Caption + Str(i) + ":" + Str(Less10Times(i)) + " "
    Less10Times(i) = 0
Next i


End Sub

Private Sub Command1_Click()
MainList.Clear
If labelList.ListCount > 3 Then
    labelList.RemoveItem 1
    labelList.RemoveItem 1
    labelList.RemoveItem 1
 ' labelList.RemoveItem 3

End If
Command1.Enabled = False
Call SaveTotaldata
yllist.Clear
Call YLCount
MainList.ListIndex = MainList.ListCount - 1

ColdButton.Enabled = True
Frame1.Enabled = True
End Sub

Private Sub Command10_Click()
Form4.Show
End Sub

Private Sub Command11_Click()
 


LongTime.Show
End Sub

Private Sub Command12_Click()
Form6.Show
'BitBlt Me.hDC, 100, 100, 800, 300, GetDC(0), 0, 0, vbSrcCopy

'SavePicture Me.Image, "C:\JieTu.BMP"

End Sub

Private Sub Command13_Click()
Form7.Show
End Sub

Private Sub Command2_Click()
ResultList.Clear
Dim i As Integer
Dim ChoStr(9) As String
Dim Result(3, 5) As Variant
For i = 0 To 9
    ChoStr(i) = Text1(i).Text
Next i
Call zuhe8(ChoStr(), Result)

For i = 0 To 3
ResultList.AddItem Result(i, 0) + " " + Result(i, 1) + " " + Result(i, 2) + " " + Result(i, 3) + " " + Result(i, 4) + " " + Result(i, 5)
'Text2.Text = Text2.Text + Result(i, 0) + " " + Result(i, 1) + " " + Result(i, 2) + " " + Result(i, 3) + " " + Result(i, 4) + " " + Result(i, 5) + vbCrLf
Next i
    
End Sub


Private Sub Command3_Click()
'LongTime.Show
Region6.Show
End Sub

Private Sub Command4_Click()
RedBall.Show
BlueBall.Show
End Sub

Private Sub Command5_Click()
HZForm.Show
End Sub

Private Sub Command6_Click()
'Call Shell("notepad save.txt", vbNormalFocus)
Form5.Show

End Sub

Private Sub Command7_Click()
'MTime.Show
Form3.Show
End Sub

Private Sub Command8_Click()
If TxtUpdateFlag = False Then
    Call Shell("notepad all.txt", vbNormalFocus)
    TxtUpdateFlag = True
End If
'Timer1.Enabled = True
Call Form_Load
End Sub

Private Sub Command9_Click()
'Option1(9).Value = True
    qishu = 200

Call ColdButton_Click
'MTime.Show
LoseForm.Show
End Sub

Private Sub Form_Load()
Dim TempStr As String
Dim i As Integer



For i = 0 To 9
Text1(i).Font.Size = 12
Next i

Open App.Path + "\GoldData.txt" For Input As #1
    
Do While Not EOF(1)
    Line Input #1, TempStr
    AllData.AddItem TempStr
Loop
Close #1
TheNumStr = Left(TempStr, 5)
    'AllData.AddItem "----------------------------"
qishu = 5
TotalNum = 300
AllData.ListIndex = AllData.ListCount - 1
labelList.AddItem "     01| 02| 03| 04| 05| 06| 07| 08| 09| 10| 11| 12| 13| 14| 15| 16| 17| 18| 19| 20| 21| 22| 23| 24| 25| 26| 27| 28| 29| 30| 31| 32| 33|"
Option2(3).Caption = Str(AllData.ListCount)
Call Command1_Click
Call ColdButton_Click
Call ShowBaiFenBi
NumStr = Str(Val(TheNumStr) + 1)

End Sub

Private Sub HotButton_Click()
yllist.Clear



End Sub

Private Sub Form_Unload(Cancel As Integer)
'MainForm.Show
End Sub

Private Sub MainList_Click()
 AllData.ListIndex = AllData.ListCount - 1 - (MainList.ListCount - 1 - MainList.ListIndex)

Dim i As Integer
For i = 0 To 5
    Label1(i).Caption = ""
Next i
Dim TempStr, TempStr1, tempstrS1, TempStr2 As String
Dim Nsum, Nsum1, Nsum2 As Integer
Dim ChooseStr(10) As String
Dim TempNum(1 To 33) As Integer
If MainList.ListCount > 1 Then
NowList.Clear
TempStr = MainList.List(MainList.ListIndex)
TempStr1 = MainList.List(MainList.ListIndex - 1)
tempstrS1 = MainList.List(MainList.ListIndex - 2)

If MainList.ListIndex = MainList.ListCount - 1 Then
    TempStr2 = TempStr
Else
    TempStr2 = MainList.List(MainList.ListIndex + 1)
End If

For i = 1 To 33
    TempNum(i) = Val(Mid(TempStr, 4 * i + 2, 2))
    If TempNum(i) < 10 Then
        ChooseStr(TempNum(i)) = ChooseStr(TempNum(i)) + Format(Str(i), "00") + " "
        If TempNum(i) = 0 Then
            Label1(0).Caption = Label1(0).Caption + Format(Str(i), "00") + " "
            Label1(1).Caption = Label1(1).Caption + Mid(TempStr1, 4 * i + 2, 2) + " "
            Nsum = Nsum + Val(Mid(TempStr1, 4 * i + 2, 2))
        End If
    Else
        ChooseStr(10) = ChooseStr(10) + Format(Str(i), "00") + " "
    End If
    TempNum(i) = Val(Mid(TempStr1, 4 * i + 2, 2))
    If TempNum(i) = 0 Then
            Label1(2).Caption = Label1(2).Caption + Format(Str(i), "00") + " "
            Label1(3).Caption = Label1(3).Caption + Mid(tempstrS1, 4 * i + 2, 2) + " "
            Nsum1 = Nsum1 + Val(Mid(tempstrS1, 4 * i + 2, 2))
    End If
    TempNum(i) = Val(Mid(TempStr2, 4 * i + 2, 2))
      If TempNum(i) = 0 Then
            Label1(4).Caption = Label1(4).Caption + Format(Str(i), "00") + " "
            Label1(5).Caption = Label1(5).Caption + Mid(TempStr, 4 * i + 2, 2) + " "
            Nsum2 = Nsum2 + Val(Mid(TempStr, 4 * i + 2, 2))
      
      
      End If
   
Next i
    Label2(0) = Str(Nsum1)
    Label2(1) = Str(Nsum)
    Label2(2) = Str(Nsum2)
For i = 0 To 10
    NowList.AddItem Format(Str(i), "00") + ": " + ChooseStr(i)
Next i
End If


End Sub

Private Sub mainlist2_Click()
If mainlist2.ListIndex <> mainlist2.ListCount - 1 Then
    MainList.ListIndex = MainList.ListCount - (mainlist2.ListCount - 1 - mainlist2.ListIndex)
    AllData.ListIndex = AllData.ListCount - (mainlist2.ListCount - 1 - mainlist2.ListIndex)
End If
End Sub

Private Sub Option1_Click(Index As Integer)
If Index < 6 Then
    qishu = Index + 5
End If
If Index = 6 Then
    qishu = 20
ElseIf Index = 7 Then
    qishu = 30
ElseIf Index = 8 Then
    qishu = 40
ElseIf Index = 9 Then
    qishu = 50
End If
Call ColdButton_Click


End Sub


Private Sub Option2_Click(Index As Integer)
If Index = 0 Then
    TotalNum = 100
ElseIf Index = 1 Then
    TotalNum = 200
ElseIf Index = 2 Then
    TotalNum = 300

ElseIf Index = 3 Then
    TotalNum = Val(Option2(Index).Caption)
End If
Call Command1_Click

'Command1.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
If (Len(Text1(Index).Text) = 2) And Index < 9 Then
    Text1(Index + 1).SetFocus
End If
If Len(Text1(9).Text) > 2 Then
    Text1(9).Text = ""
End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
Text1(Index).Text = ""
End Sub

Private Sub Timer1_Timer()
Call Form_Load
Timer1.Enabled = False

End Sub

