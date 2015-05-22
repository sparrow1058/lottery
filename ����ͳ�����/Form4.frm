VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "期选号器"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   LinkTopic       =   "Form4"
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10125
   Begin VB.Frame Frame6 
      Caption         =   "号码特性"
      Height          =   1695
      Left            =   6240
      TabIndex        =   75
      Top             =   4920
      Width           =   3615
      Begin VB.Label Label5 
         Caption         =   "遗漏分布"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   80
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "区间比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "和值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "大小比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   77
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "奇偶比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      Height          =   1500
      Left            =   8400
      TabIndex        =   74
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   6240
      TabIndex        =   73
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "数据更新"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8280
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "数据更新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "存储文本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "清除勾选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "QQ购买"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   64
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "正在选号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "自动选号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "篮球选号"
      Height          =   2055
      Left            =   0
      TabIndex        =   41
      Top             =   3360
      Width           =   6255
      Begin VB.Frame Frame4 
         Caption         =   "小号偶数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   3120
         TabIndex        =   57
         Top             =   1080
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   2280
            TabIndex        =   61
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   1560
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   840
            TabIndex        =   59
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "小号偶数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   2280
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   1560
            TabIndex        =   55
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "11111"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   840
            TabIndex        =   54
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "09"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "小号偶数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2280
            TabIndex        =   51
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "06"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   50
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "04"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   840
            TabIndex        =   49
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "02"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "小号偶数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "03"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   840
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "05"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   1560
            TabIndex        =   44
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "07"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2280
            TabIndex        =   43
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "红球区间选号"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame5 
         Caption         =   "选中号码"
         Height          =   855
         Left            =   4080
         TabIndex        =   69
         Top             =   2400
         Width           =   3975
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "选中号码"
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
            Left            =   0
            TabIndex        =   72
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FF8080&
            Caption         =   "B"
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
            Left            =   3360
            TabIndex        =   71
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3000
            TabIndex        =   70
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "31-33区间"
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
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   3735
         Begin VB.CheckBox Check1 
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   33
            Left            =   1560
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   32
            Left            =   840
            TabIndex        =   39
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   31
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "26-30区间"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   31
         Top             =   1680
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   30
            Left            =   3120
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   29
            Left            =   2400
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   28
            Left            =   1560
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   27
            Left            =   840
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "26"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   26
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "21-25区间"
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
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   25
            Left            =   3120
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   2400
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   840
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "16-20区间"
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
         Index           =   3
         Left            =   4080
         TabIndex        =   19
         Top             =   960
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   3120
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   2400
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   840
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "11-15区间"
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
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "06-10区间"
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
         Index           =   1
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   3120
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "09"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   2400
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "07"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "06"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "01-05区间"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox Check1 
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "02"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "03"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "04"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "05"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3120
            TabIndex        =   2
            Top             =   240
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BlueOK, RedOK As Boolean
Dim UpdateMode As Boolean
''拍拍网购买http://paipai.500wan.com/

Private Sub Check1_Click(Index As Integer)
Dim i As Integer
Dim NumC As Integer
Dim SumStr As Integer
Dim TempStr As String
Dim LoseStr As String
Dim ODDN, EVENN, BIGN, MidN, SmaN As Integer
Dim RegionN(1 To 7) As Integer
If RedOK Then
 If Check1(Index) = 1 Then
  Check1(Index).BackColor = vbRed
  Exit Sub
 
 Else
 Check1(Index).BackColor = &H8000000F
 End If
  
  
End If

For i = 1 To 33
 If Check1(i).Value = 1 Then
  NumC = NumC + 1
  
   TempStr = TempStr + Check1(i).Caption + " "
   Check1(i).BackColor = vbGreen
 
 
 'leaf add more
  If i Mod 2 Then           'The ODD and EVEN
   ODDN = ODDN + 1
  Else
   EVENN = EVENN + 1
  End If
  If i < 12 Then            'The Large Region
   SmaN = SmaN + 1
  ElseIf i > 21 Then
   BIGN = BIGN + 1
  Else
   MidN = MidN + 1
  End If
   SumStr = SumStr + i   'the sum of num
 
 If i > 0 And i < 6 Then
    RegionN(1) = RegionN(1) + 1
 End If
If i > 5 And i < 11 Then
    RegionN(2) = RegionN(2) + 1
 End If
If i > 10 And i < 16 Then
    RegionN(3) = RegionN(3) + 1
 End If
If i > 15 And i < 21 Then
    RegionN(4) = RegionN(4) + 1
 End If
If i > 20 And i < 26 Then
    RegionN(5) = RegionN(5) + 1
 End If
If i > 25 And i < 31 Then
    RegionN(6) = RegionN(6) + 1
 End If
If i > 30 And i < 34 Then
    RegionN(7) = RegionN(7) + 1
 End If
  
 
   LoseStr = LoseStr + Format(Str(TopLose(i)), "00") + " "   'the LostNum
 Else
   Check1(i).BackColor = &H8000000F
  
End If
  
 
  
  If NumC = 6 Then
  Label1.BackColor = &H80FF80
  RedOK = True
      If BlueOK Then
        Command2.Caption = "保存"
        Command2.BackColor = vbGreen
        If UpdateMode Then
            Check3.Value = 1
            Check3.Enabled = True
            Check3.BackColor = vbGreen
        End If
        End If
       
  Else
  RedOK = False
  Label1.BackColor = &HC0C0FF
          Command2.Caption = "待选"
        Command2.BackColor = &HC0C0FF
    
  End If
Next i
 

Label1.Caption = TempStr
'Label4.Caption = Str(SumStr)


'Show The Nums Feature
 Label5(0).Caption = Format(Str(ODDN), "0") + ":" + Format(Str(EVENN), "0")
 Label5(1).Caption = Format(Str(SmaN), "0") + ":" + Format(Str(MidN), "0") + ":" + Format(Str(BIGN), "0")
 Label5(2).Caption = Str(SumStr)
 
 TempStr = ""
 For i = 1 To 7
  TempStr = TempStr + Format(RegionN(i), "0") + ":"
 Next i
 Label5(3).Caption = TempStr
Label5(4).Caption = LoseStr
 

End Sub

Private Sub Check2_Click(Index As Integer)
Dim i As Integer
Dim BNum As Integer
 
If BlueOK Then
 If Check2(Index) = 1 Then
  Check2(Index).BackColor = vbRed
  Exit Sub
 Else
  Check2(Index).BackColor = &H8000000F
 End If
  
  
End If
 
 For i = 1 To 16
  If Check2(i).Value = 1 Then
   Label2.Caption = Check2(i).Caption
   BNum = BNum + 1
    Check2(i).BackColor = vbGreen
    Else
   Check2(i).BackColor = &H8000000F

 End If
 If BNum = 1 Then
    BlueOK = True
    Label2.BackColor = &H80FF80
    If RedOK Then
        Command2.Caption = "保存"
        Command2.BackColor = vbGreen
        If UpdateMode Then
            Check3.Value = 1
            Check3.Enabled = True
            Check3.BackColor = vbGreen
        End If
   End If
  Else
   BlueOK = False
          Command2.Caption = "待选"
        Command2.BackColor = &HC0C0FF

    Label2.BackColor = &HFF8080
 
 
 End If
 
 Next i
End Sub

Private Sub Check3_Click()
UpdateMode = False
Check3.Visible = False
End Sub

Private Sub Command2_Click()
If BlueOK And RedOK Then
    List1.AddItem Label1.Caption + "+" + Label2.Caption
    List2.AddItem Label5(0).Caption + "|" + Label5(1).Caption + "|" + Label5(2).Caption
End If
End Sub

Private Sub Command3_Click()
 Call ShellExecute(hwnd, "open", "http://paipai.500wan.com/", vbNullString, vbNullString, 1)
End Sub

Private Sub Command4_Click()
Dim i As Integer
BlueOK = False
RedOK = False

 For i = 1 To 33
  Check1(i).Value = 0
 Next i
 For i = 1 To 16
  Check2(i).Value = 0
Next i
Label1.Caption = "待选号码"
Label2.Caption = "B"

End Sub

Private Sub Command5_Click()
Dim i As Integer
Dim tmpstr As String
Open App.Path + "\savet.txt" For Append As #1
Print #1, Format(NumStr, "00000") + "期购买号码"
For i = 0 To List1.ListCount - 1
    tmpstr = List1.List(i) + List2.List(2)
 Print #1, tmpstr
Next i
'Print #1, Text1.Text
Close #1
End Sub

Private Sub Command6_Click()
If Check3.Visible = False Then
  Check3.Visible = True
  UpdateMode = True
  Check3.Caption = "Update"
   
End If

If Check3.Value = 1 Then

Open App.Path + "\all.txt" For Append As #1
 ' Print #1, vbCrLf
  Print #1, Format(NumStr, "00000") + " " + Label1.Caption + Label2.Caption
Close #1
UpdateMode = False
Check3.Visible = False
End If

End Sub

Private Sub Form_Load()
Form4.Caption = Format(NumStr, "00000") + Form4.Caption
End Sub

Private Sub List1_DblClick()
List2.RemoveItem (List1.ListIndex)
List1.RemoveItem (List1.ListIndex)

End Sub
