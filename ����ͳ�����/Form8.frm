VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "随机生成器"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form8"
   ScaleHeight     =   8910
   ScaleWidth      =   12300
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List3 
      Height          =   270
      Left            =   7560
      Style           =   1  'Checkbox
      TabIndex        =   58
      Top             =   8520
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   240
      Left            =   4320
      TabIndex        =   56
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   735
      Left            =   0
      TabIndex        =   55
      Top             =   7800
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4320
      TabIndex        =   6
      Top             =   6480
      Width           =   7815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   5520
         TabIndex        =   57
         Text            =   "1021020"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "遗漏次数选择"
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   43
         Top             =   1320
         Width           =   5415
         Begin VB.CheckBox Check2 
            Caption         =   "10"
            Height          =   255
            Index           =   10
            Left            =   4920
            TabIndex        =   54
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "9"
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   53
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "8"
            Height          =   255
            Index           =   8
            Left            =   3960
            TabIndex        =   52
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "7"
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   51
            Top             =   240
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "6"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   50
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "5"
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   49
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "4"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   48
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "3"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   47
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   46
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   45
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Value           =   1  'Checked
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "大小比"
         Height          =   615
         Left            =   1560
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton Option1 
            Caption         =   "0:6"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1:5"
            Height          =   255
            Index           =   8
            Left            =   720
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2:4"
            Height          =   255
            Index           =   9
            Left            =   1320
            TabIndex        =   40
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3:3"
            Height          =   255
            Index           =   10
            Left            =   1920
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "4:2"
            Height          =   255
            Index           =   11
            Left            =   2520
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "5:1"
            Height          =   255
            Index           =   12
            Left            =   3120
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "6:0"
            Height          =   255
            Index           =   13
            Left            =   3720
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "区间分布"
         Height          =   855
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   2415
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   0
            TabIndex        =   22
            Top             =   600
            Width           =   2415
            Begin VB.OptionButton Option2 
               Caption         =   "4"
               Height          =   255
               Index           =   9
               Left            =   1920
               TabIndex        =   27
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               Caption         =   "3"
               Height          =   255
               Index           =   8
               Left            =   1440
               TabIndex        =   26
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               Caption         =   "2"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               Caption         =   "1"
               Height          =   255
               Index           =   6
               Left            =   480
               TabIndex        =   24
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               Caption         =   "0"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.OptionButton Option2 
            Caption         =   "4"
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "3"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "长期出现次数分布"
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   4335
         Begin VB.CheckBox Check1 
            Caption         =   "7"
            Height          =   180
            Index           =   6
            Left            =   3240
            TabIndex        =   34
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6"
            Height          =   180
            Index           =   5
            Left            =   2760
            TabIndex        =   33
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5"
            Height          =   180
            Index           =   4
            Left            =   2280
            TabIndex        =   32
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4"
            Height          =   180
            Index           =   3
            Left            =   1800
            TabIndex        =   31
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3"
            Height          =   180
            Index           =   2
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2"
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   29
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Value           =   1  'Checked
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "奇偶比"
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   4335
         Begin VB.OptionButton Option1 
            Caption         =   "6:0"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "5:1"
            Height          =   255
            Index           =   5
            Left            =   3120
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "4:2"
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3:3"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2:4"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1:5"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0:6"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Image Image4 
         Height          =   420
         Left            =   7080
         Picture         =   "Form8.frx":0000
         Top             =   1440
         Width           =   420
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   1560
      Picture         =   "Form8.frx":0611
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   0
      Picture         =   "Form8.frx":5143
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   188
      Left            =   1560
      Top             =   5760
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4150
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   720
         Top             =   5400
      End
      Begin VB.Image Image1 
         Height          =   690
         Index           =   0
         Left            =   0
         Picture         =   "Form8.frx":9C56
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   2880
      Picture         =   "Form8.frx":A93A
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "期数据预测"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   3120
      Picture         =   "Form8.frx":AF4B
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label Label2 
      Caption         =   "数据类型统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FList As Integer
Dim LostR(10, 10) As Integer        ' The Lose region
Dim LostN(1 To 33) As Integer       'The Number's Lose
Dim LostMax(10) As Integer          ' The Lose Reion Num
Dim AppReg(1 To 33) As Integer
Dim DelFlag1(1 To 33) As Integer         'show del the nums of Region 测试选择符合区间号码 Leaf
Dim DelFlag2(1 To 33) As Integer         'show del the nums of the Lost Times
Dim LongR(7, 5) As Integer
'**************just for the Manual choose
Dim InStrS(7) As String

Dim UseManual As Boolean
Dim ODDS, BIGS As Integer
Dim LoseRegStr, DelLoseStr As String


Private Function ShowPic(PList As Integer, Data() As Integer)
Dim i As Integer
For i = 1 To 6
 Image1(PList * 6 + i).Picture = LoadResPicture(100 + Data(i), 0)
Next i


End Function

Private Function ChooseIN(InString As String, OutNum As Integer, Data() As String)

 Dim i, CNum As Integer
 Dim Temp(5) As String

CNum = Len(InString) \ 2
If CNum < OutNum Then
 MsgBox " Error  The Choose Num is long then In"
 Exit Function
End If
For i = 1 To CNum
  Temp(i) = Mid(InString, 2 * i - 1, 2)
Next i


'************************  choose 1
If OutNum = 1 Then
 For i = 1 To CNum
  Data(i) = Temp(i)
Next i
End If
'***************************** choose 2
If OutNum = 2 Then
 If CNum = 2 Then
  Data(1) = Temp(1) + Temp(2)
 End If
 If CNum = 3 Then
  Data(1) = Temp(1) + Temp(2)
  Data(2) = Temp(1) + Temp(3)
  Data(3) = Temp(2) + Temp(3)
 End If
 If CNum = 4 Then
  Data(1) = Temp(1) + Temp(2)
  Data(2) = Temp(1) + Temp(3)
  Data(3) = Temp(1) + Temp(4)
  Data(4) = Temp(2) + Temp(3)
  Data(5) = Temp(2) + Temp(4)
  Data(6) = Temp(3) + Temp(4)
End If

If CNum = 5 Then
  Data(1) = Temp(1) + Temp(2)
  Data(2) = Temp(1) + Temp(3)
  Data(3) = Temp(1) + Temp(4)
  Data(4) = Temp(1) + Temp(5)
  Data(5) = Temp(2) + Temp(3)
  Data(6) = Temp(2) + Temp(4)
  Data(7) = Temp(2) + Temp(5)
  Data(8) = Temp(3) + Temp(4)
  Data(9) = Temp(3) + Temp(5)
  Data(10) = Temp(4) + Temp(5)
End If
  
End If
'************************** choose 3
If OutNum = 3 Then
 If CNum = 3 Then
  Data(1) = Temp(1) + Temp(2) + Temp(3)
 End If
 
 If CNum = 4 Then
  Data(1) = Temp(1) + Temp(2) + Temp(3)
  Data(2) = Temp(1) + Temp(2) + Temp(4)
  Data(3) = Temp(1) + Temp(3) + Temp(4)
  Data(4) = Temp(2) + Temp(3) + Temp(4)
 End If
 
 If CNum = 5 Then
  Data(1) = Temp(1) + Temp(2) + Temp(3)
  Data(2) = Temp(1) + Temp(2) + Temp(4)
  Data(3) = Temp(1) + Temp(2) + Temp(5)
  Data(4) = Temp(1) + Temp(3) + Temp(4)
  Data(5) = Temp(1) + Temp(3) + Temp(5)
  Data(6) = Temp(1) + Temp(4) + Temp(5)
  Data(7) = Temp(2) + Temp(3) + Temp(4)
  Data(8) = Temp(2) + Temp(3) + Temp(5)
  Data(9) = Temp(2) + Temp(4) + Temp(5)
  Data(10) = Temp(3) + Temp(4) + Temp(5)
 End If
End If

End Function


Private Function ShowList2()

Dim i, j As Integer

Dim Data(10) As String
Dim TempStr, InString As String

For i = 1 To 7

Call ChooseIN(InStrS(i), Val(Mid(Text1.Text, i, 1)), Data())
 For j = 1 To 10
  If Not (Data(j) = "") Then
    TempStr = TempStr + Data(j) + " "
  End If
 Data(j) = ""
Next j
If Len(TempStr) = 28 Then
   List2.AddItem Left(TempStr, 14)
   List2.AddItem Mid(TempStr, 15, 14)
  
 End If

If Len(TempStr) = 30 Then
   List2.AddItem Left(TempStr, 15)
   List2.AddItem Mid(TempStr, 16, 15)
  
 End If
 If Len(TempStr) = 50 Then
   List2.AddItem Left(TempStr, 15)
   List2.AddItem Mid(TempStr, 16, 14)
   List2.AddItem Mid(TempStr, 31, 14)
   List2.AddItem Mid(TempStr, 46, 14)
 End If

 
If Len(TempStr) = 70 Then
   List2.AddItem Left(TempStr, 14)
   List2.AddItem Mid(TempStr, 15, 14)
   List2.AddItem Mid(TempStr, 29, 14)
   List2.AddItem Mid(TempStr, 43, 14)
   List2.AddItem Mid(TempStr, 57, 14)
 End If
 If Len(TempStr) < 16 Then
  List2.AddItem TempStr
End If
  List2.AddItem "----------------"
  TempStr = ""
Next i





End Function

Private Function ShowList1(Data() As Integer, TempStr As String)
Dim i, j As Integer
Dim LostStr, EVENStr, BIGStr, SumStr, AppRegStr, RegionStr, Reg3Str As String
Dim EVENN, BIGN, TheSum, RegionN(6), Reg3(2) As Integer

For i = 0 To 6
 RegionN(i) = 0
Next i

For i = 1 To 6
 If LostN(Data(i)) = 10 Then
 LostStr = LostStr + "L"
 Else
 LostStr = LostStr + Format(Str(LostN(Data(i))), "0")
 End If

If (Data(i) Mod 2) = 0 Then
  EVENN = EVENN + 1                 'the odd and even
End If
If Data(i) > 21 Then
 BIGN = BIGN + 1                    'The Big and Small
End If
TheSum = TheSum + Data(i)               ' the all sum
AppRegStr = AppRegStr + Format(AppReg(Data(i)), "0")        'the appear times
RegionN((Data(i) - 1) \ 5) = RegionN((Data(i) - 1) \ 5) + 1     'the region
Reg3((Data(i) - 1) \ 11) = Reg3((Data(i) - 1) \ 11) + 1


Next i                  'show the lost number

If Not UseManual Then

If EVENN = 6 Or EVENN = 0 Or BIGN = 6 Or BIGN = 0 Then
   FList = FList - 1
  Exit Function
End If



'******************************** the next is for Manual  Rule
Else
 If ODDS > 0 Then
   If Not (EVENN = 7 - ODDS) Then
    FList = FList - 1
    Exit Function
   End If
 End If
'If BIGS > 0 Then
 '  If Not (BIGN = BIGS - 1) Then
  '  FList = FList - 1
   ' Exit Function
   'End If
 'End If


End If



'************************************** For the Manual rule

For i = 0 To 6
 
 If RegionN(i) > 3 Then
      FList = FList - 1
  Exit Function
 End If
RegionStr = RegionStr + Format(RegionN(i), "0")

Next i


EVENStr = Format(Str(6 - EVENN), "0") + ":" + Format(Str(EVENN), "0")
BIGStr = Format(Str(BIGN), "0") + ":" + Format(Str(6 - BIGN), "0")
Reg3Str = Format(Str(Reg3(0)), "0") + ":" + Format(Str(Reg3(1)), "0") + ":" + Format(Str(Reg3(2)), "0")
SumStr = Format(Str(TheSum), "000")



TempStr = LostStr + "|" + EVENStr + "|" + BIGStr + "|" + Reg3Str + "|" + SumStr + "|" + AppRegStr + "|" + RegionStr

List1.AddItem TempStr
List1.AddItem "***************************************"





End Function


Private Sub LoadPic()
Dim i, j As Integer
For i = 1 To 48 Step 6
 
 For j = 0 To 5
 Load Image1(i + j)
 Image1(i + j).Left = j * 690
 Image1(i + j).Top = 690 * ((i - 1) / 6)
 Image1(i + j).Visible = True
 Next j
Next i



End Sub

Private Sub GetLostAndLongR()
Dim i, j As Integer
Dim TempStr As String
For i = 0 To 10
TempStr = Form7.List1.List(i)
TempStr = Replace(TempStr, " ", "")
For j = 1 To (Len(TempStr) - 3) / 2
  LostR(i, j) = Val(Mid(TempStr, 3 + j * 2 - 1, 2))
  LostN(Val(Mid(TempStr, 3 + j * 2 - 1, 2))) = i
Next j
LostMax(i) = j - 1
Next i

For i = 1 To 6
 TempStr = Form7.List3.List(i - 1)
 TempStr = Replace(TempStr, " ", "")
For j = 1 To 5
 LongR(i, j) = Val(Mid(TempStr, 3 + 2 * j - 1, 2))
 AppReg(Val(Mid(TempStr, 3 + 2 * j - 1, 2))) = i
Next j
Next i
 TempStr = Form7.List3.List(6)
TempStr = Replace(TempStr, " ", "")
For j = 1 To 3
 LongR(7, j) = Val(Mid(TempStr, 3 + 2 * j - 1, 2))
 AppReg(Val(Mid(TempStr, 3 + 2 * j - 1, 2))) = 7
Next j

End Sub

Private Function ManualGet2()



End Function
Private Function ManualGet(Data() As Integer)
 Dim i, j As Integer
 Dim TempStr1, TempStr2, TempStr3 As String
 Dim TData(1 To 6) As Integer
 Dim TempData(0 To 33) As Integer
 Dim TempNum(1 To 11) As Integer
 Dim SmallNum, MidNum, BigNum As Integer
 Dim DelRegion(1 To 10) As Integer
 Dim NumCount As Integer
  For i = 0 To 4
   If Option2(i).Value = True Then
     SmallNum = i
   End If
   If Option2(i + 5).Value = True Then
     MidNum = i
   End If
  Next i
 
  '************************************************  del the Region
 If Len(LoseRegStr) = 1 Then
      TempStr1 = Form7.List3.List(Val(Mid(LoseRegStr, 1, 1)) - 1)
      TempStr1 = Replace(TempStr1, " ", "")
     For i = 1 To 5
     TempData(Val(Mid(TempStr1, 2 * i + 2, 2))) = 1
     Next i
 End If
       
 If Len(LoseRegStr) = 2 Then
      TempStr1 = Form7.List3.List(Val(Mid(LoseRegStr, 1, 1)) - 1)
      TempStr2 = Form7.List3.List(Val(Mid(LoseRegStr, 2, 1)) - 1)
      TempStr1 = Replace(TempStr1, " ", "")
      TempStr2 = Replace(TempStr2, " ", "")
     For i = 1 To 5
       TempData(Val(Mid(TempStr1, 2 * i + 2, 2))) = 1
       TempData(Val(Mid(TempStr2, 2 * i + 2, 2))) = 1
     
     Next i
 End If

        '************************************************
For i = 1 To Len(DelLoseStr)
   For j = 1 To LostMax(Val(Mid(DelLoseStr, i, 1)))
        TempData(LostR(Val(Mid(DelLoseStr, i, 1)), j)) = 1
    Next j
Next i









 NumCount = 0
 For i = 1 To 11
  If TempData(i) = 0 Then
    TempNum(i) = i
    NumCount = NumCount + 1
  End If
 Next i
For i = 1 To SmallNum
  TData(i) = TempNum(Rnd * (NumCount - 1) + 1)
Next i
For i = 1 To 11
 TempNum(i) = 0
Next i
'******************************* For Small one
NumCount = 0
 For i = 12 To 22
  If TempData(i) = 0 Then
    TempNum(i - 11) = i
    NumCount = NumCount + 1
  End If
Next i
For i = SmallNum + 1 To SmallNum + MidNum
  TData(i) = TempNum(Rnd * (NumCount - 1) + 1)
Next i
For i = 1 To 11
 TempNum(i) = 0
Next i

'******************************* For Middle one

NumCount = 0
 For i = 23 To 33
  If TempData(i) = 0 Then
    TempNum(i - 22) = i
    NumCount = NumCount + 1
  End If
 Next i
 For i = SmallNum + MidNum + 1 To 6
  TData(i) = TempNum(Rnd * (NumCount - 1) + 1)
 Next i
'**************************************** For the big one


For i = 1 To 33
 TempData(i) = 0
Next i


For i = 1 To 6
 TempData(TData(i)) = 1
Next i
j = 0
For i = 1 To 33
 If TempData(i) = 1 Then
  j = j + 1
  Data(j) = i
 End If
Next i



End Function
Private Function GetSixNum(Data() As Integer)
Dim i, j, k, ALLM, Tm As Integer
Dim TempStr As String
Dim Temp(4, 3), TempM1, TempM2, TempM3, TempL As Integer
Dim TempMid(30) As Integer
Dim TempSix(6, 6) As Integer

Dim TempData(33) As Integer
For i = 0 To 3
 Temp(i, 1) = LostR(i, Rnd * LostMax(i))
 Temp(i, 2) = LostR(i, Rnd * LostMax(i))
 Temp(i, 3) = LostR(i, Rnd * LostMax(i))
 Next i
 'Temp4 = LostR(4, Rnd * LostMax(4))
ALLM = LostMax(4) + LostMax(5) + LostMax(6) + LostMax(7) + LostMax(8) + LostMax(9)

For i = 4 To 9
 For j = 1 To 10
  If Not (LostR(i, j) = 0) Then
   TempMid(k) = LostR(i, j)
    k = k + 1
   End If
 Next j
Next i

TempM1 = TempMid(Rnd * ALLM)
TempM2 = TempMid(Rnd * ALLM)
TempM3 = TempMid(Rnd * ALLM)

TempL = LostR(10, Rnd * LostMax(10))

'*************************              ' The First One Lost  0,0   >10 Mid
 TempSix(0, 1) = Temp(0, 1)
 TempSix(0, 2) = Temp(0, 2)
 TempSix(0, 3) = Temp(Rnd * 2 + 1, 1)
 TempSix(0, 4) = TempM2
 TempSix(0, 5) = TempM1
 TempSix(0, 6) = LostR(10, Rnd * LostMax(10))

'*************************              ' The Second  One Lost  1,1   >10 Mid
 TempSix(1, 1) = Temp(1, 1)
 TempSix(1, 2) = Temp(1, 2)
 TempSix(1, 3) = Temp(Rnd * 3, 3)
 TempSix(1, 4) = TempM2
 TempSix(1, 5) = TempM1
 TempSix(1, 6) = LostR(10, Rnd * LostMax(10))
'******************************************
'*************************              ' The Third  One Lost  2,2   >10 Mid
 TempSix(2, 1) = Temp(2, 1)
 TempSix(2, 2) = Temp(2, 2)
 TempSix(2, 3) = Temp(Rnd * 3, 3)
 TempSix(2, 4) = TempM2
 TempSix(2, 5) = TempM1
 TempSix(2, 6) = LostR(10, Rnd * LostMax(10))
'******************************************
'*************************              ' The 4 Lost  0,0   >10  No Mid
 TempSix(3, 1) = TempM2
 TempSix(3, 2) = Temp(0, 2)
 TempSix(3, 3) = Temp(Rnd * 2 + 1, 1)
 TempSix(3, 4) = Temp(Rnd * 2 + 1, 2)
 TempSix(3, 5) = TempM1
 TempSix(3, 6) = LostR(10, Rnd * LostMax(10))
'******************************************
'*************************              ' The 4 Lost  1,1   >10  No Mid
 TempSix(4, 1) = Temp(Rnd * 3, 1)
 TempSix(4, 2) = Temp(Rnd * 3, 1)
 TempSix(4, 3) = TempM1
 TempSix(4, 4) = TempM2
 TempSix(4, 5) = TempM3
 TempSix(4, 6) = LostR(10, Rnd * LostMax(10))
'******************************************
'*************************              ' The 4 Lost  0,0   >10  No Mid
 TempSix(5, 1) = Temp(Rnd * 3, 1)
 TempSix(5, 2) = Temp(Rnd * 3, 1)
 TempSix(5, 3) = Temp(Rnd * 3, 1)
 TempSix(5, 4) = TempM1
 TempSix(5, 5) = TempM2
 TempSix(5, 6) = LostR(10, Rnd * LostMax(10))
'******************************************

Tm = Rnd * 5
j = 0
For i = 1 To 6
 TempData(TempSix(Tm, i)) = 1
Next i
For i = 1 To 33
 If TempData(i) = 1 Then
   j = j + 1
   Data(j) = i
 End If
Next i
'Label1.Caption = Label1.Caption + Str(Tm) + " "

End Function


Private Sub Check1_Click(Index As Integer)
'UseManual = True
Dim i, j As Integer
For i = 1 To 33
 DelFlag1(i) = 0
Next i

For i = 0 To 5
If Check1(i).Value = 0 Then
 For j = 1 To 5
  DelFlag1(LongR(i + 1, j)) = 1
 Next j
End If
Next i
 If Check1(6).Value = 0 Then
   For j = 1 To 3
    DelFlag1(LongR(7, j)) = 1
   Next j
 End If

Call Image2_DblClick
End Sub

Private Sub Check2_Click(Index As Integer)
Dim i, j As Integer
''delflag2 show the del numbers
For i = 1 To 33
 DelFlag2(i) = 0
Next i
For i = 0 To 10
 If Check2(i).Value = 0 Then
  For j = 1 To 33
   If LostN(j) = i Then
    DelFlag2(j) = 1
   End If
  Next j
 End If
Next i
Call Image2_DblClick

End Sub

Private Sub Command1_Click()
List1.Visible = True
List2.Visible = False
List3.Visible = False

UseManual = False
'Label1.Caption = ""

Dim i As Integer
List1.Clear
For i = 1 To 48
 Image1(i).Picture = LoadResPicture(100, 0)
Next i
Timer1.Enabled = True
Timer2.Enabled = True

End Sub

Private Sub Command2_Click()
Dim i As Integer
List1.Visible = True
List2.Visible = False
List3.Visible = False
List1.Clear
ODDS = 0
BIGS = 0
UseManual = True
 For i = 0 To 6
  If Option1(i).Value = True Then
   ODDS = i + 1
  End If
  If Option1(i + 7).Value = True Then
   BIGS = i + 1
  End If
 Next i
 LoseRegStr = ""
For i = 0 To 6
 If Check1(i).Value = 0 Then
  LoseRegStr = LoseRegStr + Format(Str(i + 1), "0")
 End If
Next i

 DelLoseStr = ""
 For i = 0 To 9
  If Check2(i).Value = 0 Then
   DelLoseStr = DelLoseStr + Format(Str(i), "0")
 End If
Next i




For i = 1 To 48
 Image1(i).Picture = LoadResPicture(100, 0)
Next i
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True

End Sub

Private Sub Form_Load()
List2.Top = List1.Top
List2.Left = List1.Left
List2.Height = List1.Height
List2.Width = 1800
List3.Top = List2.Top
List3.Left = List2.Left + List2.Width
List3.Height = List2.Height
List3.Width = List1.Width - 1800

Label2.Caption = "遗漏次数|奇偶|大小|区间比|和值|出现分布|详细分布"
Me.Caption = Format(NumStr, "00000") + Me.Caption
Label1.Caption = Format(NumStr, "00000") + Label1.Caption
Call GetLostAndLongR
Call LoadPic
Randomize

List2.Visible = False
List3.Visible = False

End Sub

Private Sub Image2_DblClick()
' Call ShellExecute(hwnd, "open", "http://paipai.500wan.com/", vbNullString, vbNullString, 1)
List2.Clear
List1.Visible = False
List2.Visible = True
List3.Visible = True
Dim i, j As Integer
For i = 1 To 7
 InStrS(i) = ""
Next i

List3.Clear
Dim TempStr As String
j = 1
For i = 1 To 33
 If DelFlag1(i) = 0 And DelFlag2(i) = 0 Then
   TempStr = TempStr + Format(Str(i), "00") + " "
   InStrS(j) = InStrS(j) + Format(Str(i), "00")
   List3.AddItem Format(Str(i), "00") + ":" + Str(AppReg(i)) + "  " + Right(LoseForm.List6.List(i - 1), 29) 'leaf
   
   
   
   
  End If
 If (i Mod 5) = 0 Then
   List2.AddItem TempStr
    'TempStr = TempStr + "|" + vbCrLf
   TempStr = ""
    j = j + 1
 End If

Next i
List2.AddItem TempStr
List2.AddItem "================="
If Image4.Visible = True Then

    Call ShowList2
End If


End Sub

Private Sub Image3_DblClick()
Call ShellExecute(hwnd, "open", "http://paipai.500wan.com/", vbNullString, vbNullString, 1)
End Sub




Private Sub List1_DblClick()
List1.Clear
End Sub

Private Sub List2_DblClick()
'List2.Clear
End Sub

Private Sub List3_Click()
Dim i, j As Integer
Dim Data(6) As Integer
j = 1
 For i = 0 To List3.ListCount - 1
  If List3.Selected(i) Then
    Data(j) = i + 1
    j = j + 1
   End If
  If j = 7 Then
   Call ShowPic(1, Data())
   Exit Sub
  End If
 Next i '''leaf


End Sub

Private Sub Text1_Change()
Dim i As Integer
Dim Sum As Integer
If Len(Text1.Text) = 7 Then
 Image4.Visible = True
 For i = 1 To 7
  If Val(Mid((Text1.Text), i, 1)) > 3 Then
    Image4.Visible = False
  End If
   Sum = Val(Mid((Text1.Text), i, 1)) + Sum
 Next i
 If Not (Sum = 6) Then
    Image4.Visible = False
 End If
 
 
Else
 Image4.Visible = False
End If
If Len(Text1.Text) > 7 Then
 Text1.Text = ""
End If
If Image4.Visible = True Then
 Call Image2_DblClick
 

End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer


For i = 1 To 6
Image1(FList * 6 + i).Picture = LoadResPicture(100 + Rnd * 33, 0)
Next i
End Sub

Private Sub Timer2_Timer()
Dim Data(1 To 6) As Integer
Dim i As Integer
Timer1.Enabled = False
Dim TempStr As String
If Not UseManual Then
While ((Data(1) And Data(2) And Data(3) And Data(4) And Data(5) And Data(6)) = 0)
For i = 1 To 6
 Data(i) = 0
Next i
 Call GetSixNum(Data())
 Wend
Else
While ((Data(1) And Data(2) And Data(3) And Data(4) And Data(5) And Data(6)) = 0)
 For i = 1 To 6
  Data(i) = 0
 Next i
     Call ManualGet(Data())
 Wend
End If


Call ShowPic(FList, Data())
Call ShowList1(Data(), TempStr)






FList = FList + 1
If FList > 7 Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    FList = 0
    
   Exit Sub
End If
Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = False
Timer1.Enabled = False
End Sub
