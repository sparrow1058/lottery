VERSION 5.00
Begin VB.Form BlueBall 
   Caption         =   "篮球数据统计"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form3"
   ScaleHeight     =   8205
   ScaleWidth      =   10995
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List4 
      Height          =   3840
      Left            =   2280
      TabIndex        =   3
      Top             =   3840
      Width           =   3615
   End
   Begin VB.ListBox List3 
      Height          =   3840
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Height          =   7800
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7440
      Width           =   1575
   End
End
Attribute VB_Name = "BlueBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BlueNum(500) As Integer
Dim AllNum As Integer
Private Sub Form_Load()
Dim i, j, k As Integer
Dim TempStr As String
Dim TempLine As String
Dim ShowStr1, ShowStr2, ShowStr3 As String
Dim ShowOE, ShowRegion As String
Dim Region(4), RegionAll(4) As Integer
'ShowStr = "  ★●◆■▲ |"
ShowStr1 = "    |"
ShowStr2 = "    |" + "    |"
ShowStr3 = "    |" + "    |" + "    |"
'ShowStr = " ★ |"
ShowRegion = " ★ |"
ShowOE = " ● |"
Dim LostBlue(500, 16) As Integer
AllNum = Form2.AllData.ListCount - 1
For i = 0 To AllNum
    TempStr = Form2.AllData.List(i)
    List1.AddItem Left(TempStr, 5) + " " + Right(TempStr, 3)
    BlueNum(i) = Val(Trim(Right(TempStr, 3)))
  
Next i
For j = 1 To 16
    LostBlue(0, j) = LostBlue(0, j) + 1
    Next j
LostBlue(0, BlueNum(0)) = 0
List2.AddItem Str(BlueNum(0))
For i = 1 To AllNum
  For j = 1 To 16
    LostBlue(i, j) = LostBlue(i - 1, j) + 1
 
Next j
    LostBlue(i, BlueNum(i)) = 0
Next i

For i = 1 To AllNum
k = k + 1
If BlueNum(i) < 5 Then
    TempLine = ShowRegion + ShowStr3
    Region(1) = Region(1) + 1
ElseIf BlueNum(i) < 9 Then
    Region(2) = Region(2) + 1
    TempLine = ShowStr1 + ShowRegion + ShowStr2
ElseIf BlueNum(i) < 13 Then
    Region(3) = Region(3) + 1
    TempLine = ShowStr2 + ShowRegion + ShowStr1
Else
   Region(4) = Region(4) + 1
    TempLine = ShowStr3 + ShowRegion
End If

List3.AddItem TempLine
TempLine = ""
If k = 16 Then
    For j = 1 To 4
        TempLine = TempLine + Format(Str(Region(j)), " 00") + " |"
         RegionAll(j) = RegionAll(j) + Region(j)
        Region(j) = 0
    Next j
    List4.AddItem TempLine
    k = 0
    
End If
TempLine = ""
For j = 1 To 16
    'TempLine = TempLine + Str(LostBlue(i, j)) + "|"
    If LostBlue(i, j) = 0 Then
    List2.AddItem Format(Str(LostBlue(i - 1, j)), "00")
    End If
 Next j
 
Next i
If Not k Then
     For j = 1 To 4
        TempLine = TempLine + Format(Str(Region(j)), " 00") + " |"
         RegionAll(j) = RegionAll(j) + Region(j)
        Region(j) = 0
    Next j
    List4.AddItem TempLine
End If
List3.AddItem Format(Str(RegionAll(1)), "000") + " |" + Format(Str(RegionAll(2)), "000") + " |" + Format(Str(RegionAll(3)), "000") + " |" + Format(Str(RegionAll(4)), "000") + " |"
Label1.Caption = Str(List1.ListCount) + "/16 =" + Str(List1.ListCount / 16) + "--" + Str(List1.ListCount Mod 16)
End Sub
