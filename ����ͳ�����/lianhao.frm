VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "连号统计"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form3"
   ScaleHeight     =   8010
   ScaleWidth      =   10680
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List3 
      Height          =   7080
      Left            =   6840
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   7080
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ShowList2()
Dim i, j As Integer
Dim Diff(5) As Integer
Dim TempStr, TempZ As String
For i = 0 To List1.ListCount - 1
 TempStr = Right(Replace(List1.List(i), " ", ""), 14)
  For j = 1 To 5
   Diff(j) = Val(Mid(TempStr, 2 * j + 1, 2)) - Val(Mid(TempStr, 2 * j - 1, 2)) - 1
  Next j
  TempStr = ""
  For j = 1 To 5
   TempStr = TempStr + Format(Str(Diff(j)), "00") + " "
   If Diff(j) = 0 Then
     TempZ = TempZ + " Z"
   End If
  Next j
  List2.AddItem TempStr + "|" + TempZ
  TempZ = ""
Next i

  



End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To Form2.AllData.ListCount - 1
    List1.AddItem Form2.AllData.List(i)
Next i
Call ShowList2
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex

End Sub
