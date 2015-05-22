VERSION 5.00
Begin VB.Form MTime 
   Caption         =   "中期数据统计"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   14385
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List3 
      Height          =   2940
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   7080
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   11175
   End
   Begin VB.ListBox List1 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "MTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalNum As Integer
Private Function ShowMTime()
Dim SaveData(1 To 1000, 1 To 33) As Integer
Dim LoseNum(1 To 1000, 1 To 33) As Integer
Dim tdata(7) As String
Dim NumStr(1 To 33) As String
Dim NumTimes As Integer
Dim tempstr As String
Dim i, j As Integer
For i = 1 To TotalNum
    tempstr = List1.List(i - 1)
    Call GetNum(tempstr, tdata())
 For j = 1 To 6
    SaveData(i, Val(tdata(j))) = 1
 Next j
Next i
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
tempstr = ""
    For j = 1 To 33
       NumStr(j) = Format(Str(LoseNum(TotalNum, j)), "00") + "| "
    For i = TotalNum To 2 Step -1
            If LoseNum(i, j) = 0 Then
               NumTimes = NumTimes + 1
                NumStr(j) = Format(Str(LoseNum(i - 1, j)), "00") + "| " + NumStr(j)
            End If
          
            If NumTimes > 10 Then
                NumTimes = 0
                Exit For
            End If
            
           
    Next i
           List2.AddItem Format(Str(j), "00") + ": |" + NumStr(j)
        
    Next j
    
    
    
  ' saveData(i,val(mid(tempstr,j


End Function

Private Sub Form_Load()
Dim i As Integer
TotalNum = Form2.AllData.ListCount
For i = 0 To Form2.AllData.ListCount - 1
    List1.AddItem Form2.AllData.List(i)
Next i
Call ShowMTime
For i = 0 To Form2.yllist.ListCount - 1
    List3.AddItem Form2.yllist.List(i)
Next i
List1.ListIndex = List1.ListCount - 1
End Sub
