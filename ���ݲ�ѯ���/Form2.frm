VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   ScaleHeight     =   8430
   ScaleWidth      =   10560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   6000
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim i As Integer
   
Open App.Path & "/00.txt" For Append As #1

    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
    Next i
Close #1

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim pos(7) As Integer
Dim tempstr, newStr As String
Open App.Path & "/03.txt" For Input As #1
While Not EOF(1)

Line Input #1, tempstr


pos(0) = InStr(tempstr, " ")
For i = 1 To 7
    pos(i) = InStr(pos(i - 1) + 1, tempstr, " ")

Next i
    newStr = Mid(tempstr, 1, pos(0) - 1)
For i = 0 To 6
    If (pos(i + 1) - pos(i)) = 2 Then
        newStr = newStr + " " + "0" + Mid(tempstr, pos(i) + 1, 1)
   Else
        newStr = newStr + " " + Mid(tempstr, pos(i) + 1, 2)
   End If
Next i
List1.AddItem newStr
newStr = ""

'i = i + 1
Wend
Close #1
End Sub

