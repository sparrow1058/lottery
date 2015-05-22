Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public NumStr As String
Public TopLose(1 To 33) As Integer
Type Str_Sum
   str1 As String
   Sum1 As Integer
 
End Type

Type RegAttr
    Reg3Attr As String
    Reg4Attr As String
    Reg7Attr As String
End Type

Type NumAttr
    BSAttr As String
    OEAttr As String
    Reg7Attr As String
    Reg3Attr As String
    Reg4Attr As String
    
    LostAttr As String
    LostSum As String
    AppearAttr As String
    
End Type
Private Function CheckBS(InputStr As String) As String
Dim i As Integer
Dim Big As Integer
 For i = 1 To 6
  If Val(Mid(InputStr, 2 * i - 1, 2)) > 16 Then
    Big = Big + 1
  End If
 Next i
 CheckBS = Format(Str(Big), "0") + Format(Str(6 - Big), "0")
End Function
Private Function CheckOE(InputStr As String) As String
Dim i As Integer
Dim ODD As Integer
 For i = 1 To 6
  If (Val(Mid(InputStr, 2 * i - 1, 2)) Mod 2) Then
    ODD = ODD + 1
  End If
 Next i
 CheckOE = Format(Str(ODD), "0") + Format(Str(6 - ODD), "0")
End Function
Private Function CheckReg(InputStr As String) As RegAttr
Dim i As Integer
Dim Reg(6) As Integer
 For i = 1 To 6
  Reg(Val(Mid(InputStr, 2 * i - 1, 2) - 1) \ 5) = Reg(Val(Mid(InputStr, 2 * i - 1, 2) - 1) \ 5) + 1

  
 Next i
For i = 0 To 6


Next i
End Function





Private Function GetNumAttr(InputStr As String) As NumAttr
 GetNumAttr.BSAttr = CheckBS(InputStr)
 GetNumAttr.OEAttr = CheckOE(InputStr)
 



End Function
Public Function GetNumOld(Lstr As String, Data() As String)
Dim Pos1, Pos2, pos3, pos4, pos5, pos6, pos7 As Integer
'If Len(Lstr) >= 26 Then
Pos1 = InStr(Lstr, " ")
Pos2 = InStr(Pos1 + 1, Lstr, " ")
pos3 = InStr(Pos2 + 1, Lstr, " ")
pos4 = InStr(pos3 + 1, Lstr, " ")
pos5 = InStr(pos4 + 1, Lstr, " ")
pos6 = InStr(pos5 + 1, Lstr, " ")
pos7 = InStr(pos6 + 1, Lstr, " ")

Data(0) = Mid(Lstr, 1, Pos1)
Data(1) = Mid(Lstr, Pos1 + 1, Pos2 - Pos1 - 1)
Data(2) = Mid(Lstr, Pos2 + 1, pos3 - Pos2 - 1)
Data(3) = Mid(Lstr, pos3 + 1, pos4 - pos3 - 1)
Data(4) = Mid(Lstr, pos4 + 1, pos5 - pos4 - 1)
Data(5) = Mid(Lstr, pos5 + 1, pos6 - pos5 - 1)
Data(6) = Mid(Lstr, pos6 + 1, pos7 - pos6 - 1)
Data(7) = Mid(Lstr, pos7 + 1, Len(Lstr) - pos7 - 1)

'End If
End Function
Public Function GetNum(Lstr As String, Data() As String)
Dim i As Integer


Lstr = Replace(Lstr, " ", "")

If Len(Lstr) = 19 Then
Data(0) = Left(Lstr, 5)
Lstr = Right(Lstr, Len(Lstr) - 5)

 For i = 1 To 7
  Data(i) = Mid(Lstr, i * 2 - 1, 2)
 Next i



End If



'End If
End Function


Public Function DaPaixu(InData() As Integer, NumData() As Integer) ', OutData As Variant)

Dim i, j As Integer
Dim tempmax, TempNum As Integer
'For i = LBound(InData()) To UBound(InData())
   ' OutData(0, i) = i
  '  OutData(1, i) = InData(i)


'Next i
For i = LBound(InData()) To UBound(InData())
    NumData(i) = i
Next i
For j = LBound(InData()) To UBound(InData()) - 1
For i = LBound(InData()) To UBound(InData()) - j
    If InData(i) < InData(i + 1) Then
        tempmax = InData(i + 1)
        InData(i + 1) = InData(i)
        InData(i) = tempmax
        TempNum = NumData(i + 1)
        NumData(i + 1) = NumData(i)
        NumData(i) = TempNum
    End If
Next i
Next j



End Function

Public Function PPX(InputStr As String, NumData() As Variant)
    Dim i, j As Integer
    Dim SLen As Integer
    Dim TData() As Integer
    SLen = Len(InputStr)
    For i = 1 To SLen / 2
        TData(i) = Val(Mid(InputStr, i * 2 - 1, 2))
    Next i
    
    For i = 1 To SLen / 2
        
        
    Next i
    

End Function
Public Function zuhe8(ChooseStr() As String, Result As Variant)

'Result 0
Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)

'Next i


End Function
Public Function zuhe9(ChooseStr() As String, Result As Variant)

Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)



End Function
Public Function zuhe10(ChooseStr() As String, Result As Variant)

Result(0, 0) = ChooseStr(0)
Result(0, 1) = ChooseStr(1)
Result(0, 2) = ChooseStr(2)
Result(0, 3) = ChooseStr(4)
Result(0, 4) = ChooseStr(6)
Result(0, 5) = ChooseStr(7)
'Result 1
Result(1, 0) = ChooseStr(0)
Result(1, 1) = ChooseStr(1)
Result(1, 2) = ChooseStr(3)
Result(1, 3) = ChooseStr(5)
Result(1, 4) = ChooseStr(6)
Result(1, 5) = ChooseStr(7)
'Result 2
Result(2, 0) = ChooseStr(0)
Result(2, 1) = ChooseStr(2)
Result(2, 2) = ChooseStr(3)
Result(2, 3) = ChooseStr(4)
Result(2, 4) = ChooseStr(5)
Result(2, 5) = ChooseStr(7)
'Result 3
Result(3, 0) = ChooseStr(1)
Result(3, 1) = ChooseStr(2)
Result(3, 2) = ChooseStr(3)
Result(3, 3) = ChooseStr(4)
Result(3, 4) = ChooseStr(5)
Result(3, 5) = ChooseStr(6)
End Function

Public Function SumCount(InData() As String, Result() As Str_Sum)
 Dim i, j, k As Integer
 Dim StrSum As Integer
 StrSum = 1
 For i = LBound(InData()) To UBound(InData())
    If Not (InData(i) = "") Then
    For j = i + 1 To UBound(InData())
           If InData(i) = InData(j) Then
                StrSum = StrSum + 1
                InData(j) = ""
           End If
     Next j
     Result(k).str1 = InData(i)
     Result(k).Sum1 = StrSum
     k = k + 1
     StrSum = 1
    End If
  Next i
            
            
End Function
