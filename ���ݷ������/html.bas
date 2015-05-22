Attribute VB_Name = "Module2"
Option Explicit


'Appear Red Ball <TD bgColor=#99FF33>1</TD>
'Appear Blue Ball <TD bgColor="#FF00FF">12</TD>
'Not Appear Ball <TD></TD>
'行首   <TR onMouseOver="this.style.backgroundColor='#ffff00'" onmouseout="this.style.backgroundColor=''">
'行末 </TR>
Public Const TRHead = "<TR onMouseOver=""this.style.backgroundColor='#ffff00'"" onmouseout=""this.style.backgroundColor=''"">"
Public Const TREnd = "</TR>"
Public Const SRTDHead = "<TD bgColor=#99FF33>"
Public Const SBTDHead = "<TD bgColor=#FF00FF>"
Public Const TDEnd = "</TD>"
Public Const HtmlHeadFile = "DB1.html"
Public Const HtmlEndFile = "DB2.html"
Public Const DivHead = "<TD><DIV align=center>"
Public Const DivEnd = "</DIV></TD>"
Public Const NOTD = "<TD></TD>"
Public Const AddTD = "<TD class=p5></TD>"
Public Const HeadFile = "/html/DB1.html"
Public Const EndFile = "/html/DB2.html"

Public Const Blue1 = "<TD>RR1</TD><TD class=p1><B>RR2</B></TD><TD class=p2><B>RR3</B></TD>"
Public Const P3R = "<TD class=p3>RR</TD>"
Public Const p4R = "<TD class=p4>RR</TD>"
Public Const p5R = "<TD class=p5>RR</TD>"






Public Function WriteRTR(IStr As String) As String
Dim LineStr As String
Dim i As Integer
Dim Rnum(33) As String
Rnum(0) = Left(IStr, 5)
For i = 1 To 33
 Rnum(i) = NOTD
Next i
For i = 1 To 6
    Rnum(Val(Mid(IStr, 2 * i + 4, 2))) = SRTDHead + Mid(IStr, 2 * i + 4, 2) + TDEnd
Next i
LineStr = TRHead + DivHead + Rnum(0) + DivEnd
For i = 1 To 33
   LineStr = LineStr + Rnum(i)

    If (i Mod 5 = 0) Then
         LineStr = LineStr + AddTD
    End If
Next i
WriteRTR = LineStr

End Function

Public Function WriteRegTR(IStr As String) As String        ''IStr="10100|2111100|420|3210|2:4|3:3|
Dim RegStr(11) As String
Dim RN(7) As Integer
Dim LineStr As String
Dim i As Integer
RegStr(0) = "<TD class=p0>" + Left(IStr, 5) + "</TD>"

For i = 1 To 7
 RN(i) = Val(Mid(IStr, 6 + i, 1))
Next i

RegStr(1) = "<TD class=p1 >" + CheckXX(RN(1)) + "</TD>"
RegStr(2) = "<TD class=p2 >" + CheckXX(RN(2)) + "</TD>"
RegStr(3) = "<TD class=p1 >" + CheckXX(RN(3)) + "</TD>"
RegStr(4) = "<TD class=p2 >" + CheckXX(RN(4)) + "</TD>"
RegStr(5) = "<TD class=p1 >" + CheckXX(RN(5)) + "</TD>"
RegStr(6) = "<TD class=p2 >" + CheckXX(RN(6)) + "</TD>"
RegStr(7) = "<TD class=p1 >" + CheckXX(RN(7)) + "</TD>"

'' Regstr8  3区分布  4 区分布  大小比 ，奇偶比，
RegStr(8) = "<TD class=p3 >" + Mid(IStr, 15, 3) + "</TD>"   '3区分布
RegStr(9) = "<TD class=p4 >" + Mid(IStr, 19, 4) + "</TD>"   '4
RegStr(10) = "<TD class=p3 >" + Mid(IStr, 24, 3) + "</TD>"  '大小比
RegStr(11) = "<TD class=p4 >" + Mid(IStr, 28, 3) + "</TD>"  '奇偶比

For i = 0 To 11
WriteRegTR = WriteRegTR + RegStr(i)
Next i
WriteRegTR = "<TR>" + WriteRegTR + "</TR>"


End Function
Public Function CheckXX(i As Integer) As String
Select Case i
Case 0:
   CheckXX = ""
Case 1:
    CheckXX = "★"
Case 2:
    CheckXX = "★★"
Case 3:
    CheckXX = "★★★"
Case 4:
    CheckXX = "★★★★"
End Select




End Function

Public Function WriteLostTR(IStr As String) As String
Dim WLost(28) As String
Dim i As Integer
WLost(0) = "<TD class=p0>" + Left(IStr, 3) + "</TD>"
For i = 1 To (Len(IStr) - 3) \ 2
    If (i - 1) \ 4 = 0 Or (i - 1) \ 4 = 4 Then
        WLost(i) = "<TD class=p1>" + Mid(IStr, 2 * i - 1 + 3, 2) + "</TD>"
    End If
    If (i - 1) \ 4 = 1 Or (i - 1) \ 4 = 5 Then
        WLost(i) = "<TD class=p2>" + Mid(IStr, 2 * i - 1 + 3, 2) + "</TD>"
    End If
    If (i - 1) \ 4 = 2 Or (i - 1) \ 4 = 6 Then
        WLost(i) = "<TD class=p3>" + Mid(IStr, 2 * i - 1 + 3, 2) + "</TD>"
    End If
    If (i - 1) \ 4 = 3 Or (i - 1) \ 4 = 7 Then
        WLost(i) = "<TD class=p4>" + Mid(IStr, 2 * i - 1 + 3, 2) + "</TD>"
    End If
Next i
For i = (Len(IStr) - 3) \ 2 To 28
        If (i - 1) \ 4 = 0 Or (i - 1) \ 4 = 4 Then
        WLost(i) = "<TD class=p1>" + "</TD>"
    End If
    If (i - 1) \ 4 = 1 Or (i - 1) \ 4 = 5 Then
        WLost(i) = "<TD class=p2>" + "</TD>"
    End If
    If (i - 1) \ 4 = 2 Or (i - 1) \ 4 = 6 Then
        WLost(i) = "<TD class=p3>" + "</TD>"
    End If
    If (i - 1) \ 4 = 3 Or (i - 1) \ 4 = 7 Then
        WLost(i) = "<TD class=p4>" + "</TD>"
    End If
Next i
For i = 0 To 28
  WriteLostTR = WriteLostTR + WLost(i)
Next i
WriteLostTR = "<TR>" + WriteLostTR + "</TR>" + vbCrLf



End Function



Public Function CreateHtmlFile(OutString As String, OutFileName As String, RID As Integer)
Dim RString As String
Dim bytImage()     As Byte
If Dir(OutFileName) <> "" Then
    Kill OutFileName
End If

bytImage() = LoadResData(RID, "CUSTOM")
RString = StrConv(bytImage(), vbUnicode)
RString = Replace(RString, "RRR", OutString)


Open OutFileName For Append As #1
   Print #1, RString
Close #1


End Function
Public Function WBlue1(IStr As String) As String
Dim tp1, tp2, tp3 As String
Dim i As Integer
tp1 = Left(IStr, 5)
For i = 1 To 6
 tp2 = tp2 + Mid(IStr, 5 + 2 * i - 1, 2) + " "
Next i
tp3 = Right(IStr, 2)
WBlue1 = Replace(Blue1, "RR1", tp1)
WBlue1 = Replace(WBlue1, "RR2", tp2)
WBlue1 = Replace(WBlue1, "RR3", tp3)





End Function
Public Function WBlue2(BLuenum As Integer) As String
Dim i As Integer
Dim tempstr As String
For i = 1 To 16
   If i = BLuenum Then
     tempstr = tempstr + Replace(p4R, "RR", i)
   Else
     tempstr = tempstr + Replace(P3R, "RR", "")
   End If
    
Next i
WBlue2 = tempstr
End Function
Public Function WBlue3(BLuenum As Integer) As String          ''◆★☆▲■
Dim tp1, tp2 As String
Dim tp3(4) As String
Dim i As Integer
If (BLuenum Mod 2) Then
 tp1 = Replace(p5R, "RR", "▲")
 tp1 = tp1 + Replace(p5R, "RR", "")
Else
  tp1 = Replace(p5R, "RR", "")
 tp1 = tp1 + Replace(p5R, "RR", "▲")
End If

If (BLuenum < 8) Then
 tp2 = Replace(P3R, "RR", "◆")
 tp2 = tp2 + Replace(P3R, "RR", "")
Else
 tp2 = Replace(P3R, "RR", "")
 tp2 = tp2 + Replace(P3R, "RR", "◆")
End If
For i = 1 To 4
 tp3(i) = Replace(p5R, "RR", "")
Next i
  tp3(BLuenum \ 4) = Replace(p5R, "RR", "★")
 
WBlue3 = tp1 + tp2 + tp3(1) + tp3(2) + tp3(3) + tp3(4)

End Function
Public Function WriteBLueAll(IStr As String) As String

Dim BLuenum As Integer
BLuenum = Val(Right(IStr, 2))
WriteBLueAll = "<TR>" + WBlue1(IStr) + WBlue2(BLuenum) + WBlue3(BLuenum) + "</TR>"


End Function


