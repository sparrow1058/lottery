Attribute VB_Name = "Module1"
Option Explicit
Public Const SumRegNums = 6
Type DataAttrib
   BSAttrib As String
   OEAttrib As String
   R3Attrib As String
   R4Attrib As String
   R6Attrib As String
   APAttribs As String
   LostTimes As Integer
   AppearRegions As Integer
End Type
Type SumType
    TheSum  As String
    TheLostSum As String
    TheAppear As String
    TheReg3  As String
    TheReg4 As String


End Type

Type StrSum
    StrChr As String
    StrSum As Integer
End Type
Type RBMatrixType
    RBAll As Integer    'inlcude num nums
    RBIN  As String     'include num
    RBNOT As String     'not include num
    RBList(1 To 33) As Integer
End Type
Type MaxMin
 max As Integer
 min As Integer
End Type


Public Function GetAllCombin(StrFlag As String, OutStr() As String)
Dim i, j As Integer
Dim TempOutstr1(100) As String
Dim TempOutstr2(100) As String
Dim TempOutstr3(100) As String
Dim TempOutstr4(10000) As String
Dim TempOutstr5(10000) As String

Dim RegoutStr0(0 To 9) As String
Dim RegoutStr1(0 To 9) As String
Dim RegoutStr2(0 To 9) As String
Dim RegoutStr3(0 To 9) As String
Dim RegoutStr4(0 To 9) As String
Dim RegoutStr5(0 To 9) As String
Dim RegoutStr6(0 To 9) As String
Dim RegionAllStr(0 To 6) As String
For i = 0 To 5
    For j = 1 To 5
        RegionAllStr(i) = RegionAllStr(i) + Format(Str(5 * i + j), "00")
    Next j
Next i
    RegionAllStr(6) = "313233"


   Call CombinCheck(RegionAllStr(0), Val(Mid(StrFlag, 1, 1)), RegoutStr0())
   Call CombinCheck(RegionAllStr(1), Val(Mid(StrFlag, 2, 1)), RegoutStr1())
   Call CombinCheck(RegionAllStr(2), Val(Mid(StrFlag, 3, 1)), RegoutStr2())
   Call CombinCheck(RegionAllStr(3), Val(Mid(StrFlag, 4, 1)), RegoutStr3())
   Call CombinCheck(RegionAllStr(4), Val(Mid(StrFlag, 5, 1)), RegoutStr4())
   Call CombinCheck(RegionAllStr(5), Val(Mid(StrFlag, 6, 1)), RegoutStr5())
   Call CombinCheck(RegionAllStr(6), Val(Mid(StrFlag, 7, 1)), RegoutStr6())
   Call PlayList(RegoutStr0(), RegoutStr1(), TempOutstr1())
   Call PlayList(RegoutStr2(), RegoutStr3(), TempOutstr2())
   Call PlayList(RegoutStr4(), RegoutStr5(), TempOutstr3())
   
   Call PlayList(TempOutstr1(), TempOutstr2(), TempOutstr4())
   Call PlayList(TempOutstr3(), RegoutStr6(), TempOutstr5())
   Call PlayList(TempOutstr4(), TempOutstr5(), OutStr())








End Function
Public Function PlayList(InputStr1() As String, Inputstr2() As String, OutStr() As String)
Dim len1, len2 As Integer
Dim i, j, k As Integer
For i = 0 To UBound(InputStr1())
    If Not (InputStr1(i) = "") Then
        len1 = i
    End If
Next i
For i = 0 To UBound(Inputstr2())
    If Not (Inputstr2(i) = "") Then
        len2 = i
    End If
 Next i
 
 For i = 0 To len1
    For j = 0 To len2
        OutStr(k) = InputStr1(i) + Inputstr2(j)
        k = k + 1
    Next j
 Next i
 


End Function

Public Function CombinCheck(TempStr As String, num As Integer, OutStr() As String)
    'tempstr as 10 char lenght
  If Len(TempStr) = 10 Then
 Select Case num
    Case 0:
        OutStr(0) = ""
    Case 1:
        OutStr(0) = Mid(TempStr, 1, 2)
        OutStr(1) = Mid(TempStr, 3, 2)
        OutStr(2) = Mid(TempStr, 5, 2)
        OutStr(3) = Mid(TempStr, 7, 2)
        OutStr(4) = Mid(TempStr, 9, 2)
    Case 2:
        OutStr(0) = Mid(TempStr, 1, 4)
        OutStr(1) = Mid(TempStr, 1, 2) + Mid(TempStr, 5, 2)
        OutStr(2) = Mid(TempStr, 1, 2) + Mid(TempStr, 7, 2)
        OutStr(3) = Mid(TempStr, 1, 2) + Mid(TempStr, 9, 2)
        
        OutStr(4) = Mid(TempStr, 3, 4)
        OutStr(5) = Mid(TempStr, 3, 2) + Mid(TempStr, 7, 2)
        OutStr(6) = Mid(TempStr, 3, 2) + Mid(TempStr, 9, 2)
        
        OutStr(7) = Mid(TempStr, 5, 4)
        OutStr(8) = Mid(TempStr, 5, 2) + Mid(TempStr, 9, 2)
        
        OutStr(9) = Mid(TempStr, 7, 4)
        
  
    Case 3:
        OutStr(0) = Mid(TempStr, 1, 6)
        OutStr(1) = Mid(TempStr, 1, 4) + Mid(TempStr, 7, 2)
        OutStr(2) = Mid(TempStr, 1, 4) + Mid(TempStr, 9, 2)
        OutStr(3) = Mid(TempStr, 1, 2) + Mid(TempStr, 5, 4)
        OutStr(4) = Mid(TempStr, 1, 2) + Mid(TempStr, 7, 4)
        OutStr(5) = Mid(TempStr, 1, 2) + Mid(TempStr, 5, 2) + Mid(TempStr, 9, 2)
        
        
        OutStr(6) = Mid(TempStr, 3, 6)
        OutStr(7) = Mid(TempStr, 3, 2) + Mid(TempStr, 7, 4)
       
        OutStr(8) = Mid(TempStr, 3, 4) + Mid(TempStr, 9, 2)
        OutStr(9) = Mid(TempStr, 5, 6)
         
   End Select
   End If
   If Len(TempStr) = 6 Then
    Select Case num
        Case 0:
            OutStr(0) = ""
        Case 1:
            OutStr(0) = "31"
            OutStr(1) = "32"
            OutStr(2) = "33"
        Case 2:
            OutStr(0) = "3132"
            OutStr(1) = "3133"
            OutStr(2) = "3233"
   End Select
   
   End If
 End Function
Public Function CheckBSNum(INBS As String) As Integer
    Dim i As Integer
    Dim Big As Integer
    For i = 1 To 6
    If Val(Mid(INBS, 2 * i - 1, 2)) > 16 Then
        Big = Big + 1
    End If
    Next i
    CheckBSNum = Big
    
End Function
Public Function CheckOENum(INOE As String) As Integer
    Dim i As Integer
    Dim Odd As Integer
    For i = 1 To 6
    If (Val(Mid(INOE, 2 * i - 1, 2)) Mod 2) Then
        Odd = Odd + 1
    End If
    Next i
    CheckOENum = Odd
End Function

Public Function CheckBS(INBS As String) As String
    Dim i As Integer
    Dim Big, Small As Integer
    For i = 1 To 6
    If Val(Mid(INBS, 2 * i - 1, 2)) > 16 Then
        Big = Big + 1
     Else
        Small = Small + 1
    End If
    Next i
    
    CheckBS = Format(Str(Big), "0") + ":" + Format(Str(Small), "0")
 

End Function
Public Function CheckOE(INOE As String) As String
    Dim i As Integer
    Dim Odd, EVEN As Integer
    For i = 1 To 6
    If (Val(Mid(INOE, 2 * i - 1, 2)) Mod 2) Then
        Odd = Odd + 1
    Else
        EVEN = EVEN + 1
       
    End If
        CheckOE = Format(Str(Odd), "0") + ":" + Format(Str(EVEN), "0")
    Next i
End Function

Public Function CheckR6(inStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    For i = 0 To 6
        CheckR6 = CheckR6 + Format(Str(Reg(i)), "0")
    Next i
End Function
Public Function CheckR3(inStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    CheckR3 = CheckR3 + Format(Str(Reg(0) + Reg(1) + Reg(2)), "0")
    CheckR3 = CheckR3 + Format(Str(Reg(3) + Reg(4) + Reg(5)), "0")
    CheckR3 = CheckR3 + Format(Str(Reg(6)), "0")
End Function
Public Function CheckR4(inStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    CheckR4 = CheckR4 + Format(Str(Reg(0) + Reg(1)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(2) + Reg(3)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(4) + Reg(5)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(6)), "0")
End Function
Public Function CheckR11(inStr1 As String) As String
    Dim i As Integer
    Dim Reg(3) As Integer
    For i = 1 To 6
        Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 11) = Reg((Val(Mid(inStr1, 2 * i - 1, 2)) - 1) \ 11) + 1
    Next i
    For i = 0 To 2
    CheckR11 = CheckR11 + Format(Str(Reg(i)), "0")
    Next i
End Function

Public Function AddSum(InputStr() As String, OutStr() As StrSum)
 Dim i, j, Out, Total As Integer
 Dim TempStr As String
 Out = 0
 For i = LBound(OutStr()) To UBound(OutStr())
    OutStr(i).StrChr = ""
    OutStr(i).StrSum = 0
Next i
 
' Total = UBound(InStr()) - LBound(InStr())
  For i = LBound(InputStr()) To UBound(InputStr())
   If Not (InputStr(i) = "") Then
    OutStr(Out).StrChr = InputStr(i)
    OutStr(Out).StrSum = 1
    For j = i + 1 To UBound(InputStr())
        If (InputStr(i) = InputStr(j)) Then
            OutStr(Out).StrSum = OutStr(Out).StrSum + 1
            InputStr(j) = ""
        End If
    Next j

    Out = Out + 1
  End If
   
   Next i
   
End Function
Public Function CheckAppearRegions(InputStr As String, SumRegion() As Integer, MaxRegStr As String) As Integer
    Dim i As Integer
    Dim max As Integer
    Dim Nums As Integer
    Dim region(1 To SumRegNums) As Integer
    For i = 1 To SumRegNums
       'leaf
       region(SumRegion(Val(Mid(InputStr, 2 * i - 1, 2)))) = region(SumRegion(Val(Mid(InputStr, 2 * i - 1, 2)))) + 1
       
    Next i

    For i = 1 To SumRegNums
        If region(i) = 0 Then
            Nums = Nums + 1
        End If
        If max < region(i) Then
            max = region(i)
            MaxRegStr = Str(i)
        End If
    Next i

    MaxRegStr = MaxRegStr + "--" + Str(max)
        CheckAppearRegions = SumRegNums - Nums
End Function
Public Function SumSort(InputData() As StrSum)
   Dim i, j As Integer
   Dim TempStr As StrSum
   For i = LBound(InputData()) To UBound(InputData())
       For j = UBound(InputData()) To i + 1 Step -1
            If InputData(j).StrSum > InputData(j - 1).StrSum Then
                TempStr.StrChr = InputData(j - 1).StrChr
                TempStr.StrSum = InputData(j - 1).StrSum
                InputData(j - 1).StrChr = InputData(j).StrChr
                InputData(j - 1).StrSum = InputData(j).StrSum
                InputData(j).StrChr = TempStr.StrChr
                InputData(j).StrSum = TempStr.StrSum
            End If
        Next j
    Next i
    
End Function
Public Function CheckAttrib(InputStr As String, OutAttrib As DataAttrib)
    OutAttrib.BSAttrib = CheckBS(InputStr)
    OutAttrib.OEAttrib = CheckOE(InputStr)
    OutAttrib.R3Attrib = CheckR3(InputStr)
    OutAttrib.R6Attrib = CheckR6(InputStr)
    
    
End Function
Public Function RBMatrix(InData() As String, OutData As RBMatrixType)
    Dim i, j As Integer
    Dim RBData(1 To 33) As Integer
    For i = 1 To 6
        If Len(InData(i)) = 12 Then
            For j = 1 To 6
                OutData.RBList(Val(Mid(InData(i), 2 * i - 1, 2))) = OutData.RBList(Val(Mid(InData(i), 2 * i - 1, 2))) + 1
            Next j
        End If
    Next i
End Function
Public Function AddAllSum(TempStr As String, OutSum() As Integer)
Dim i As Integer
For i = 1 To 6
    OutSum(Val(Mid(TempStr, 2 * i - 1, 2))) = OutSum(Val(Mid(TempStr, 2 * i - 1, 2))) + 1
Next i
End Function
Public Function GetNewData(Data As String) As String

Dim i, cpos As Integer
Dim tnumstr, LineStr As String
cpos = InStr(Data, "<DIValign=center>")
LineStr = Mid(Data, cpos + 17, 5) + " "
For i = 1 To 6
    cpos = InStr(cpos + 1, Data, "<TDbgColor=#FF33>")
    tnumstr = Mid(Data, cpos + 18, 2)
 If InStr(tnumstr, "<") Then
    tnumstr = "0" + Left(tnumstr, 1)
 End If
    LineStr = LineStr + tnumstr + " "

Next i
 cpos = InStr(cpos + 1, Data, "FF00FF")
 tnumstr = Mid(Data, cpos + 8, 2)
 If InStr(tnumstr, "<") Then
    tnumstr = "0" + Left(tnumstr, 1)
 End If
    LineStr = LineStr + tnumstr + " "
    GetNewData = LineStr
End Function
Public Function GetSumRegion(SumNumber() As Integer, SumRegion() As Integer)
Dim i, j, RegSize As Integer
Dim max, min As Integer
max = SumNumber(1)
min = max
For i = 2 To 33
    If SumNumber(i) > max Then
        max = SumNumber(i)
    End If
    If SumNumber(i) < min Then
        min = SumNumber(i)
    End If
Next i

RegSize = (max - min + 1) / SumRegNums
For i = 1 To 33
 For j = 1 To SumRegNums
    If SumNumber(i) >= min + (j - 1) * RegSize And SumNumber(i) < min + (j + 1) * RegSize Then
        SumRegion(i) = j
    End If
 Next j
Next i

End Function
Public Function GetLost(SumList() As String, FindSum As String) As String
Dim i, j As Integer
Dim TempStr As String
TempStr = FindSum + "- "
For i = UBound(SumList()) To LBound(SumList()) Step -1
   If Not (SumList(i) = "") Then
    j = j + 1
    If SumList(i) = FindSum Then
        TempStr = TempStr + Format(Str(j - 1), "00") + " "
        j = 0
   End If
   End If
Next i
GetLost = TempStr

End Function
Public Function GetMaxMin(InNum() As Integer) As MaxMin
Dim i As Integer
Dim max, min As Integer
min = InNum(2)
For i = LBound(InNum()) To UBound(InNum())
    If Not (InNum(i) = 0) Then
        If max < InNum(i) Then
            max = InNum(i)
        End If
        If min > InNum(i) Then
            min = InNum(i)
        End If
    End If
Next i
GetMaxMin.max = max
GetMaxMin.min = min

End Function
Public Function FindLost(InNum() As String, CheckNUm As String, Outnum() As Integer) As Integer
 Dim i, k As Integer
 For i = LBound(Outnum()) To UBound(Outnum())
  Outnum(i) = 0
Next i
 
 For i = LBound(InNum()) To UBound(InNum())
    If Not (InNum(i) = "") Then
      If InNum(i) = CheckNUm Then
        k = k + 1
        Outnum(k) = 0
      Else
        Outnum(k) = Outnum(k) + 1
      End If
    End If
    
 Next i
 FindLost = k
End Function
Public Function FindNowLost(InNum() As String, Pos As Integer) As Integer
Dim i As Integer
 For i = Pos + 1 To UBound(InNum())
       If InNum(i) = InNum(Pos) Then
        FindNowLost = i - Pos - 1
      Exit For
     End If
     
 Next i

End Function

Public Function ChangeStr(IStr As String) As String
Dim i  As Integer
Dim Count As Integer
Dim Pos As Integer
Dim Pleft, Pright As Integer
Dim PLen As Integer
Dim TempStr As String
PLen = Len(IStr)
Pos = InStr(IStr, "10")
Pleft = Pos - 1
Pright = PLen - Pos - 1
If Pos = 0 Then
    Exit Function
End If
Count = 0
For i = 1 To Pleft
 If (InStr(i, IStr, "1") = i) Then
    Count = Count + 1
 End If
Next i
TempStr = ""
For i = 1 To Count
    TempStr = TempStr + "1"
Next i
For i = Count + 1 To Pleft
    TempStr = TempStr + "0"
Next i
TempStr = TempStr + "01"
TempStr = TempStr + Right(IStr, Pright)

ChangeStr = TempStr
End Function
Public Function CCInOut(InNum As Integer, Outnum As Integer) As Integer
Dim Temp1, Temp2 As Long
Temp1 = 1
Temp2 = 1
Dim i As Integer
For i = 1 To Outnum
 Temp1 = Temp1 * (InNum - i + 1)
 Next i
For i = 1 To Outnum
 Temp2 = Temp2 * i
Next i

CCInOut = Temp1 \ Temp2
End Function

Public Function CombinNum(IStr As String, FlagStr As String) As String          ' 从字符串中取出 字符
Dim TotalNum As Integer
Dim LenTemp As Integer
Dim TempStr As String
Dim i As Integer
If Len(FlagStr) = Len(IStr) \ 2 Then
 For i = 1 To Len(FlagStr)
    If (InStr(i, FlagStr, "1") = i) Then
        TempStr = TempStr + Mid(IStr, 2 * i - 1, 2)
    End If
Next i
End If
CombinNum = TempStr
End Function
Public Function FindLHSum(IStr() As String, FindStr As String) As Integer

    Dim i As Integer
    Dim Fcount As Integer
    For i = LBound(IStr) To UBound(IStr)
        If IStr(i) = "" Then
            Exit For
        End If
        If InStr(IStr(i), FindStr) Then
            Fcount = Fcount + 1
        End If
        
    Next i
FindLHSum = Fcount
End Function

Public Function TheMainValue(IStr() As String, OStr() As String)
  Dim i, j As Integer
Dim ICount(33) As Integer
Dim TempStr As String
Dim INums  As Integer
INums = 4

'*************************************************************计算5期内， 数字出现0 - 5次的次数，
For i = 0 To 5
    For j = 1 To 6
        ICount(Val(Mid(IStr(i), 2 * j - 1, 2))) = ICount(Val(Mid(IStr(i), 2 * j - 1, 2))) + 1
    Next j
Next i
For i = 1 To 33
    OStr(ICount(i)) = OStr(ICount(i)) + Format(Str(i), "00") + " "
    ICount(i) = 0

Next i

'*******************************************************************
For i = 0 To 9
    For j = 1 To 6
        ICount(Val(Mid(IStr(i), 2 * j - 1, 2))) = ICount(Val(Mid(IStr(i), 2 * j - 1, 2))) + 1
    Next j
Next i
For i = 1 To 33
 OStr(ICount(i) + 5) = OStr(ICount(i) + 5) + Format(Str(i), "00") + " "
Next i

    


End Function

Public Function CheckReg6(IStr As String) As String
Dim i As Integer
Dim Reg(6) As Integer
Dim TempStr As String
For i = 1 To 6
 Reg((Val(Mid(IStr, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(IStr, 2 * i - 1, 2)) - 1) \ 5) + 1
Next i
For i = 0 To 6
TempStr = TempStr + Format(Str(Reg(i)))
Next i
CheckReg6 = TempStr
End Function








