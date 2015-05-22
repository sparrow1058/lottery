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

Type ZhiBiao
BMS As Integer ' big middle small
R012 As Integer   ' region 0 1 2
CHFlag As Integer  ' over  BMS , R012
BSFlag As Integer ' Big or small
OEFlag As Integer ' Odd  or even
ZHFlag As Integer  ' 质数 或者合数

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

Public Function CombinCheck(tempstr As String, num As Integer, OutStr() As String)
    'tempstr as 10 char lenght
  If Len(tempstr) = 10 Then
 Select Case num
    Case 0:
        OutStr(0) = ""
    Case 1:
        OutStr(0) = Mid(tempstr, 1, 2)
        OutStr(1) = Mid(tempstr, 3, 2)
        OutStr(2) = Mid(tempstr, 5, 2)
        OutStr(3) = Mid(tempstr, 7, 2)
        OutStr(4) = Mid(tempstr, 9, 2)
    Case 2:
        OutStr(0) = Mid(tempstr, 1, 4)
        OutStr(1) = Mid(tempstr, 1, 2) + Mid(tempstr, 5, 2)
        OutStr(2) = Mid(tempstr, 1, 2) + Mid(tempstr, 7, 2)
        OutStr(3) = Mid(tempstr, 1, 2) + Mid(tempstr, 9, 2)
        
        OutStr(4) = Mid(tempstr, 3, 4)
        OutStr(5) = Mid(tempstr, 3, 2) + Mid(tempstr, 7, 2)
        OutStr(6) = Mid(tempstr, 3, 2) + Mid(tempstr, 9, 2)
        
        OutStr(7) = Mid(tempstr, 5, 4)
        OutStr(8) = Mid(tempstr, 5, 2) + Mid(tempstr, 9, 2)
        
        OutStr(9) = Mid(tempstr, 7, 4)
        
  
    Case 3:
        OutStr(0) = Mid(tempstr, 1, 6)
        OutStr(1) = Mid(tempstr, 1, 4) + Mid(tempstr, 7, 2)
        OutStr(2) = Mid(tempstr, 1, 4) + Mid(tempstr, 9, 2)
        OutStr(3) = Mid(tempstr, 1, 2) + Mid(tempstr, 5, 4)
        OutStr(4) = Mid(tempstr, 1, 2) + Mid(tempstr, 7, 4)
        OutStr(5) = Mid(tempstr, 1, 2) + Mid(tempstr, 5, 2) + Mid(tempstr, 9, 2)
        
        
        OutStr(6) = Mid(tempstr, 3, 6)
        OutStr(7) = Mid(tempstr, 3, 2) + Mid(tempstr, 7, 4)
       
        OutStr(8) = Mid(tempstr, 3, 4) + Mid(tempstr, 9, 2)
        OutStr(9) = Mid(tempstr, 5, 6)
         
   End Select
   End If
   If Len(tempstr) = 6 Then
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

Public Function CheckR6(InStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    For i = 0 To 6
        CheckR6 = CheckR6 + Format(Str(Reg(i)), "0")
    Next i
End Function
Public Function CheckR3(InStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    CheckR3 = CheckR3 + Format(Str(Reg(0) + Reg(1) + Reg(2)), "0")
    CheckR3 = CheckR3 + Format(Str(Reg(3) + Reg(4) + Reg(5)), "0")
    CheckR3 = CheckR3 + Format(Str(Reg(6)), "0")
End Function
Public Function CheckR4(InStr1 As String) As String
    Dim i As Integer
    Dim Reg(7) As Integer
    For i = 1 To 6
        Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) = Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 5) + 1
    Next i
    CheckR4 = CheckR4 + Format(Str(Reg(0) + Reg(1)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(2) + Reg(3)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(4) + Reg(5)), "0")
    CheckR4 = CheckR4 + Format(Str(Reg(6)), "0")
End Function
Public Function CheckR11(InStr1 As String) As String
    Dim i As Integer
    Dim Reg(3) As Integer
    For i = 1 To 6
        Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 11) = Reg((Val(Mid(InStr1, 2 * i - 1, 2)) - 1) \ 11) + 1
    Next i
    For i = 0 To 2
    CheckR11 = CheckR11 + Format(Str(Reg(i)), "0")
    Next i
End Function
Public Function CheckAC(InStr1 As String) As Integer
 Dim i As Integer
 Dim ACData(32) As Integer
 Dim AC(6) As Integer
 Dim ACNum As Integer
  For i = 1 To 6
   AC(i) = Val(Mid(InStr1, 2 * i - 1, 2))
  Next i


  ACData(AC(6) - AC(5)) = 1
  ACData(AC(6) - AC(4)) = 1
  ACData(AC(6) - AC(3)) = 1
  ACData(AC(6) - AC(2)) = 1
  ACData(AC(6) - AC(1)) = 1
  
  ACData(AC(5) - AC(4)) = 1
  ACData(AC(5) - AC(3)) = 1
  ACData(AC(5) - AC(2)) = 1
  ACData(AC(5) - AC(1)) = 1
      
  ACData(AC(4) - AC(3)) = 1
  ACData(AC(4) - AC(2)) = 1
  ACData(AC(4) - AC(1)) = 1
 
  ACData(AC(3) - AC(2)) = 1
  ACData(AC(3) - AC(1)) = 1
  
  ACData(AC(2) - AC(1)) = 1


 For i = 1 To 32
  If ACData(i) = 1 Then
   ACNum = ACNum + 1
   ACData(i) = 0
  End If
Next i
 CheckAC = ACNum - 5

End Function


Public Function AddSum(InputStr() As String, OutStr() As StrSum)
 Dim i, j, Out, Total As Integer
 Dim tempstr As String
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
   Dim tempstr As StrSum
   For i = LBound(InputData()) To UBound(InputData())
       For j = UBound(InputData()) To i + 1 Step -1
            If InputData(j).StrSum > InputData(j - 1).StrSum Then
                tempstr.StrChr = InputData(j - 1).StrChr
                tempstr.StrSum = InputData(j - 1).StrSum
                InputData(j - 1).StrChr = InputData(j).StrChr
                InputData(j - 1).StrSum = InputData(j).StrSum
                InputData(j).StrChr = tempstr.StrChr
                InputData(j).StrSum = tempstr.StrSum
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
Public Function AddAllSum(tempstr As String, OutSum() As Integer)
Dim i As Integer
For i = 1 To 6
    OutSum(Val(Mid(tempstr, 2 * i - 1, 2))) = OutSum(Val(Mid(tempstr, 2 * i - 1, 2))) + 1
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
Dim tempstr As String
tempstr = FindSum + "- "
For i = UBound(SumList()) To LBound(SumList()) Step -1
   If Not (SumList(i) = "") Then
    j = j + 1
    If SumList(i) = FindSum Then
        tempstr = tempstr + Format(Str(j - 1), "00") + " "
        j = 0
   End If
   End If
Next i
GetLost = tempstr

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
Dim tempstr As String
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
tempstr = ""
For i = 1 To Count
    tempstr = tempstr + "1"
Next i
For i = Count + 1 To Pleft
    tempstr = tempstr + "0"
Next i
tempstr = tempstr + "01"
tempstr = tempstr + Right(IStr, Pright)

ChangeStr = tempstr
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
Public Function Combin1Num(IStr As String, flagstr As String) As String          ' 从字符串中取出 字符
Dim TotalNum As Integer
Dim LenTemp As Integer
Dim tempstr As String
Dim i As Integer
If Len(flagstr) = Len(IStr) Then
 For i = 1 To Len(flagstr)
    If (InStr(i, flagstr, "1") = i) Then
        tempstr = tempstr + Mid(IStr, i, 1)
    End If
Next i
End If
Combin1Num = tempstr
End Function
Public Function CombinNum(IStr As String, flagstr As String) As String          ' 从字符串中取出 字符
Dim TotalNum As Integer
Dim LenTemp As Integer
Dim tempstr As String
Dim i As Integer
If Len(flagstr) = Len(IStr) \ 2 Then
 For i = 1 To Len(flagstr)
    If (InStr(i, flagstr, "1") = i) Then
        tempstr = tempstr + Mid(IStr, 2 * i - 1, 2)
    End If
Next i
End If
CombinNum = tempstr
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
Dim tempstr As String
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

Public Function CheckLHB(TempstrB As String, TempstrN As String) As Integer         '检验上期连号
Dim i, j As Integer
Dim LHCount As Integer
For i = 1 To 6
 For j = 1 To 6
  If Abs(Val(Mid(TempstrN, 2 * i - 1, 2)) - Val(Mid(TempstrB, 2 * j - 1, 2))) = 1 Then
     LHCount = LHCount + 1
     Exit For
  End If
  Next j
Next i
 CheckLHB = LHCount
End Function
Public Function CheckTHB(TempstrB As String, TempstrN As String) As Integer    ' 检验本期连号
Dim i, j As Integer
Dim THCount As Integer
For i = 1 To 6
 For j = 1 To 6
  If Abs(Val(Mid(TempstrN, 2 * i - 1, 2)) - Val(Mid(TempstrB, 2 * j - 1, 2))) = 0 Then
     THCount = THCount + 1
     Exit For
  End If
  Next j
Next i
 CheckTHB = THCount
End Function

Public Function AddStrSum(IStr As String) As Integer        '检验分布次数和值， 如112 =4  110 =2   为6区分布 方便排序
   Dim i, Sum As Integer
   For i = 1 To Len(IStr)
   Sum = Sum + Val(Mid(IStr, i, 1))
   Next i
    AddStrSum = Sum
End Function

Public Function SubStr(IStr1 As String, IStr2 As String, SubSum As Integer) As String
Dim tempstr As String
Dim i As Integer
Dim SubValue As Integer
For i = 1 To 6
SubValue = Val(Mid(IStr1, 2 * i - 1, 2)) - Val(Mid(IStr2, 2 * i - 1, 2))
SubSum = SubSum + SubValue
tempstr = tempstr + Format(Str(SubValue), "@@@") + " "
Next i

SubStr = tempstr

End Function

Public Function NextNum(IStr As String) As String


End Function
Public Function CheckZB(INum As Integer, zbr As ZhiBiao)
Dim InNum As Integer
InNum = INum
InNum = InNum Mod 10
' 0, 1  , 2 means small ,   3 4 5 6  means  middle ,, 7 8 9 means Big Number

If InNum < 3 Then
 zbr.BMS = 0
ElseIf InNum > 6 Then
 zbr.BMS = 2
 Else
 zbr.BMS = 1
End If

' 0 1 2  路数据分布 ，，
If (InNum = 0 Or InNum = 3 Or InNum = 6 Or InNum = 9) Then
 zbr.R012 = 0
ElseIf (InNum = 1 Or INum = 5 Or InNum = 7) Then
 zbr.R012 = 1
 Else
 zbr.R012 = 2
 End If

'重合码  1 3   6  8  ，， 区间分布 与 路数分布 重合
If (InNum = 1 Or InNum = 3 Or InNum = 6 Or InNum = 8) Then
 zbr.CHFlag = 1
 Else
 zbr.CHFlag = 0
End If

' 大小数 分布
If (InNum < 5) Then
 zbr.BSFlag = 0
 Else
 zbr.BSFlag = 1
End If


' 奇偶数分布
If (InNum Mod 2) = 0 Then
 zbr.OEFlag = 0
Else
 zbr.OEFlag = 1
End If


' 质数 合数 分布
If (InNum = 1 Or InNum = 2 Or InNum = 3 Or InNum = 5 Or InNum = 7) Then
  zbr.ZHFlag = 0
Else
  zbr.ZHFlag = 1
End If


End Function
Public Function ShowZBR(zbr As ZhiBiao) As String
 Dim tempstr(6) As String
 Dim i As Integer
If zbr.BMS = 0 Then
  tempstr(0) = "  ●  |      |      |"     ''★◆ ●■〓 ★▲■●▲
ElseIf zbr.BMS = 1 Then
  tempstr(0) = "      |  ●  |      |"
 Else
  tempstr(0) = "      |      |  ●  |"
 End If
 
If zbr.R012 = 0 Then
  tempstr(1) = "|  ●  |      |      |"       ''★◆
  ElseIf zbr.R012 = 1 Then
  tempstr(1) = "|      |  ●  |      |"
 Else
  tempstr(1) = "|      |      |  ●  |"
End If
  
' 数据重合标志
 If zbr.CHFlag = 0 Then
   tempstr(2) = "| ▲  |      |"
  Else
   tempstr(2) = "|      |  ▲ |"
 End If
 '大数标志
If zbr.BSFlag = 0 Then
   tempstr(3) = "|  ■  |      |"
  Else
   tempstr(3) = "|      |  ■  |"
 End If
'奇偶数标志
If zbr.OEFlag = 0 Then
   tempstr(4) = "|  ▲ |      |"
  Else
   tempstr(4) = "|      |  ▲ |"
 End If

'质数，合数标志
If zbr.ZHFlag = 0 Then
   tempstr(5) = "|  ■  |      |"
  Else
   tempstr(5) = "|      |  ■  |"
 End If
For i = 0 To 5
 ShowZBR = ShowZBR + tempstr(i) + "|"
Next i
End Function

Public Function RegionFenbu(IStr As String) As String



End Function

