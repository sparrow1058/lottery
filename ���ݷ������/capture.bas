Attribute VB_Name = "Module3"
'要用API?
'在窗体中放一个Text1 , 一个List1, 一个Command1?
'---------放在模块中-------------
Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type
Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal _
        iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, _
        ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries _
        As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) _
        As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject _
        As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As _
        Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop _
        As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette _
        As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As _
        RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As _
        Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As _
        PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'创建BMP位图
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long
  Dim Pic As PicBmp
  Dim IPic As IPicture
  Dim IID_IDispatch As GUID
  '填充IDispatch界面
  With IID_IDispatch
     .Data1 = &H20400
     .Data4(0) = &HC0
     .Data4(7) = &H46
  End With
  '填充Pic
  With Pic
     .Size = Len(Pic)          '注释： Pic结构长度
     .Type = vbPicTypeBitmap   '注释： 图象类型
     .hBmp = hBmp              '注释： 位图句柄
     .hPal = hPal              '注释： 调色板句柄
  End With
  '建立Picture图象
  r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
  '返回Picture对象
  Set CreateBitmapPicture = IPic
End Function
'截图处理
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal _
    LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc _
    As Long) As Picture
    Dim hDCMemory As Long '保存截取图象的目标设备
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long '要截取图象的源设备
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
'GetDC传回用于写入窗口显示区域的设备内容句柄,而GetWindowDC传回写入整个窗口的设备内容句柄
'区别在于GetDC不包括边框、滚动条、标题栏、菜单等，而GetWindowDC则包括
 If Client Then '如果为真，即指定是客户区（不包括标题栏等）
        hDCSrc = GetDC(hWndSrc)  'GetDC检索一指定窗口的客户区域或整个屏幕的显示设备上下文的句柄
    Else   '否则用GetWindowDC寻找后获取
        hDCSrc = GetWindowDC(hWndSrc)
    End If
    hDCMemory = CreateCompatibleDC(hDCSrc) '创建一块与hDCSrc设备场景一样的内存区
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc) '创建一幅与设备有关位图
    hBmpPrev = SelectObject(hDCMemory, hBmp) 'SelectObject将位图放入设备场景中
    '获得屏幕属性
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) '根据指定设备场景代表的设备的功能返回信息
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    '如果屏幕对象有调色板则获得屏幕调色板
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        '建立屏幕调色板的拷贝
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0)) '获取系统调色板
        hPal = CreatePalette(LogPal) 'CreatePalette调色板函数
        '将新建立的调色板选入建立的内存绘图句柄中
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory) 'RealizePalette函数使系统恢复当前选中的逻辑调色板中的值
    End If
    '拷贝图象
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    '释放资源
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
'capturescreen函数捕捉整个屏幕图象
Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long
'获得桌面的窗口句柄
    hWndScreen = GetDesktopWindow()
    'Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width _
        \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
        
        Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width _
        \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
'捕捉当前活动窗口图象
Public Function CaptureActiveWindow() As Picture
    Dim hWndActive As Long
    Dim r As Long
    Dim RectActive As RECT
  
    hWndActive = GetForegroundWindow()
    r = GetWindowRect(hWndActive, RectActive)
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
        RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function
Public Function CaptureSel(pLeft As Long, pTop As Long, pWidth As Long, pHeight As Long) As Picture
    Dim hWndScreen As Long
'获得桌面的窗口句柄
    hWndScreen = GetDesktopWindow()
    
        Set CaptureSel = CaptureWindow(hWndScreen, False, pLeft, pTop, pWidth, pHeight)
End Function
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
  
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait
    Else
        Prn.Orientation = vbPRORLandscape
    End If
  
    PicRatio = Pic.Width / Pic.Height
  
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    PrnRatio = PrnWidth / PrnHeight
  
    If PicRatio >= PrnRatio Then
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
  
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub

'截图模块结束Module1.bas
'------------放在窗体中----------------
'Private Sub Command1_Click()
'Dim tmpPicture As Picture
'Public Const snapFolder = "c:\Snap"  '截图保存位置
'If List1.ListIndex = 0 Then
    '捕捉活动窗口
 '   Set tmpPicture = CaptureActiveWindow()
    
    '捕捉整个屏幕
'ElseIf List1.ListIndex = 1 Then
'Set tmpPicture = CaptureScreen()
'Else: MsgBox "                请选择一种方式！                  "
'Exit Sub
'保存时
'End If
 'SavePicture tmpPicture, Text1.Text
'End Sub



