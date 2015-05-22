Attribute VB_Name = "Module3"
'Ҫ��API?
'�ڴ����з�һ��Text1 , һ��List1, һ��Command1?
'---------����ģ����-------------
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
'����BMPλͼ
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long
  Dim Pic As PicBmp
  Dim IPic As IPicture
  Dim IID_IDispatch As GUID
  '���IDispatch����
  With IID_IDispatch
     .Data1 = &H20400
     .Data4(0) = &HC0
     .Data4(7) = &H46
  End With
  '���Pic
  With Pic
     .Size = Len(Pic)          'ע�ͣ� Pic�ṹ����
     .Type = vbPicTypeBitmap   'ע�ͣ� ͼ������
     .hBmp = hBmp              'ע�ͣ� λͼ���
     .hPal = hPal              'ע�ͣ� ��ɫ����
  End With
  '����Pictureͼ��
  r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
  '����Picture����
  Set CreateBitmapPicture = IPic
End Function
'��ͼ����
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal _
    LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc _
    As Long) As Picture
    Dim hDCMemory As Long '�����ȡͼ���Ŀ���豸
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long 'Ҫ��ȡͼ���Դ�豸
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
'GetDC��������д�봰����ʾ������豸���ݾ��,��GetWindowDC����д���������ڵ��豸���ݾ��
'��������GetDC�������߿򡢹����������������˵��ȣ���GetWindowDC�����
 If Client Then '���Ϊ�棬��ָ���ǿͻ������������������ȣ�
        hDCSrc = GetDC(hWndSrc)  'GetDC����һָ�����ڵĿͻ������������Ļ����ʾ�豸�����ĵľ��
    Else   '������GetWindowDCѰ�Һ��ȡ
        hDCSrc = GetWindowDC(hWndSrc)
    End If
    hDCMemory = CreateCompatibleDC(hDCSrc) '����һ����hDCSrc�豸����һ�����ڴ���
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc) '����һ�����豸�й�λͼ
    hBmpPrev = SelectObject(hDCMemory, hBmp) 'SelectObject��λͼ�����豸������
    '�����Ļ����
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) '����ָ���豸����������豸�Ĺ��ܷ�����Ϣ
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    '�����Ļ�����е�ɫ��������Ļ��ɫ��
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        '������Ļ��ɫ��Ŀ���
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0)) '��ȡϵͳ��ɫ��
        hPal = CreatePalette(LogPal) 'CreatePalette��ɫ�庯��
        '���½����ĵ�ɫ��ѡ�뽨�����ڴ��ͼ�����
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory) 'RealizePalette����ʹϵͳ�ָ���ǰѡ�е��߼���ɫ���е�ֵ
    End If
    '����ͼ��
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    '�ͷ���Դ
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
'capturescreen������׽������Ļͼ��
Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long
'�������Ĵ��ھ��
    hWndScreen = GetDesktopWindow()
    'Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width _
        \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
        
        Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width _
        \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
'��׽��ǰ�����ͼ��
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
'�������Ĵ��ھ��
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

'��ͼģ�����Module1.bas
'------------���ڴ�����----------------
'Private Sub Command1_Click()
'Dim tmpPicture As Picture
'Public Const snapFolder = "c:\Snap"  '��ͼ����λ��
'If List1.ListIndex = 0 Then
    '��׽�����
 '   Set tmpPicture = CaptureActiveWindow()
    
    '��׽������Ļ
'ElseIf List1.ListIndex = 1 Then
'Set tmpPicture = CaptureScreen()
'Else: MsgBox "                ��ѡ��һ�ַ�ʽ��                  "
'Exit Sub
'����ʱ
'End If
 'SavePicture tmpPicture, Text1.Text
'End Sub



