Attribute VB_Name = "modGdiPlus"
' =====================================================================
'modGdiPlus.bas--GDI+平面API的薄包装
'《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
'twinBASIC（VB7）源代码。使用LongPtr，以便编译相同的代码
'适用于Win32和Win64目标。
' =====================================================================
Option Explicit

Public Type CLSID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Public Const PixelFormat32bppARGB   As Long = &H26200A
Public Const UnitPixel              As Long = 2
Public Const InterpolationModeNearestNeighbor As Long = 5
Public Const PixelOffsetModeHalf    As Long = 4
Public Const GpOk                   As Long = 0

Public Const FontStyleRegular       As Long = 0
Public Const FontStyleBold          As Long = 1
Public Const StringFormatFlagsNoClip As Long = &H4000

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As LongPtr, ByVal InputBuf As LongPtr, ByVal OutputBuf As LongPtr) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" ( _
    ByVal Token As LongPtr) As Long

Public Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" ( _
    ByVal FileName As LongPtr, ByRef Bitmap As LongPtr) As Long
Public Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" ( _
    ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, _
    ByVal PixelFormat As Long, ByVal Scan0 As LongPtr, ByRef Bitmap As LongPtr) As Long
Public Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" ( _
    ByVal Image As LongPtr, ByRef Graphics As LongPtr) As Long
Public Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" ( _
    ByVal Graphics As LongPtr) As Long
Public Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" ( _
    ByVal Image As LongPtr) As Long
Public Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, ByVal Color As Long) As Long
Public Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, ByVal Mode As Long) As Long
Public Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, ByVal Mode As Long) As Long
Public Declare PtrSafe Function GdipDrawImageRectRectI Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, ByVal Image As LongPtr, _
    ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, _
    ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, _
    ByVal srcUnit As Long, ByVal imageAttributes As LongPtr, _
    ByVal Callback As LongPtr, ByVal callbackData As LongPtr) As Long
Public Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" ( _
    ByVal Color As Long, ByRef Brush As LongPtr) As Long
Public Declare PtrSafe Function GdipFillRectangleI Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, ByVal Brush As LongPtr, _
    ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" ( _
    ByVal Brush As LongPtr) As Long
Public Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" ( _
    ByVal Image As LongPtr, ByVal FileName As LongPtr, _
    ByRef clsidEncoder As CLSID, ByVal encoderParams As LongPtr) As Long
Public Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal Bitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal Background As Long) As Long

' ----text/font（由CJK回退渲染器使用）----
Public Declare PtrSafe Function GdipCreateFontFamilyFromName Lib "gdiplus" ( _
    ByVal Name As LongPtr, ByVal fontCollection As LongPtr, _
    ByRef fontFamily As LongPtr) As Long
Public Declare PtrSafe Function GdipDeleteFontFamily Lib "gdiplus" ( _
    ByVal fontFamily As LongPtr) As Long
Public Declare PtrSafe Function GdipCreateFont Lib "gdiplus" ( _
    ByVal fontFamily As LongPtr, ByVal emSize As Single, ByVal style As Long, _
    ByVal unit As Long, ByRef font As LongPtr) As Long
Public Declare PtrSafe Function GdipDeleteFont Lib "gdiplus" ( _
    ByVal font As LongPtr) As Long
Public Declare PtrSafe Function GdipCreateStringFormat Lib "gdiplus" ( _
    ByVal formatAttributes As Long, ByVal language As Long, _
    ByRef strFormat As LongPtr) As Long
Public Declare PtrSafe Function GdipDeleteStringFormat Lib "gdiplus" ( _
    ByVal strFormat As LongPtr) As Long
Public Declare PtrSafe Function GdipMeasureString Lib "gdiplus" ( _
    ByVal graphics As LongPtr, ByVal strLong As LongPtr, ByVal length As Long, _
    ByVal font As LongPtr, ByRef layoutRect As RECTF, ByVal strFormat As LongPtr, _
    ByRef boundingBox As RECTF, ByRef codepointsFitted As Long, _
    ByRef linesFilled As Long) As Long
Public Declare PtrSafe Function GdipDrawString Lib "gdiplus" ( _
    ByVal graphics As LongPtr, ByVal strLong As LongPtr, ByVal length As Long, _
    ByVal font As LongPtr, ByRef layoutRect As RECTF, ByVal strFormat As LongPtr, _
    ByVal brush As LongPtr) As Long

Private m_Token As LongPtr
Private m_Started As Boolean

'------------------------------------------------------------------
Public Function GdiPlus_Ensure() As Boolean
    If m_Started Then
        GdiPlus_Ensure = True
        Exit Function
    End If
    '将启动块构建为原始字节缓冲区，使布局
    '在x86（16字节）和x64（32字节）上都是正确的。
    Dim buf(0 To 63) As Byte
    buf(0) = 1                                  ' GdiplusVersion = 1
    If GdiplusStartup(m_Token, VarPtr(buf(0)), 0) = GpOk Then m_Started = True
    GdiPlus_Ensure = m_Started
End Function

Public Sub GdiPlus_Shutdown()
    If m_Started Then
        GdiplusShutdown m_Token
        m_Started = False
    End If
End Sub

'------------------------------------------------------------------
Public Function GdiPlus_LoadPng(ByVal FileName As String) As LongPtr
    Dim bmp As LongPtr
    If GdipCreateBitmapFromFile(StrPtr(FileName), bmp) <> GpOk Then
        GdiPlus_LoadPng = 0
        Exit Function
    End If
    GdiPlus_LoadPng = bmp
End Function

Public Function GdiPlus_NewBitmap(ByVal Width As Long, ByVal Height As Long, _
                                  ByRef Graphics As LongPtr) As LongPtr
    Dim bmp As LongPtr
    If GdipCreateBitmapFromScan0(Width, Height, 0, PixelFormat32bppARGB, 0, bmp) <> GpOk Then
        GdiPlus_NewBitmap = 0
        Exit Function
    End If
    GdipGetImageGraphicsContext bmp, Graphics
    GdiPlus_NewBitmap = bmp
End Function

Public Function GdiPlus_SavePng(ByVal Bitmap As LongPtr, ByVal FileName As String) As Boolean
    Dim cls As CLSID
    cls.Data1 = &H557CF406
    cls.Data2 = &H1A04
    cls.Data3 = &H11D3
    cls.Data4(0) = &H9A
    cls.Data4(1) = &H73
    cls.Data4(2) = &H0
    cls.Data4(3) = &H0
    cls.Data4(4) = &HF8
    cls.Data4(5) = &H1E
    cls.Data4(6) = &HF3
    cls.Data4(7) = &H2E
    GdiPlus_SavePng = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), cls, 0) = GpOk)
End Function
