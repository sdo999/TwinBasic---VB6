Attribute VB_Name = "modCJK"
' =====================================================================
' 'modCJK.bas--回退文本呈现器
' '《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
' “HoMM2/HoMM3精灵表仅包含拉丁语、西里尔语和波兰语
' ”和匈牙利象形文字。其他（中文、日文、韩文），
' “表格中缺少的任何字符）都用实数绘制在这里
' “Windows字体不会被静默删除。
' '
' “外观与每场比赛相匹配：
' 'HoMM3:奶油填充（#FFF3DE），1像素黑色轮廓
' 'HoMM2：近白色填充（#FCFCFC），带有柔和的深色阴影
' '
' “测量值是按代码点缓存的，所以有一条很长的中文消息
'每个不同字符花费一个GdipMeasureString。
' =====================================================================
Option Explicit

Public gCJK_FontName   As String
Public gCJK_FontSize   As Long
Public gCJK_ExtraY     As Long      '手动垂直轻推（像素）
Public gCJK_UseOutline As Boolean    ' False = 仅绘制纯字形

Private m_family    As LongPtr
Private m_font      As LongPtr
Private m_sf        As LongPtr
Private m_gfx       As LongPtr      ' 用于测量的划痕图形
Private m_bmp       As LongPtr
Private m_brFill    As LongPtr
Private m_brLine    As LongPtr

Private m_fillColor As Long
Private m_lineColor As Long
Private m_lineMode  As Long      ' 0=无，1=8向轮廓，2=阴影

Private m_ready     As Boolean
Private m_tried     As Boolean

Private m_refOffY   As Single       ' 参考字形的度量（U+4E2D）
Private m_refH      As Single

Private Type CJK_METRIC
    Code    As Long
    Advance As Long
    OffX    As Single
    OffY    As Single
    InkW    As Single
    InkH    As Single
End Type

Private m_met()     As CJK_METRIC
Private m_metCount  As Long

' ------------------------------------------------------------------
Public Sub CJK_Reset()
    m_ready = False
    m_tried = False
    m_metCount = 0
    If m_font <> 0 Then GdipDeleteFont m_font
    If m_family <> 0 Then GdipDeleteFontFamily m_family
    If m_sf <> 0 Then GdipDeleteStringFormat m_sf
    If m_brFill <> 0 Then GdipDeleteBrush m_brFill
    If m_brLine <> 0 Then GdipDeleteBrush m_brLine
    ' 测量用的临时位图与图形（以前一直没释放，退出时统一回收）
    If m_bmp <> 0 Then GdipDisposeImage m_bmp
    If m_gfx <> 0 Then GdipDeleteGraphics m_gfx
    m_font = 0: m_family = 0: m_sf = 0: m_brFill = 0: m_brLine = 0
    m_bmp = 0: m_gfx = 0
End Sub

' “每次渲染调用一次；在HoMM2和
' “HoMM3的外观。便宜的时候什么都没变。
Public Sub CJK_SetColors(ByVal FillColor As Long, ByVal LineColor As Long, _
                         ByVal LineMode As Long)
    If Not CJK_Ensure() Then Exit Sub
    If FillColor <> m_fillColor Then
        If m_brFill <> 0 Then GdipDeleteBrush m_brFill
        m_brFill = 0
        GdipCreateSolidFill FillColor, m_brFill
        m_fillColor = FillColor
    End If
    If LineColor <> m_lineColor Then
        If m_brLine <> 0 Then GdipDeleteBrush m_brLine
        m_brLine = 0
        GdipCreateSolidFill LineColor, m_brLine
        m_lineColor = LineColor
    End If
    m_lineMode = LineMode
End Sub

Public Function CJK_Ready() As Boolean
    CJK_Ready = CJK_Ensure()
End Function

' ------------------------------------------------------------------
Private Function CJK_Ensure() As Boolean
    If m_ready Then
        CJK_Ensure = True
        Exit Function
    End If
    If m_tried Then Exit Function          ' GDI+不可用-悄悄放弃
    m_tried = True

    If Not GdiPlus_Ensure() Then Exit Function

    Dim names() As String, k As Long
    names = Split(gCJK_FontName & "|SimSun|NSimSun|SimHei|Microsoft YaHei|MingLiU|Tahoma|Arial", "|")

    For k = LBound(names) To UBound(names)
        If Len(names(k)) > 0 Then
            If GdipCreateFontFamilyFromName(StrPtr(names(k)), 0, m_family) = GpOk Then
                If m_family <> 0 Then Exit For
            End If
        End If
    Next k
    If m_family = 0 Then Exit Function

    Dim sz As Long
    sz = gCJK_FontSize
    If sz < 6 Then sz = 6
    If sz > 96 Then sz = 96

    If GdipCreateFont(m_family, CSng(sz), FontStyleRegular, UnitPixel, m_font) <> GpOk Then Exit Function
    GdipCreateStringFormat(StringFormatFlagsNoClip, 0, m_sf)

    ' 划痕位图+图形，仅用于测量
    m_bmp = GdiPlus_NewBitmap(8, 8, m_gfx)
    If m_bmp = 0 Then Exit Function

    GdipCreateSolidFill &HFFFFFFFF, m_brFill
    GdipCreateSolidFill &HFF000000, m_brLine
    m_fillColor = &HFFFFFFFF
    m_lineColor = &HFF000000

    ' 参考度量：U+4E2D=“中”字
    Dim m As CJK_METRIC
    If MeasureOne(ChrW$(&H4E2D), m) Then
        m_refOffY = m.OffY
        m_refH = m.InkH
    Else
        m_refOffY = 0
        m_refH = CSng(sz)
    End If

    m_ready = True
    CJK_Ensure = True
End Function

' ------------------------------------------------------------------
Private Function CJK_CharStr(ByVal Code As Long) As String
    Dim n As Long
    n = Code
    If n > 32767 Then n = n - 65536
    CJK_CharStr = ChrW$(n)
End Function

Private Function MeasureOne(ByRef s As String, ByRef m As CJK_METRIC) As Boolean
    Dim rc As RECTF, bb As RECTF
    Dim fitted As Long, lines As Long
    rc.Left = 0
    rc.Top = 0
    rc.Width = 100000
    rc.Height = 100000
    If GdipMeasureString(m_gfx, StrPtr(s), Len(s), m_font, rc, m_sf, bb, fitted, lines) <> GpOk Then
        Exit Function
    End If
    m.OffX = bb.Left
    m.OffY = bb.Top
    m.InkW = bb.Width
    m.InkH = bb.Height
    ' 天花板，因此亚像素的进步永远不会四舍五入到零
    m.Advance = CLng(-Int(-(bb.Left + bb.Width)))
    If m.Advance < 1 Then m.Advance = 1
    MeasureOne = True
End Function

Private Function CJK_Get(ByVal Code As Long, ByRef m As CJK_METRIC) As Boolean
    Dim i As Long
    For i = 0 To m_metCount - 1
        If m_met(i).Code = Code Then
            m = m_met(i)
            CJK_Get = True
            Exit Function
        End If
    Next i

    m.Code = Code
    Dim s As String
    s = CJK_CharStr(Code)
    If Not MeasureOne(s, m) Then
        m.Advance = gCJK_FontSize
        m.OffX = 0
        m.OffY = 0
        m.InkW = CSng(gCJK_FontSize)
        m.InkH = CSng(gCJK_FontSize)
    End If

    If m_metCount = 0 Then ReDim m_met(0 To 255)
    If m_metCount > UBound(m_met) Then ReDim Preserve m_met(0 To m_metCount + 255)
    m_met(m_metCount) = m
    m_metCount = m_metCount + 1
    CJK_Get = True
End Function

' ------------------------------------------------------------------
Public Function CJK_Width(ByVal Code As Long) As Long
    If Not CJK_Ensure() Then
        CJK_Width = gCJK_FontSize
        If CJK_Width < 4 Then CJK_Width = 4
        Exit Function
    End If
    Dim m As CJK_METRIC
    If Not CJK_Get(Code, m) Then
        CJK_Width = gCJK_FontSize
        Exit Function
    End If
    CJK_Width = m.Advance
End Function

'绘制一个字符，其左边缘位于x，基线位于
'“基线Y。
Public Sub CJK_Draw(ByVal gfx As LongPtr, ByVal Code As Long, _
                    ByVal x As Long, ByVal baselineY As Long)
    If gfx = 0 Then Exit Sub
    If Not CJK_Ensure() Then Exit Sub
    If m_brFill = 0 Then Exit Sub

    Dim m As CJK_METRIC
    If Not CJK_Get(Code, m) Then Exit Sub

    Dim s As String
    s = CJK_CharStr(Code)

    '通过将单元格锚定到以下位置，将每个字符放在同一基线上
    ' “参考字形，而不是它自己的墨盒。
    Dim cellTop As Single
    cellTop = CSng(baselineY + gCJK_ExtraY) - (m_refOffY + m_refH)

    Dim rc As RECTF
    rc.Width = 100000
    rc.Height = 100000

    If m_brLine <> 0 And m_lineMode <> 0 And gCJK_UseOutline Then
        If m_lineMode = 1 Then
            Dim ox As Long, oy As Long
            For oy = -1 To 1
                For ox = -1 To 1
                    If ox <> 0 Or oy <> 0 Then
                        rc.Left = CSng(x) - m.OffX + CSng(ox)
                        rc.Top = cellTop - m.OffY + CSng(oy)
                        GdipDrawString gfx, StrPtr(s), 1, m_font, rc, m_sf, m_brLine
                    End If
                Next ox
            Next oy
        Else
            ' HoMM2：左下一个像素的柔和阴影
            rc.Left = CSng(x) - m.OffX - 1
            rc.Top = cellTop - m.OffY + 1
            GdipDrawString gfx, StrPtr(s), 1, m_font, rc, m_sf, m_brLine
        End If
    End If

    rc.Left = CSng(x) - m.OffX
    rc.Top = cellTop - m.OffY
    GdipDrawString gfx, StrPtr(s), 1, m_font, rc, m_sf, m_brFill
End Sub
