Attribute VB_Name = "modHomm3"
' =====================================================================
'modHomm3.bas——魔法门英雄III风格信息窗口
'
'js/Homm3MessageGenerator.js的忠实移植（作者：Alexei Sergienko）
'加上网络版本不能做的一件事：
'>>>您喜欢的任何输出宽度/高度，任何整数缩放<<<
'
'需要：modGdiPlus.bas、modSplit.bas、modCJK.bas、modGlyphs3.bas，
'img/homm3-prite.png
' =====================================================================
Option Explicit

' ---- fixed geometry taken from the original game --------------------
Private Const BORDER_SIZE   As Long = 64
Private Const BG_SIZE       As Long = 256
Private Const LINE_HEIGHT   As Long = 18
Private Const LETTER_SPACING As Long = 1

Private Const H_BORDER_H    As Long = 15     ' 水平边界条高度
Private Const V_BORDER_W    As Long = 14     ' 垂直边界条宽度
Private Const H_COLOR_H     As Long = 5      ' coloured accent, horizontal
Private Const V_COLOR_W     As Long = 4      ' coloured accent, vertical
Private Const COLOR_CORNER  As Long = 56     ' coloured corner ornament

Private Const BTN_W         As Long = 66
Private Const BTN_H         As Long = 32
Private Const BTN_MARGIN    As Long = 16

Private Const SHADOW_X      As Long = 8
Private Const SHADOW_Y      As Long = 8

Private Const SCROLL_SIDE   As Long = 16
Private Const SCROLL_M_T    As Long = 20
Private Const SCROLL_M_R    As Long = 11
Private Const SCROLL_M_B    As Long = 23

Private Const PAD_TOP       As Long = 11
Private Const PAD_BOTTOM    As Long = 15
Private Const PAD_RIGHT     As Long = 11
Private Const PAD_LEFT      As Long = 10

Private Const PAD_T         As Long = 0
Private Const PAD_B         As Long = 1
Private Const PAD_L         As Long = 2
Private Const PAD_R         As Long = 3

' ---- 子画面坐标 ---------------------------------------
Private Const SP_BG_X       As Long = 648
Private Const SP_BG_Y       As Long = 0
Private Const SP_OK_X       As Long = 516
Private Const SP_OK_Y       As Long = 228
Private Const SP_CANCEL_Y   As Long = 259
Private Const SP_SC_UP_X    As Long = 714
Private Const SP_SC_MID_X   As Long = 731
Private Const SP_SC_DN_X    As Long = 748
Private Const SP_SC_Y       As Long = 257

Private Const MIN_MANUAL_W  As Long = 160
Private Const MIN_MANUAL_H  As Long = 128
Private Const MAX_MANUAL_W  As Long = 4096
Private Const MAX_MANUAL_H  As Long = 4096

Private Const MAX_LINES     As Long = 64
Private Const MAX_WORDS     As Long = 256
Private Const RULE_SLOTS    As Long = 16

Public Const HOMM3_COLOR_COUNT As Long = 8

' =====================================================================
Public Type HOMM3_OPTIONS
    Text        As String
    ColorIndex  As Long          '0..7（见Homm3_ColorName）
    ShowOK      As Boolean
    ShowCancel  As Boolean
    DrawShadow  As Boolean
    ManualSize  As Boolean       ' True => use ManualWidth / ManualHeight
    ManualWidth As Long          ' 逻辑像素
    ManualHeight As Long
    ForceScroll As Boolean
    Zoom        As Long          ' 1..12, nearest-neighbour upscale
End Type

Public Type HOMM3_RESULT
    PopupWidth  As Long
    PopupHeight As Long
    OutWidth    As Long          ' 缩放后包括阴影
    OutHeight   As Long
    LineCount   As Long
    ScrollVisible As Boolean
    Clipped     As Boolean       ' 文本不合适，被剪切了
End Type

Private Type SIZE_RULE
    W As Long
    H As Long
    MaxTextLines As Long
    Scroll As Boolean
    PadT As Long
    PadR As Long
    PadB As Long
    PadL As Long
    HalfAt(0 To RULE_SLOTS - 1) As Long      ' -1 = 未使用的
    ResizeAt(0 To RULE_SLOTS - 1) As Long
End Type

' ---- module state ---------------------------------------------------
Private m_opt        As HOMM3_OPTIONS
Private m_size       As SIZE_RULE
Private m_rules()    As SIZE_RULE
Private m_ruleCount  As Long

Private m_colName()  As String
Private m_colCX()    As Long
Private m_colCY()    As Long
Private m_colHY()    As Long
Private m_colVX()    As Long
Private m_colCount   As Long

Private m_sprite     As LongPtr
Private m_spriteFile As String
Private m_gfx        As LongPtr

Private m_popupW     As Long
Private m_popupH     As Long
Private m_scrollVis  As Boolean
Private m_raiseHalf  As Boolean
Private m_overflow   As Boolean     ' 文本不符合 -> result.Clipped
Private m_tryNext    As Boolean     ' 此预设值太小 -> 试试下一个

Private m_lineCount  As Long
Private m_wordCount(0 To MAX_LINES - 1) As Long
Private m_words()    As String


Private m_ready      As Boolean

' =====================================================================
'                            T A B L E S
' =====================================================================
Private Sub AddColor(ByVal Name As String, ByVal CX As Long, ByVal CY As Long, _
                     ByVal HY As Long, ByVal VX As Long)
    If m_colCount = 0 Then ReDim m_colName(0 To 15): ReDim m_colCX(0 To 15)
    If m_colCount = 0 Then ReDim m_colCY(0 To 15): ReDim m_colHY(0 To 15): ReDim m_colVX(0 To 15)
    m_colName(m_colCount) = Name
    m_colCX(m_colCount) = CX
    m_colCY(m_colCount) = CY
    m_colHY(m_colCount) = HY
    m_colVX(m_colCount) = VX
    m_colCount = m_colCount + 1
End Sub

Private Sub AddSize(ByVal W As Long, ByVal H As Long, ByVal MaxLines As Long, _
                    ByVal Scroll As Long, ByVal PT As Long, ByVal PR As Long, _
                    ByVal PB As Long, ByVal PL As Long)
    If m_ruleCount = 0 Then ReDim m_rules(0 To 31)
    Dim i As Long
    With m_rules(m_ruleCount)
        .W = W
        .H = H
        .MaxTextLines = MaxLines
        .Scroll = (Scroll <> 0)
        .PadT = PT: .PadR = PR: .PadB = PB: .PadL = PL
        For i = 0 To RULE_SLOTS - 1
            .HalfAt(i) = -1
            .ResizeAt(i) = -1
        Next i
    End With
    m_ruleCount = m_ruleCount + 1
End Sub

Private Sub SetHalf(ByVal LineIndex As Long, ByVal Value As Long)
    If LineIndex >= 0 And LineIndex < RULE_SLOTS Then
        m_rules(m_ruleCount - 1).HalfAt(LineIndex) = Value
    End If
End Sub

Private Sub SetResize(ByVal LineIndex As Long, ByVal Value As Long)
    If LineIndex >= 0 And LineIndex < RULE_SLOTS Then
        m_rules(m_ruleCount - 1).ResizeAt(LineIndex) = Value
    End If
End Sub

Private Sub InitTables()
    If m_ready Then Exit Sub
    Glyphs3_Init
    AddColor "红色", 192, 0, 95, 0
    AddColor "蓝色", 192, 114, 107, 10
    AddColor "褐色", 306, 0, 119, 20'棕色
    AddColor "绿色", 306, 114, 131, 30
    AddColor "橙色", 420, 0, 143, 40
    AddColor "紫色", 420, 114, 155, 50 '紫罗兰
    AddColor "青色", 534, 0, 167, 60'凫蓝
    AddColor "粉色", 534, 114, 179, 70
    AddSize 4, 2, 1, 0, -1, -1, -1, -1
    AddSize 5, 2, 1, 0, -1, 19, -1, 18
    AddSize 5, 3, 4, 0, -1, -1, -1, -1
    SetHalf 2, 242
    SetHalf 3, 222
    SetResize 4, 206
    AddSize 5, 4, 8, 0, 16, -1, -1, -1
    SetHalf 5, 190
    SetHalf 6, 174
    SetHalf 7, 158
    SetResize 8, 142
    AddSize 5, 5, 9, 0, 21, -1, -1, -1
    SetResize 9, 126
    AddSize 7, 4, 8, 0, 16, -1, -1, -1
    SetHalf 7, 286
    SetResize 8, 270
    AddSize 7, 5, 11, 0, 21, -1, -1, -1
    SetHalf 9, 254
    SetHalf 10, 238
    SetResize 11, 222
    AddSize 10, 4, 8, 0, 25, -1, -1, -1
    SetResize 8, 466
    AddSize 10, 5, 11, 0, 30, -1, -1, -1
    SetHalf 9, 446
    SetHalf 10, 430
    SetResize 11, 414
    AddSize 10, 5, 11, 1, 29, 38, -1, 11
    ReDim m_words(0 To MAX_LINES * MAX_WORDS - 1)
    m_ready = True
End Sub

Public Function Homm3_ColorName(ByVal Index As Long) As String
    InitTables
    If Index < 0 Or Index >= m_colCount Then
        Homm3_ColorName = ""
    Else
        Homm3_ColorName = m_colName(Index)
    End If
End Function

' =====================================================================
'                          P U B L I C   A P I
' =====================================================================
' Renders the message window and returns a GDI+ bitmap handle.
' The caller owns the bitmap and must call GdipDisposeImage on it.
Public Function Homm3_Render(ByVal SpritePath As String, _
                             ByRef opt As HOMM3_OPTIONS, _
                             ByRef res As HOMM3_RESULT) As LongPtr
    Homm3_Render = 0
    If Not GdiPlus_Ensure() Then Exit Function
    InitTables
    If Not EnsureSprite(SpritePath) Then Exit Function

    m_opt = opt
    If m_opt.Zoom < 1 Then m_opt.Zoom = 1
    If m_opt.Zoom > 12 Then m_opt.Zoom = 12
    If m_opt.ColorIndex < 0 Or m_opt.ColorIndex >= m_colCount Then m_opt.ColorIndex = 0

    SplitWords m_opt.Text

    If m_opt.ManualSize Then
        SetupManualSize
    Else
        SetupAutoSize
    End If

    Dim cw As Long, ch As Long
    cw = m_popupW + IIf(m_opt.DrawShadow, SHADOW_X, 0)
    ch = m_popupH + IIf(m_opt.DrawShadow, SHADOW_Y, 0)

    Dim bmpLogical As LongPtr, gfx As LongPtr
    bmpLogical = GdiPlus_NewBitmap(cw, ch, gfx)
    If bmpLogical = 0 Then Exit Function
    m_gfx = gfx

    CJK_SetColors &HFFFFF3DE, &HFF000000, 1      ' 奶油字形，8向黑色轮廓
    DrawMessageWindow
    DrawText

    GdipDeleteGraphics gfx
    m_gfx = 0

    Dim bmpOut As LongPtr
    If m_opt.Zoom = 1 Then
        bmpOut = bmpLogical
    Else
        Dim zw As Long, zh As Long, gfx2 As LongPtr
        zw = cw * m_opt.Zoom
        zh = ch * m_opt.Zoom
        bmpOut = GdiPlus_NewBitmap(zw, zh, gfx2)
        If bmpOut = 0 Then
            GdipDisposeImage bmpLogical
            Exit Function
        End If
        GdipSetInterpolationMode gfx2, InterpolationModeNearestNeighbor
        GdipSetPixelOffsetMode gfx2, PixelOffsetModeHalf
        GdipDrawImageRectRectI gfx2, bmpLogical, 0, 0, zw, zh, _
                               0, 0, cw, ch, UnitPixel, 0, 0, 0
        GdipDeleteGraphics gfx2
        GdipDisposeImage bmpLogical
    End If

    res.PopupWidth = m_popupW
    res.PopupHeight = m_popupH
    res.OutWidth = cw * m_opt.Zoom
    res.OutHeight = ch * m_opt.Zoom
    res.LineCount = m_lineCount
    res.ScrollVisible = m_scrollVis
    res.Clipped = m_overflow
    Homm3_Render = bmpOut
End Function

Public Sub Homm3_Cleanup()
    ' 注意：这里千万不要再调 GdiPlus_Shutdown！
    If m_sprite <> 0 Then GdipDisposeImage m_sprite
    m_sprite = 0
    m_spriteFile = ""
End Sub

' =====================================================================
'                        S P R I T E   C A C H E
' =====================================================================
Private Function EnsureSprite(ByVal path As String) As Boolean
    If m_sprite <> 0 And m_spriteFile = path Then
        EnsureSprite = True
        Exit Function
    End If
    If m_sprite <> 0 Then GdipDisposeImage m_sprite
    m_sprite = 0
    m_spriteFile = ""
    If Len(Dir$(path)) = 0 Then Exit Function
    m_sprite = GdiPlus_LoadPng(path)
    If m_sprite = 0 Then Exit Function
    m_spriteFile = path
    EnsureSprite = True
End Function

' =====================================================================
'                          M E A S U R I N G
' =====================================================================
Private Function GetStringWidth(ByRef s As String) As Long
    Dim w As Long, i As Long, gi As Long, n As Long, cc As Long
    w = 0
    n = Len(s)
    For i = 1 To n
        If i = n And Mid$(s, i, 1) = " " Then Exit For   '修剪一个尾随空格
        cc = CharCodeAt(s, i)
        gi = Glyph3_Find(cc)
        If gi >= 0 Then
            If w <> 0 Then w = w + LETTER_SPACING
            w = w + gGlyphs3(gi).Width
            If gGlyphs3(gi).HasMarginLeft And i <> 1 Then w = w + gGlyphs3(gi).MarginLeft
            If gGlyphs3(gi).HasMarginRight Then w = w + gGlyphs3(gi).MarginRight
        ElseIf CJK_Ready() Then
            ' 不是sprite字体-用真正的Windows字体绘制
            If w <> 0 Then w = w + LETTER_SPACING
            w = w + CJK_Width(cc)
        End If
    Next i
    GetStringWidth = w
End Function

Private Function ButtonsVisible() As Boolean
    ButtonsVisible = (m_opt.ShowOK Or m_opt.ShowCancel)
End Function

' k = PAD_T / PAD_B / PAD_L / PAD_R
Private Function GetPadding(ByVal k As Long) As Long
    Dim v As Long
    If k = PAD_B And Not ButtonsVisible() Then
        GetPadding = GetPadding(PAD_T)
        Exit Function
    End If
    Select Case k
        Case PAD_T: v = IIf(m_size.PadT >= 0, m_size.PadT, PAD_TOP)
        Case PAD_B: v = IIf(m_size.PadB >= 0, m_size.PadB, PAD_BOTTOM)
        Case PAD_L: v = IIf(m_size.PadL >= 0, m_size.PadL, PAD_LEFT)
        Case PAD_R: v = IIf(m_size.PadR >= 0, m_size.PadR, PAD_RIGHT)
    End Select
    If k = PAD_T Or k = PAD_B Then
        GetPadding = v + H_BORDER_H
    Else
        GetPadding = v + V_BORDER_W
    End If
End Function

Private Function GetPopupWidth() As Long
    GetPopupWidth = m_popupW
End Function

Private Function GetPopupHeight() As Long
    GetPopupHeight = m_popupH
End Function

Private Function GetAvailWidth() As Long
    GetAvailWidth = m_popupW - GetPadding(PAD_L) - GetPadding(PAD_R)
End Function

Private Function GetRuleAt(ByVal LineIndex As Long, ByVal WantHalf As Boolean) As Long
    If LineIndex < 0 Or LineIndex >= RULE_SLOTS Then
        GetRuleAt = -1
        Exit Function
    End If
    If WantHalf Then
        GetRuleAt = m_size.HalfAt(LineIndex)
    Else
        GetRuleAt = m_size.ResizeAt(LineIndex)
    End If
End Function

' =====================================================================
'                       S I Z E   S E L E C T I O N
' =====================================================================
Private Sub SetupAutoSize()
    Dim i As Long
    For i = 0 To m_ruleCount - 1
        m_size = m_rules(i)
        m_scrollVis = m_size.Scroll
        m_popupW = m_size.W * BORDER_SIZE
        m_popupH = m_size.H * BORDER_SIZE
        WrapText
        If Not m_tryNext Then Exit For
    Next i
End Sub

Private Sub SetupManualSize()
    Dim w As Long, h As Long, i As Long
    w = m_opt.ManualWidth
    h = m_opt.ManualHeight
    If w < MIN_MANUAL_W Then w = MIN_MANUAL_W
    If h < MIN_MANUAL_H Then h = MIN_MANUAL_H
    If w > MAX_MANUAL_W Then w = MAX_MANUAL_W
    If h > MAX_MANUAL_H Then h = MAX_MANUAL_H

    m_popupW = w
    m_popupH = h

    With m_size
        .W = w
        .H = h
        .Scroll = m_opt.ForceScroll
        .PadT = -1: .PadR = -1: .PadB = -1: .PadL = -1
        For i = 0 To RULE_SLOTS - 1
            .HalfAt(i) = -1
            .ResizeAt(i) = -1
        Next i
        If m_opt.ForceScroll Then .PadR = 38      '为滚动条留出空间
    End With

    Dim cap As Long
    cap = (m_popupH - GetPadding(PAD_T) - GetPadding(PAD_B) - _
           IIf(ButtonsVisible(), BTN_H, 0)) \ LINE_HEIGHT
    If cap < 1 Then cap = 1
    If cap > MAX_LINES Then cap = MAX_LINES
    m_size.MaxTextLines = cap

    m_scrollVis = m_opt.ForceScroll
    WrapText
End Sub

' =====================================================================
'                        W O R D   W R A P P I N G
' =====================================================================
Private Sub WrapText()
    m_lineCount = 0
    m_overflow = False
    m_tryNext = False
    m_raiseHalf = False
    If gSplitCount = 0 Then Exit Sub

    Dim avail As Long, curCount As Long, spaceStr As String
    Dim wSpace As Long, wBlock As Long
    Dim i As Long, j As Long, block As String
    Dim blkIsNewLine As Boolean, blkIsSpace As Boolean
    Dim joint As String, halfAt As Long, resizeAt As Long

    avail = GetAvailWidth()
    curCount = 0
    spaceStr = ""
    wSpace = 0

    For i = 0 To gSplitCount - 1
        m_raiseHalf = False
        halfAt = GetRuleAt(m_lineCount, True)
        resizeAt = GetRuleAt(m_lineCount, False)

        block = gSplit(i)
        blkIsNewLine = IsBreakBlock(block)
        blkIsSpace = (Not blkIsNewLine) And IsSpaceBlock(block)

        joint = JoinCurrent(curCount) & block
        wBlock = GetStringWidth(joint)

        If blkIsNewLine Then
            PushLine curCount
            spaceStr = ""
        ElseIf blkIsSpace Then
            For j = 1 To Len(block)
                spaceStr = spaceStr & Mid$(block, j, 1)
                wSpace = GetStringWidth(JoinCurrent(curCount) & spaceStr)
                If wSpace > avail Then
                    AddWord curCount, spaceStr
                    PushLine curCount
                    spaceStr = ""
                End If
            Next j
            If spaceStr <> "" Then
                AddWord curCount, spaceStr
                spaceStr = ""
            End If
        Else
            If wBlock < avail Then
                AddWord curCount, block
            Else
                If curCount > 0 Then PushLine curCount
                If Len(block) > 1 And BlockHasFallback(block, 3) Then
                    BreakBlockByChar block, avail, curCount
                Else
                    AddWord curCount, block
                End If
            End If
        End If

        If resizeAt >= 0 Then
            If wSpace >= resizeAt Or wBlock >= resizeAt Then Exit For
        End If
        If halfAt >= 0 Then
            If wSpace >= halfAt Or wBlock >= halfAt Then m_raiseHalf = True
        End If
    Next i

    If curCount > 0 Then
        PushLine curCount
        Dim lastW As Long
        lastW = GetStringWidth(JoinLineOf(m_lineCount - 1))
        resizeAt = GetRuleAt(m_lineCount, False)
        halfAt = GetRuleAt(m_lineCount, True)
        If resizeAt >= 0 And lastW >= resizeAt Then m_tryNext = True
        If halfAt >= 0 Then m_raiseHalf = (lastW >= halfAt)
    End If

    If m_lineCount > m_size.MaxTextLines Then
        m_overflow = True
        m_tryNext = True
        m_lineCount = m_size.MaxTextLines
    End If
End Sub

Private Sub BreakBlockByChar(ByRef block As String, ByVal avail As Long, _
                             ByRef Count As Long)
    Dim k As Long, chunk As String
    chunk = ""
    For k = 1 To Len(block)
        If chunk <> "" Then
            If GetStringWidth(chunk & Mid$(block, k, 1)) > avail Then
                AddWord Count, chunk
                PushLine Count
                chunk = ""
            End If
        End If
        chunk = chunk & Mid$(block, k, 1)
    Next k
    If chunk <> "" Then AddWord Count, chunk
End Sub

Private Sub PushLine(ByRef Count As Long)
    If m_lineCount >= MAX_LINES Then
        Count = 0
        Exit Sub
    End If
    m_wordCount(m_lineCount) = Count
    m_lineCount = m_lineCount + 1
    Count = 0
End Sub

Private Sub AddWord(ByRef Count As Long, ByVal s As String)
    If Count >= MAX_WORDS Then Exit Sub
    m_words(m_lineCount * MAX_WORDS + Count) = s
    Count = Count + 1
End Sub

Private Function JoinCurrent(ByVal Count As Long) As String
    Dim s As String, k As Long, base As Long
    base = m_lineCount * MAX_WORDS
    s = ""
    For k = 0 To Count - 1
        s = s & m_words(base + k)
    Next k
    JoinCurrent = s
End Function

Private Function JoinLineOf(ByVal LineIndex As Long) As String
    Dim s As String, k As Long, base As Long
    If LineIndex < 0 Or LineIndex >= m_lineCount Then Exit Function
    base = LineIndex * MAX_WORDS
    s = ""
    For k = 0 To m_wordCount(LineIndex) - 1
        s = s & m_words(base + k)
    Next k
    JoinLineOf = s
End Function

' =====================================================================
'                          D R A W I N G
' =====================================================================
Private Sub Blit(ByVal dx As Long, ByVal dy As Long, ByVal dw As Long, ByVal dh As Long, _
                 ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long)
    If dw <= 0 Or dh <= 0 Or sw <= 0 Or sh <= 0 Then Exit Sub
    GdipDrawImageRectRectI m_gfx, m_sprite, dx, dy, dw, dh, sx, sy, sw, sh, UnitPixel, 0, 0, 0
End Sub

' 在dstW像素上水平重复BORDER_SIZE宽条带。
Private Sub TileH(ByVal sx As Long, ByVal sy As Long, ByVal dstW As Long, ByVal sh As Long, _
                  ByVal dx As Long, ByVal dy As Long)
    Dim remain As Long, w As Long, x As Long
    remain = dstW
    x = dx
    Do While remain > 0
        w = BORDER_SIZE
        If w > remain Then w = remain
        Blit x, dy, w, sh, sx, sy, w, sh
        x = x + w
        remain = remain - w
    Loop
End Sub

' 在dstH像素上垂直重复BORDER_SIZE高条。
Private Sub TileV(ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal dstH As Long, _
                  ByVal dx As Long, ByVal dy As Long)
    Dim remain As Long, h As Long, y As Long
    remain = dstH
    y = dy
    Do While remain > 0
        h = BORDER_SIZE
        If h > remain Then h = remain
        Blit dx, y, sw, h, sx, sy, sw, h
        y = y + h
        remain = remain - h
    Loop
End Sub

Private Sub DrawMessageWindow()
    If m_opt.DrawShadow Then DrawShadow
    DrawBackground
    DrawCorners
    DrawHorizontalBorders
    DrawVerticalBorders
    If ButtonsVisible() Then DrawButtons
End Sub

Private Sub DrawShadow()
    Dim br As LongPtr
    If GdipCreateSolidFill(&H19000000, br) <> GpOk Then Exit Sub
    GdipFillRectangleI m_gfx, br, SHADOW_X, SHADOW_Y, m_popupW, m_popupH
    GdipFillRectangleI m_gfx, br, SHADOW_X + 1, SHADOW_Y + 1, m_popupW - 2, m_popupH - 2
    GdipDeleteBrush br
End Sub

Private Sub DrawBackground()
    Dim x As Long, y As Long, w As Long, h As Long
    y = 0
    Do While y < m_popupH
        h = BG_SIZE
        If y + h > m_popupH Then h = m_popupH - y
        x = 0
        Do While x < m_popupW
            w = BG_SIZE
            If x + w > m_popupW Then w = m_popupW - x
            Blit x, y, w, h, SP_BG_X, SP_BG_Y, w, h
            x = x + w
        Loop
        y = y + h
    Loop
End Sub

Private Sub DrawCorners()
    Dim ix As Long, iy As Long
    Dim dx As Long, dy As Long, sx As Long, sy As Long
    Dim cx As Long, cy As Long, ci As Long
    ci = m_opt.ColorIndex
    For iy = 0 To 1
        For ix = 0 To 1
            If ix = 1 Then
                dx = m_popupW - BORDER_SIZE
                sx = BORDER_SIZE
                cx = m_popupW - 8 - COLOR_CORNER
            Else
                dx = 0
                sx = 0
                cx = 8
            End If
            If iy = 1 Then
                dy = m_popupH - BORDER_SIZE
                sy = BORDER_SIZE
                cy = m_popupH - 8 - COLOR_CORNER
            Else
                dy = 0
                sy = 0
                cy = 8
            End If
            Blit dx, dy, BORDER_SIZE, BORDER_SIZE, sx, sy, BORDER_SIZE, BORDER_SIZE
            Blit cx, cy, COLOR_CORNER, COLOR_CORNER, _
                 m_colCX(ci) + (1 + COLOR_CORNER) * ix, _
                 m_colCY(ci) + (1 + COLOR_CORNER) * iy, _
                 COLOR_CORNER, COLOR_CORNER
        Next ix
    Next iy
End Sub

Private Sub DrawHorizontalBorders()
    Dim span As Long, ci As Long
    ci = m_opt.ColorIndex
    span = m_popupW - 2 * BORDER_SIZE
    If span <= 0 Then Exit Sub
    TileH BORDER_SIZE * 2, 0, span, H_BORDER_H, BORDER_SIZE, 0
    TileH BORDER_SIZE * 2, H_BORDER_H + BORDER_SIZE, span, H_BORDER_H, _
          BORDER_SIZE, m_popupH - H_BORDER_H
    TileH BORDER_SIZE * 2, m_colHY(ci), span, H_COLOR_H, BORDER_SIZE, 8
    TileH BORDER_SIZE * 2, m_colHY(ci) + 1 + H_COLOR_H, span, H_COLOR_H, _
          BORDER_SIZE, m_popupH - H_BORDER_H + 2
End Sub

Private Sub DrawVerticalBorders()
    Dim span As Long, ci As Long
    ci = m_opt.ColorIndex
    span = m_popupH - 2 * BORDER_SIZE
    If span <= 0 Then Exit Sub
    TileV BORDER_SIZE * 2, H_BORDER_H, V_BORDER_W, span, 0, BORDER_SIZE
    TileV BORDER_SIZE * 3 - V_BORDER_W, H_BORDER_H, V_BORDER_W, span, _
          m_popupW - V_BORDER_W, BORDER_SIZE
    TileV m_colVX(ci), BORDER_SIZE * 2, V_COLOR_W, span, 8, BORDER_SIZE
    TileV m_colVX(ci), BORDER_SIZE * 2, V_COLOR_W, span, _
          m_popupW - V_BORDER_W + 2, BORDER_SIZE
End Sub

Private Sub DrawButtons()
    Dim n As Long, by As Long, bx As Long
    n = IIf(m_opt.ShowOK, 1, 0) + IIf(m_opt.ShowCancel, 1, 0)
    by = m_popupH - BTN_H - GetPadding(PAD_B)
    If by < 0 Then by = 0
    If m_opt.ShowOK Then
        If n = 1 Then
            bx = (m_popupW - BTN_W) \ 2
        Else
            bx = m_popupW \ 2 - (BTN_W + BTN_MARGIN \ 2)
        End If
        Blit bx, by, BTN_W, BTN_H, SP_OK_X, SP_OK_Y, BTN_W, BTN_H
    End If
    If m_opt.ShowCancel Then
        If n = 2 Then
            bx = (m_popupW + BTN_MARGIN) \ 2
        Else
            bx = (m_popupW - BTN_W) \ 2
        End If
        Blit bx, by, BTN_W, BTN_H, SP_OK_X, SP_CANCEL_Y, BTN_W, BTN_H
    End If
End Sub

Private Sub DrawScroll()
    Dim sx As Long, y0 As Long, y1 As Long, br As LongPtr
    sx = m_popupW - V_BORDER_W - SCROLL_SIDE - SCROLL_M_R
    y0 = H_BORDER_H + SCROLL_M_T
    y1 = m_popupH - GetPadding(PAD_B) - _
         IIf(ButtonsVisible(), BTN_H + SCROLL_M_B, 0)
    Blit sx, y0, SCROLL_SIDE, SCROLL_SIDE, SP_SC_UP_X, SP_SC_Y, SCROLL_SIDE, SCROLL_SIDE
    Blit sx, y0 + SCROLL_SIDE, SCROLL_SIDE, SCROLL_SIDE, SP_SC_MID_X, SP_SC_Y, SCROLL_SIDE, SCROLL_SIDE
    Blit sx, y1 - SCROLL_SIDE, SCROLL_SIDE, SCROLL_SIDE, SP_SC_DN_X, SP_SC_Y, SCROLL_SIDE, SCROLL_SIDE
    If GdipCreateSolidFill(&HFF000000, br) = GpOk Then
        GdipFillRectangleI m_gfx, br, sx, y0 + 32, SCROLL_SIDE, y1 - y0 - 48
        GdipDeleteBrush br
    End If
End Sub

' =====================================================================
'                            T E X T
' =====================================================================
Private Sub DrawText()
    Dim lineIdx As Long, wordIdx As Long, charIdx As Long
    Dim wordCount As Long, word As String, joint As String
    Dim curX As Long, x0 As Long, y As Long
    Dim gi As Long, cc As Long, avail As Long
    Dim skipWord As Boolean

    If m_scrollVis Then DrawScroll

    avail = GetAvailWidth()

    For lineIdx = 0 To m_lineCount - 1
        wordCount = m_wordCount(lineIdx)
        curX = 0
        If Not m_scrollVis Then
            joint = JoinLineOf(lineIdx)
            curX = Int((avail - GetStringWidth(joint)) / 2)
        End If

        For wordIdx = 0 To wordCount - 1
            word = m_words(lineIdx * MAX_WORDS + wordIdx)
            skipWord = (word = " " And wordIdx + 1 = wordCount)
            If Not skipWord Then
                For charIdx = 1 To Len(word)
                    cc = CharCodeAt(word, charIdx)
                    gi = Glyph3_Find(cc)
                    If gi >= 0 Then
                        If gGlyphs3(gi).HasMarginLeft Then
                            If wordIdx <> 0 Or charIdx <> 1 Or gGlyphs3(gi).MarginLeft > 0 Then
                                curX = curX + gGlyphs3(gi).MarginLeft
                            End If
                        End If
                        If gGlyphs3(gi).Drawable Then
                            x0 = GetPadding(PAD_L) + curX
                            y = GetLetterY(lineIdx, gi)
                            Blit x0, y, gGlyphs3(gi).Width, gGlyphs3(gi).Height, _
                                 gGlyphs3(gi).X, gGlyphs3(gi).Y, _
                                 gGlyphs3(gi).Width, gGlyphs3(gi).Height
                        End If
                        curX = curX + gGlyphs3(gi).Width
                        curX = curX + LETTER_SPACING
                        If gGlyphs3(gi).HasMarginRight Then curX = curX + gGlyphs3(gi).MarginRight
                    ElseIf CJK_Ready() Then
                        ' not in the sprite font - draw it with a real Windows font
                        CJK_Draw m_gfx, cc, GetPadding(PAD_L) + curX, _
                                 GetLineTopY(lineIdx) + LINE_HEIGHT
                        curX = curX + CJK_Width(cc) + LETTER_SPACING
                    End If
                Next charIdx
            End If
        Next wordIdx
    Next lineIdx
End Sub

' 文本行顶部的垂直偏移量。
' 自动模式会还原原版（颇为古怪的）《英雄无敌3》放置方式；
' 手动模式仅将整个模块居中，这正是你所期望的
' 窗口可以是任意大小。
Private Function GetLineTopY(ByVal lineIdx As Long) As Long
    Dim y As Long
    If Not m_opt.ManualSize Then
        If Not ButtonsVisible() Then
            GetLineTopY = RoundHalf((m_popupH - m_lineCount * LINE_HEIGHT) / 2) - 7
            Exit Function
        End If
        y = GetPadding(PAD_T) - (m_lineCount - 1) * (LINE_HEIGHT \ 2)
        If m_size.H > 2 Then y = y + 14 + (m_size.H - 2) * LINE_HEIGHT
        Select Case m_size.W
            Case 5
                If m_lineCount >= 9 Then
                    y = y + LINE_HEIGHT
                ElseIf m_lineCount >= 5 Then
                    y = y + LINE_HEIGHT \ 2
                End If
            Case 7
                If m_lineCount >= 9 Then
                    y = y + LINE_HEIGHT
                ElseIf m_lineCount >= 7 Then
                    y = y + LINE_HEIGHT \ 2
                End If
            Case 10
                If m_lineCount >= 9 Then y = y + LINE_HEIGHT \ 2
        End Select
        If m_raiseHalf Then y = y - LINE_HEIGHT \ 2
        GetLineTopY = y
        Exit Function
    End If

    ' ---- 手动尺寸：垂直居中文本块----
    Dim top0 As Long, bot0 As Long, blockH As Long, y0 As Long
    top0 = GetPadding(PAD_T)
    bot0 = m_popupH - GetPadding(PAD_B) - IIf(ButtonsVisible(), BTN_H, 0)
    blockH = m_lineCount * LINE_HEIGHT
    y0 = top0 + (bot0 - top0 - blockH) \ 2
    If y0 < top0 Then y0 = top0
    GetLineTopY = y0
End Function

' 一个精灵字形的最后一个Y。
Private Function GetLetterY(ByVal lineIdx As Long, ByVal gi As Long) As Long
    Dim y As Long
    y = GetLineTopY(lineIdx) + lineIdx * LINE_HEIGHT + (LINE_HEIGHT - gGlyphs3(gi).Height)
    If gGlyphs3(gi).HasTranslateY Then y = y + gGlyphs3(gi).TranslateY
    GetLetterY = y
End Function

Private Function RoundHalf(ByVal v As Double) As Long
    RoundHalf = CLng(Int(v + 0.5))
End Function
