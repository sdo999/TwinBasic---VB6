Attribute VB_Name = "modHomm2"
' =====================================================================
'modHomm2.bas--魔法门英雄II风格信息窗口
'《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
' “js/Homm2MessageGenerator.js的端口（由Alexei Sergienko编写），带有
' HoMM3渲染器获得了相同的附加功能：
' '*精灵字体（中文，…）中缺少的字符是
' '通过modCJK绘制而不是丢弃
' *窗口高度可以强制为任何值
' '*导出时任意整数缩放
' '
' 关于原始JavaScript的两个注意事项，都固定在这里：
' '-drawChar（）根据*HoMM3*字形测试了字符
' “表，所以HoMM2只测量了字符，但从未绘制。
' '-homm2-spite.png（y=499）中的“字母阴影”条为
' 空的，所以影子通道总是禁止通行。港口仍然
' “发布它；它什么都不花。
'
'需要：modGdiPlus.bas、modSplit.bas、modCJK.bas、modGlyphs2.bas，
'img/homm2-spite.png
' =====================================================================
Option Explicit

Private Const LINE_HEIGHT     As Long = 16
Private Const LETTER_SPACING  As Long = 2
Private Const LETTER_CW       As Long = 19      ' 子画面单元格宽度
Private Const LTR_OX          As Long = 1
Private Const LTR_OY          As Long = 313
Private Const LTR_SOY         As Long = 499

Private Const BTN_W           As Long = 95
Private Const BTN_H           As Long = 25
Private Const BTN_MARGIN      As Long = 28

Private Const MAX_LINES       As Long = 64
Private Const MAX_WORDS       As Long = 256
Private Const MAX_MIDDLES     As Long = 40

Private Const PAD_TOP         As Long = 0
Private Const PAD_BOTTOM      As Long = 1
Private Const PAD_LEFT        As Long = 2
Private Const PAD_RIGHT       As Long = 3
Private Const PAD_LEFT_SHADOW As Long = 4

Public Const HOMM2_THEME_GOOD As Long = 0
Public Const HOMM2_THEME_EVIL As Long = 1

' =====================================================================
Public Type HOMM2_OPTIONS
    Text         As String
    Theme        As Long          ' 0 = 善, 1 = 恶
    ShowOK       As Boolean
    ShowCancel   As Boolean
    DrawShadow   As Boolean
    ManualHeight As Boolean       ' True => 使用强制高度
    ForcedHeight As Long
    Zoom         As Long          ' 1..12
End Type

Public Type HOMM2_RESULT
    PopupWidth   As Long
    PopupHeight  As Long
    OutWidth     As Long
    OutHeight    As Long
    LineCount    As Long
    Middles      As Long
    Clipped      As Boolean
End Type

Private Type SEC_SHADOW
    ox As Long
    oy As Long
    sw As Long
    sh As Long
    px As Long
    py As Long
End Type

Private Type SEC_PART
    ox      As Long
    oy      As Long
    w       As Long
    h       As Long
    shCount As Long
    sh(0 To 2) As SEC_SHADOW
End Type

Private Type SEC_THEME
    Name         As String
    Width        As Long
    ShadowW      As Long
    padTop       As Long
    padRight     As Long
    padBottom    As Long
    padBottomBtn As Long
    padLeft      As Long
    btnOKX       As Long
    btnOKY       As Long
    btnCancelX   As Long
    btnCancelY   As Long
    top          As SEC_PART
    middle       As SEC_PART
    bottom       As SEC_PART
End Type

' ----模块状态 ---------------------------------------------------
Private m_opt        As HOMM2_OPTIONS
Private m_th         As SEC_THEME
Private m_good       As SEC_THEME
Private m_evil       As SEC_THEME

Private m_sprite     As LongPtr
Private m_spriteFile As String
Private m_gfx        As LongPtr

Private m_popupW     As Long
Private m_popupH     As Long
Private m_middles    As Long

Private m_lineCount  As Long
Private m_wordCount(0 To MAX_LINES - 1) As Long
Private m_words()    As String

Private m_overflow   As Boolean
Private m_ready      As Boolean

' =====================================================================
'                            T H E M E S
' =====================================================================
Private Sub InitThemes()
    If m_ready Then Exit Sub
    Glyphs2_Init

    ' ---------------- 善 ----------------
    m_good.Name = "good"
    m_good.Width = 290
    m_good.ShadowW = 17
    m_good.padTop = 64
    m_good.padRight = 32
    m_good.padBottom = 32
    m_good.padBottomBtn = 58
    m_good.padLeft = 32
    m_good.btnOKX = 345
    m_good.btnOKY = 243
    m_good.btnCancelX = 441
    m_good.btnCancelY = 243

    m_good.top.ox = 346: m_good.top.oy = 0:   m_good.top.w = 290: m_good.top.h = 99
    m_good.top.shCount = 2
    m_good.top.sh(0).ox = 312: m_good.top.sh(0).oy = 18
    m_good.top.sh(0).sw = 33:  m_good.top.sh(0).sh = 20
    m_good.top.sh(0).px = 97:  m_good.top.sh(0).py = 18
    m_good.top.sh(1).ox = 328: m_good.top.sh(1).oy = 50
    m_good.top.sh(1).sw = 16:  m_good.top.sh(1).sh = 49
    m_good.top.sh(1).px = -15: m_good.top.sh(1).py = 50

    m_good.middle.ox = 346: m_good.middle.oy = 100: m_good.middle.w = 290: m_good.middle.h = 45
    m_good.middle.shCount = 1
    m_good.middle.sh(0).ox = 327: m_good.middle.sh(0).oy = 100
    m_good.middle.sh(0).sw = 17:  m_good.middle.sh(0).sh = 45
    m_good.middle.sh(0).px = -16: m_good.middle.sh(0).py = 0

    m_good.bottom.ox = 346: m_good.bottom.oy = 146: m_good.bottom.w = 290: m_good.bottom.h = 65
    m_good.bottom.shCount = 2
    m_good.bottom.sh(0).ox = 327: m_good.bottom.sh(0).oy = 146
    m_good.bottom.sh(0).sw = 17:  m_good.bottom.sh(0).sh = 61
    m_good.bottom.sh(0).px = -16: m_good.bottom.sh(0).py = 0
    m_good.bottom.sh(1).ox = 331: m_good.bottom.sh(1).oy = 217
    m_good.bottom.sh(1).sw = 288: m_good.bottom.sh(1).sh = 20
    m_good.bottom.sh(1).px = -15: m_good.bottom.sh(1).py = 61

    ' ---------------- 恶----------------
    m_evil.Name = "evil"
    m_evil.Width = 288
    m_evil.ShadowW = 16
    m_evil.padTop = 53
    m_evil.padRight = 32
    m_evil.padBottom = 32
    m_evil.padBottomBtn = 61
    m_evil.padLeft = 32
    m_evil.btnOKX = 19
    m_evil.btnOKY = 243
    m_evil.btnCancelX = 115
    m_evil.btnCancelY = 243

    m_evil.top.ox = 19: m_evil.top.oy = 11:  m_evil.top.w = 288: m_evil.top.h = 88
    m_evil.top.shCount = 2
    m_evil.top.sh(0).ox = 3:   m_evil.top.sh(0).oy = 27
    m_evil.top.sh(0).sw = 14:  m_evil.top.sh(0).sh = 10
    m_evil.top.sh(0).px = 107: m_evil.top.sh(0).py = 16
    m_evil.top.sh(1).ox = 0:   m_evil.top.sh(1).oy = 48
    m_evil.top.sh(1).sw = 16:  m_evil.top.sh(1).sh = 51
    m_evil.top.sh(1).px = -16: m_evil.top.sh(1).py = 37

    m_evil.middle.ox = 19: m_evil.middle.oy = 100: m_evil.middle.w = 288: m_evil.middle.h = 45
    m_evil.middle.shCount = 1
    m_evil.middle.sh(0).ox = 0:   m_evil.middle.sh(0).oy = 100
    m_evil.middle.sh(0).sw = 16:  m_evil.middle.sh(0).sh = 45
    m_evil.middle.sh(0).px = -16: m_evil.middle.sh(0).py = 0

    m_evil.bottom.ox = 19: m_evil.bottom.oy = 146: m_evil.bottom.w = 288: m_evil.bottom.h = 68
    m_evil.bottom.shCount = 2
    m_evil.bottom.sh(0).ox = 0:   m_evil.bottom.sh(0).oy = 146
    m_evil.bottom.sh(0).sw = 16:  m_evil.bottom.sh(0).sh = 61
    m_evil.bottom.sh(0).px = -16: m_evil.bottom.sh(0).py = 0
    m_evil.bottom.sh(1).ox = 3:   m_evil.bottom.sh(1).oy = 216
    m_evil.bottom.sh(1).sw = 288: m_evil.bottom.sh(1).sh = 20
    m_evil.bottom.sh(1).px = -16: m_evil.bottom.sh(1).py = 61

    ReDim m_words(0 To MAX_LINES * MAX_WORDS - 1)
    m_ready = True
End Sub

' =====================================================================
'                          P U B L I C   A P I
' =====================================================================
Public Function Homm2_Render(ByVal SpritePath As String, _
                             ByRef opt As HOMM2_OPTIONS, _
                             ByRef res As HOMM2_RESULT) As LongPtr
    Homm2_Render = 0
    If Not GdiPlus_Ensure() Then Exit Function
    InitThemes
    If Not EnsureSprite(SpritePath) Then Exit Function

    m_opt = opt
    If m_opt.Zoom < 1 Then m_opt.Zoom = 1
    If m_opt.Zoom > 12 Then m_opt.Zoom = 12
    If m_opt.Theme <> HOMM2_THEME_EVIL Then m_opt.Theme = HOMM2_THEME_GOOD

    If m_opt.Theme = HOMM2_THEME_EVIL Then
        m_th = m_evil
    Else
        m_th = m_good
    End If

    SplitWords m_opt.Text
    WrapText
    SetSize

    Dim cw As Long, ch As Long
    cw = m_popupW + IIf(m_opt.DrawShadow, m_th.ShadowW, 0)
    ch = m_popupH + IIf(m_opt.DrawShadow, m_th.bottom.sh(1).sh, 0)

    Dim bmpLogical As LongPtr, gfx As LongPtr
    bmpLogical = GdiPlus_NewBitmap(cw, ch, gfx)
    If bmpLogical = 0 Then Exit Function
    m_gfx = gfx

    CJK_SetColors &HFFFCFCFC, &H90000000, 2       '近白色字形，投下阴影
    DrawSections
    If ButtonsVisible() Then DrawButtons
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
    res.Middles = m_middles
    res.Clipped = m_overflow
    Homm2_Render = bmpOut
End Function

Public Sub Homm2_Cleanup()
    If m_sprite <> 0 Then GdipDisposeImage m_sprite
    m_sprite = 0
    m_spriteFile = ""
End Sub

Public Function Homm2_ThemeWidth(ByVal Theme As Long) As Long
    InitThemes
    If Theme = HOMM2_THEME_EVIL Then
        Homm2_ThemeWidth = m_evil.Width
    Else
        Homm2_ThemeWidth = m_good.Width
    End If
End Function

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

Private Function ButtonsVisible() As Boolean
    ButtonsVisible = (m_opt.ShowOK Or m_opt.ShowCancel)
End Function

' k = PAD_TOP / PAD_BOTTOM / PAD_LEFT / PAD_RIGHT / PAD_LEFT_SHADOW
Private Function GetPadding(ByVal k As Long) As Long
    Select Case k
        Case PAD_TOP
            GetPadding = m_th.padTop
        Case PAD_BOTTOM
            If ButtonsVisible() Then
                GetPadding = m_th.padBottomBtn
            Else
                GetPadding = m_th.padBottom
            End If
        Case PAD_LEFT
            GetPadding = m_th.padLeft
        Case PAD_RIGHT
            GetPadding = m_th.padRight
        Case PAD_LEFT_SHADOW
            GetPadding = m_th.padLeft + IIf(m_opt.DrawShadow, m_th.ShadowW, 0)
    End Select
End Function

Private Function GetAvailWidth() As Long
    GetAvailWidth = m_th.Width - GetPadding(PAD_LEFT) - GetPadding(PAD_RIGHT)
End Function

' =====================================================================
'                       S I Z E   S E L E C T I O N
' =====================================================================
Private Sub SetSize()
    Dim textH As Long, i As Long, n As Long, spare As Long

    textH = m_lineCount * LINE_HEIGHT

    If m_opt.ManualHeight Then
        Dim h As Long
        h = m_opt.ForcedHeight
        If h < m_th.top.h + m_th.bottom.h Then h = m_th.top.h + m_th.bottom.h
        n = (h - m_th.top.h - m_th.bottom.h + m_th.middle.h - 1) \ m_th.middle.h
        If n < 0 Then n = 0
        If n > MAX_MIDDLES Then n = MAX_MIDDLES
        m_middles = n
    Else
        spare = m_th.top.h - GetPadding(PAD_TOP) + m_th.bottom.h - GetPadding(PAD_BOTTOM)
        m_middles = 9                       ' 超长文本安全网
        For i = 0 To 9
            If i > 0 Then spare = spare + m_th.middle.h
            If spare >= textH Then
                m_middles = i
                Exit For
            End If
        Next i
    End If

    m_popupW = m_th.Width
    m_popupH = m_th.top.h + m_th.bottom.h + m_middles * m_th.middle.h

    Dim cap As Long
    cap = (m_popupH - GetPadding(PAD_TOP) - GetPadding(PAD_BOTTOM)) \ LINE_HEIGHT
    If cap < 1 Then cap = 1
    If m_lineCount > cap Then
        m_overflow = True
        m_lineCount = cap
    End If
End Sub

' =====================================================================
'                        W O R D   W R A P P I N G
' =====================================================================
Private Sub WrapText()
    m_lineCount = 0
    m_overflow = False
    If gSplitCount = 0 Then Exit Sub

    Dim avail As Long, curCount As Long, spaceStr As String
    Dim i As Long, j As Long, block As String
    Dim blkIsNewLine As Boolean, blkIsSpace As Boolean
    Dim joint As String

    avail = GetAvailWidth()
    curCount = 0
    spaceStr = ""

    For i = 0 To gSplitCount - 1
        block = gSplit(i)
        blkIsNewLine = IsBreakBlock(block)
        blkIsSpace = (Not blkIsNewLine) And IsSpaceBlock(block)

        joint = JoinCurrent(curCount) & block

        If blkIsNewLine Then
            PushLine curCount
            spaceStr = ""
        ElseIf blkIsSpace Then
            For j = 1 To Len(block)
                spaceStr = spaceStr & Mid$(block, j, 1)
                If GetStringWidth(JoinCurrent(curCount) & spaceStr) > avail Then
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
            If GetStringWidth(joint) < avail Then
                AddWord curCount, block
            Else
                If curCount > 0 Then PushLine curCount
                If Len(block) > 1 And BlockHasFallback(block, 2) Then
                    BreakBlockByChar block, avail, curCount
                Else
                    AddWord curCount, block
                End If
            End If
        End If
    Next i

    If curCount > 0 Then PushLine curCount
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

Private Function GetStringWidth(ByRef s As String) As Long
    Dim w As Long, i As Long, gi As Long, n As Long, cc As Long
    w = 0
    n = Len(s)
    For i = 1 To n
        If i = n And Mid$(s, i, 1) = " " Then Exit For   ' 修剪一个尾随空格
        cc = CharCodeAt(s, i)
        gi = Glyph2_Find(cc)
        If gi >= 0 Then
            If w <> 0 Then w = w + LETTER_SPACING
            w = w + gGlyphs2(gi).Width
        ElseIf CJK_Ready() Then
            If w <> 0 Then w = w + LETTER_SPACING
            w = w + CJK_Width(cc)
        End If
    Next i
    GetStringWidth = w
End Function

' =====================================================================
'                          D R A W I N G
' =====================================================================
Private Sub Blit(ByVal dx As Long, ByVal dy As Long, ByVal dw As Long, ByVal dh As Long, _
                 ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long)
    If dw <= 0 Or dh <= 0 Or sw <= 0 Or sh <= 0 Then Exit Sub
    GdipDrawImageRectRectI m_gfx, m_sprite, dx, dy, dw, dh, sx, sy, sw, sh, UnitPixel, 0, 0, 0
End Sub

Private Sub DrawSections()
    Dim x0 As Long, y As Long, j As Long
    x0 = IIf(m_opt.DrawShadow, m_th.ShadowW, 0)
    y = 0
    DrawPart m_th.top, x0, y
    y = y + m_th.top.h
    For j = 1 To m_middles
        DrawPart m_th.middle, x0, y
        y = y + m_th.middle.h
    Next j
    DrawPart m_th.bottom, x0, y
End Sub

Private Sub DrawPart(ByRef p As SEC_PART, ByVal x0 As Long, ByVal y As Long)
    Blit x0, y, p.w, p.h, p.ox, p.oy, p.w, p.h
    If m_opt.DrawShadow Then
        Dim k As Long
        For k = 0 To p.shCount - 1
            Blit x0 + p.sh(k).px, y + p.sh(k).py, p.sh(k).sw, p.sh(k).sh, _
                 p.sh(k).ox, p.sh(k).oy, p.sh(k).sw, p.sh(k).sh
        Next k
    End If
End Sub

Private Sub DrawButtons()
    Dim n As Long, by As Long, bx As Long, x0 As Long
    x0 = IIf(m_opt.DrawShadow, m_th.ShadowW, 0)
    n = IIf(m_opt.ShowOK, 1, 0) + IIf(m_opt.ShowCancel, 1, 0)
    by = m_popupH - GetPadding(PAD_BOTTOM)
    If by < 0 Then by = 0

    If m_opt.ShowOK Then
        If n = 1 Then
            bx = x0 + (m_popupW - BTN_W) \ 2
        Else
            bx = x0 + m_popupW \ 2 - (BTN_W + BTN_MARGIN \ 2)
        End If
        Blit bx, by, BTN_W, BTN_H, m_th.btnOKX, m_th.btnOKY, BTN_W, BTN_H
    End If

    If m_opt.ShowCancel Then
        If n = 2 Then
            bx = x0 + (m_popupW + BTN_MARGIN) \ 2
        Else
            bx = x0 + (m_popupW - BTN_W) \ 2
        End If
        Blit bx, by, BTN_W, BTN_H, m_th.btnCancelX, m_th.btnCancelY, BTN_W, BTN_H
    End If
End Sub

Private Sub DrawText()
    Dim avail As Long, li As Long, wi As Long, ci As Long
    Dim wc As Long, curX As Long, joint As String, word As String
    Dim gi As Long, cc As Long, lineTopY As Long, drawX As Long
    Dim eUp As Long, eDown As Long, hh As Long, cx As Long, cy As Long, sy As Long
    Dim skipWord As Boolean

    avail = GetAvailWidth()

    For li = 0 To m_lineCount - 1
        wc = m_wordCount(li)
        joint = JoinLineOf(li)
        curX = Int((avail - GetStringWidth(joint)) / 2)
        lineTopY = GetLineTopY(li)

        For wi = 0 To wc - 1
            word = m_words(li * MAX_WORDS + wi)
            skipWord = (word = " " And wi + 1 = wc)
            If Not skipWord Then
                For ci = 1 To Len(word)
                    cc = CharCodeAt(word, ci)
                    gi = Glyph2_Find(cc)
                    If gi >= 0 Then
                        If gGlyphs2(gi).Drawable Then
                            If gGlyphs2(gi).HasExtraUp Then eUp = gGlyphs2(gi).ExtraUp Else eUp = 0
                            If gGlyphs2(gi).HasExtraDown Then eDown = gGlyphs2(gi).ExtraDown Else eDown = 0
                            hh = LINE_HEIGHT + eUp + eDown
                            cx = LTR_OX + gGlyphs2(gi).X * LETTER_CW
                            cy = LTR_OY + gGlyphs2(gi).Y * LINE_HEIGHT
                            sy = LTR_SOY + gGlyphs2(gi).Y * LINE_HEIGHT
                            drawX = GetPadding(PAD_LEFT_SHADOW) + curX
                            Blit drawX, lineTopY - eUp, gGlyphs2(gi).Width, hh, _
                                 cx, cy - eUp, gGlyphs2(gi).Width, hh
                            Blit drawX - 1, lineTopY - eUp + 1, gGlyphs2(gi).Width, hh - 1, _
                                 cx - 1, sy - eUp + 1, gGlyphs2(gi).Width, hh - 1
                        End If
                        curX = curX + gGlyphs2(gi).Width + LETTER_SPACING
                    ElseIf CJK_Ready() Then
                        CJK_Draw m_gfx, cc, GetPadding(PAD_LEFT_SHADOW) + curX, _
                                 lineTopY + LINE_HEIGHT
                        curX = curX + CJK_Width(cc) + LETTER_SPACING
                    End If
                Next ci
            End If
        Next wi
    Next li
End Sub

Private Function GetLineTopY(ByVal lineIdx As Long) As Long
    If Not ButtonsVisible() Then
        GetLineTopY = lineIdx * LINE_HEIGHT + GetPadding(PAD_TOP) + _
            RoundHalf((m_popupH - GetPadding(PAD_TOP) - GetPadding(PAD_BOTTOM) - _
                       m_lineCount * LINE_HEIGHT) / 2)
        Exit Function
    End If
    GetLineTopY = GetPadding(PAD_TOP) + lineIdx * LINE_HEIGHT
End Function

Private Function RoundHalf(ByVal v As Double) As Long
    RoundHalf = CLng(Int(v + 0.5))
End Function
