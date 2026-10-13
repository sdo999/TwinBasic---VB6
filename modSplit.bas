Attribute VB_Name = "modSplit"
' =====================================================================
'modSplit.bas--将输入文本拆分为“块”
'《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
'块是一系列相同的字符类：
'符号-任何可打印的东西
'空格-空格/制表符
'break-一次硬线突破
'
'“这由HoMM2和HoMM3渲染器共享，是一个
'HommMessageGeneratorDispatcher.breakInputIntoWordsAndSpaces（）的直接端口。
' =====================================================================
Option Explicit

Public gSplit()     As String
Public gSplitCount  As Long

' ------------------------------------------------------------------
Public Sub SplitWords(ByVal Text As String)
    gSplitCount = 0
    ReDim gSplit(0 To 63)

    Dim n As Long, i As Long
    n = Len(Text)
    If n = 0 Then Exit Sub

    Dim phase As String, seq As String, c As String, ct As String
    phase = ""
    seq = ""

    For i = 1 To n
        c = Mid$(Text, i, 1)
        If IsBreakChar(c) Then
            ct = "break"
        ElseIf IsSpaceChar(c) Then
            ct = "space"
        Else
            ct = "symbol"
        End If
        If phase <> ct Then
            If seq <> "" Then PushPart(seq)
            phase = ct
            seq = ""
        End If
        seq = seq & c
        If ct = "break" Then
            PushPart(c)
            seq = ""
        End If
    Next i
    If seq <> "" Then PushPart(seq)
End Sub

Private Sub PushPart(ByVal s As String)
    If gSplitCount > UBound(gSplit) Then ReDim Preserve gSplit(0 To gSplitCount + 127)
    gSplit(gSplitCount) = s
    gSplitCount = gSplitCount + 1
End Sub

' ------------------------------------------------------------------
Public Function IsBreakChar(ByVal c As String) As Boolean
    Dim n As Long
    n = CharCodeAt(c, 1)
    IsBreakChar = (n = 10 Or n = 13)
End Function

Public Function IsSpaceChar(ByVal c As String) As Boolean
    Dim n As Long
    n = CharCodeAt(c, 1)
    Select Case n
        Case 9, 10, 11, 12, 13, 32, 160
            IsSpaceChar = True
        Case Else
            IsSpaceChar = False
    End Select
End Function

Public Function IsBreakBlock(ByRef s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsBreakChar(Mid$(s, i, 1)) Then Exit Function
    Next i
    IsBreakBlock = True
End Function

Public Function IsSpaceBlock(ByRef s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsSpaceChar(Mid$(s, i, 1)) Then Exit Function
    Next i
    IsSpaceBlock = True
End Function

' 索引位置处从1开始的字符的Unicode码点。
Public Function CharCodeAt(ByRef s As String, ByVal Index As Long) As Long
    Dim c As Long
    c = AscW(Mid$(s, Index, 1))
    If c < 0 Then c = c + 65536
    CharCodeAt = c
End Function

'当块中至少有一个字符不在
'位图字体--即它必须经过CJK/fallback渲染器
'因此需要字符级换行。
Public Function BlockHasFallback(ByRef s As String, ByVal GlyphTable As Long) As Boolean
    Dim i As Long, gi As Long
    For i = 1 To Len(s)
        If GlyphTable = 3 Then
            gi = Glyph3_Find(CharCodeAt(s, i))
        Else
            gi = Glyph2_Find(CharCodeAt(s, i))
        End If
        If gi < 0 Then
            BlockHasFallback = True
            Exit Function
        End If
    Next i
    BlockHasFallback = False
End Function
