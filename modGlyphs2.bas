Attribute VB_Name = "modGlyphs2"
' =====================================================================
'modGlyphs2.bas-HoMM2位图字体度量
'《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
'从js/Homm2MessageGenerator.js（Homm2MessageGenerator.letters）移植而来
'每个字形都由Unicode码位存储，因此源文件保持不变
'“纯ASCII，适用于ANSI和Unicode版本。
'
'  Sprite sheet: img/homm2-sprite.png   (654 x 680)
' =====================================================================
Option Explicit

Public Type GLYPH2_INFO
    Code        As Long      ' Unicode code point
    Width       As Long      ' advance / glyph width in px
    X           As Long      ' column inside the sprite sheet
    Y           As Long      ' row inside the sprite sheet
    ExtraUp      As Long
    ExtraDown    As Long
    HasExtraUp    As Boolean
    HasExtraDown  As Boolean
    Drawable    As Boolean   ' has a bitmap to blit
End Type

Public gGlyphs2()    As GLYPH2_INFO
Public gGlyphs2Count As Long

Private m_AddIndex  As Long
Private m_Ready     As Boolean

' ------------------------------------------------------------------
Public Sub Glyphs2_Init()
    If m_Ready Then Exit Sub
    m_AddIndex = 0
    ReDim gGlyphs2(0 To 511)
    Add2 &H20, 4, -1, -1, 0, 0, 0
    Add2 &H21, 2, &H1, &H0, 0, 0, 0
    Add2 &H22, 6, &H2, &H0, 0, 0, 0
    Add2 &H24, 8, &H4, &H0, 0, 0, 0
    Add2 &H25, 11, &H5, &H0, 0, 0, 0
    Add2 &H26, 10, &H6, &H0, 0, 0, 0
    Add2 &H27, 2, &H7, &H0, 0, 0, 0
    Add2 &H28, 4, &H8, &H0, 0, 0, 0
    Add2 &H29, 4, &H9, &H0, 0, 0, 0
    Add2 &H2A, 11, &HA, &H0, 0, 0, 0
    Add2 &H2B, 8, &HB, &H0, 0, 0, 0
    Add2 &H2C, 3, &HC, &H0, 0, 0, 0
    Add2 &H2D, 9, &HD, &H0, 0, 0, 0
    Add2 &H2E, 2, &H0, &H0, 0, 0, 0
    Add2 &H2F, 7, &HF, &H0, 0, 0, 0
    Add2 &H30, 8, &H0, &H1, 0, 0, 0
    Add2 &H31, 4, &H1, &H1, 0, 0, 0
    Add2 &H32, 7, &H2, &H1, 0, 0, 0
    Add2 &H33, 7, &H3, &H1, 0, 0, 0
    Add2 &H34, 8, &H4, &H1, 0, 0, 0
    Add2 &H35, 7, &H5, &H1, 0, 0, 0
    Add2 &H36, 7, &H6, &H1, 0, 0, 0
    Add2 &H37, 7, &H7, &H1, 0, 0, 0
    Add2 &H38, 7, &H8, &H1, 0, 0, 0
    Add2 &H39, 7, &H9, &H1, 0, 0, 0
    Add2 &H3A, 2, &H0, &H2, 0, 0, 0
    Add2 &H3B, 3, &H1, &H2, 0, 0, 0
    Add2 &H3D, 9, &H3, &H2, 0, 0, 0
    Add2 &H3F, 8, &H5, &H2, 0, 0, 0
    Add2 &H41, 14, &H0, &H3, 0, 0, 0
    Add2 &H42, 10, &H1, &H3, 0, 0, 0
    Add2 &H43, 12, &H2, &H3, 0, 0, 0
    Add2 &H44, 11, &H3, &H3, 0, 0, 0
    Add2 &H45, 10, &H4, &H3, 0, 0, 0
    Add2 &H46, 10, &H5, &H3, 0, 0, 0
    Add2 &H47, 13, &H6, &H3, 0, 0, 0
    Add2 &H48, 11, &H7, &H3, 0, 0, 0
    Add2 &H49, 6, &H8, &H3, 0, 0, 0
    Add2 &H4A, 9, &H9, &H3, 0, 0, 0
    Add2 &H4B, 11, &HA, &H3, 0, 0, 0
    Add2 &H4C, 10, &HB, &H3, 0, 0, 0
    Add2 &H4D, 15, &HC, &H3, 0, 0, 0
    Add2 &H4E, 14, &HD, &H3, 0, 0, 0
    Add2 &H4F, 12, &HE, &H3, 0, 0, 0
    Add2 &H50, 10, &HF, &H3, 0, 0, 0
    Add2 &H51, 14, &H10, &H3, 0, 0, 0
    Add2 &H52, 11, &H11, &H3, 0, 0, 0
    Add2 &H53, 9, &H12, &H3, 0, 0, 0
    Add2 &H54, 11, &H13, &H3, 0, 0, 0
    Add2 &H55, 11, &H14, &H3, 0, 0, 0
    Add2 &H56, 12, &H15, &H3, 0, 0, 0
    Add2 &H57, 16, &H16, &H3, 0, 0, 0
    Add2 &H58, 15, &H17, &H3, 0, 0, 0
    Add2 &H59, 13, &H18, &H3, 0, 0, 0
    Add2 &H5A, 10, &H19, &H3, 0, 0, 0
    Add2 &H5B, 5, &H0, &H4, 0, 0, 0
    Add2 &H5D, 5, &H2, &H4, 0, 0, 0
    Add2 &H5F, 8, &H4, &H4, 0, 0, 0
    Add2 &H61, 9, &H0, &H5, 0, 0, 0
    Add2 &H62, 10, &H1, &H5, 0, 0, 0
    Add2 &H63, 8, &H2, &H5, 0, 0, 0
    Add2 &H64, 9, &H3, &H5, 0, 0, 0
    Add2 &H65, 8, &H4, &H5, 0, 0, 0
    Add2 &H66, 7, &H5, &H5, 0, 0, 0
    Add2 &H67, 9, &H6, &H5, 0, 0, 0
    Add2 &H68, 9, &H7, &H5, 0, 0, 0
    Add2 &H69, 4, &H8, &H5, 0, 0, 0
    Add2 &H6A, 7, &H9, &H5, 0, 0, 0
    Add2 &H6B, 9, &HA, &H5, 0, 0, 0
    Add2 &H6C, 4, &HB, &H5, 0, 0, 0
    Add2 &H6D, 14, &HC, &H5, 0, 0, 0
    Add2 &H6E, 9, &HD, &H5, 0, 0, 0
    Add2 &H6F, 8, &HE, &H5, 0, 0, 0
    Add2 &H70, 9, &HF, &H5, 0, 0, 0
    Add2 &H71, 10, &H10, &H5, 0, 0, 0
    Add2 &H72, 8, &H11, &H5, 0, 0, 0
    Add2 &H73, 7, &H12, &H5, 0, 0, 0
    Add2 &H74, 6, &H13, &H5, 0, 0, 0
    Add2 &H75, 9, &H14, &H5, 0, 0, 0
    Add2 &H76, 8, &H15, &H5, 0, 0, 0
    Add2 &H77, 12, &H16, &H5, 0, 0, 0
    Add2 &H78, 9, &H17, &H5, 0, 0, 0
    Add2 &H79, 10, &H18, &H5, 0, 0, 0
    Add2 &H7A, 8, &H19, &H5, 0, 0, 0
    Add2 &HC1, 14, &H1A, &HA, 4, 0, 1
    Add2 &HC9, 10, &H1B, &HA, 4, 0, 1
    Add2 &HCD, 6, &H1C, &HA, 4, 0, 1
    Add2 &HD3, 12, &HE, &HA, 4, 0, 1
    Add2 &HD6, 12, &H1D, &HA, 4, 0, 1
    Add2 &HDA, 11, &H1F, &HA, 4, 0, 1
    Add2 &HDC, 11, &H21, &HA, 4, 0, 1
    Add2 &HE1, 9, &H10, &HA, 0, 0, 0
    Add2 &HE9, 8, &H13, &HA, 0, 0, 0
    Add2 &HED, 4, &H14, &HA, 0, 0, 0
    Add2 &HF3, 8, &H5, &HA, 0, 0, 0
    Add2 &HF6, 8, &H15, &HA, 0, 0, 0
    Add2 &HFA, 9, &H17, &HA, 0, 0, 0
    Add2 &HFC, 9, &H19, &HA, 0, 0, 0
    Add2 &H104, 14, &H9, &HA, 0, 0, 0
    Add2 &H105, 9, &H0, &HA, 0, 0, 0
    Add2 &H106, 12, &HA, &HA, 0, 0, 0
    Add2 &H107, 8, &H1, &HA, 0, 0, 0
    Add2 &H118, 10, &HB, &HA, 0, 1, 2
    Add2 &H119, 8, &H2, &HA, 0, 0, 0
    Add2 &H141, 10, &HC, &HA, 0, 0, 0
    Add2 &H142, 7, &H3, &HA, 0, 0, 0
    Add2 &H143, 14, &HD, &HA, 4, 0, 1
    Add2 &H144, 9, &H4, &HA, 0, 0, 0
    Add2 &H150, 12, &H1E, &HA, 4, 0, 1
    Add2 &H151, 8, &H16, &HA, 0, 0, 0
    Add2 &H15A, 9, &HF, &HA, 4, 0, 1
    Add2 &H15B, 7, &H6, &HA, 0, 0, 0
    Add2 &H170, 11, &H20, &HA, 4, 0, 1
    Add2 &H171, 9, &H18, &HA, 0, 0, 0
    Add2 &H179, 10, &H11, &HA, 4, 0, 1
    Add2 &H17A, 8, &H7, &HA, 0, 0, 0
    Add2 &H17B, 10, &H12, &HA, 4, 0, 1
    Add2 &H17C, 8, &H8, &HA, 0, 0, 0
    Add2 &H401, 10, &H20, &H7, 0, 0, 0
    Add2 &H410, 14, &H0, &H7, 0, 0, 0
    Add2 &H411, 11, &H1, &H7, 0, 0, 0
    Add2 &H412, 10, &H2, &H7, 0, 0, 0
    Add2 &H413, 11, &H3, &H7, 0, 0, 0
    Add2 &H414, 11, &H4, &H7, 0, 0, 0
    Add2 &H415, 10, &H5, &H7, 0, 0, 0
    Add2 &H416, 15, &H6, &H7, 0, 0, 0
    Add2 &H417, 8, &H7, &H7, 0, 0, 0
    Add2 &H418, 11, &H8, &H7, 0, 0, 0
    Add2 &H419, 11, &H9, &H7, 0, 0, 0
    Add2 &H41A, 11, &HA, &H7, 0, 0, 0
    Add2 &H41B, 9, &HB, &H7, 0, 0, 0
    Add2 &H41C, 14, &HC, &H7, 0, 0, 0
    Add2 &H41D, 11, &HD, &H7, 0, 0, 0
    Add2 &H41E, 12, &HE, &H7, 0, 0, 0
    Add2 &H41F, 11, &HF, &H7, 0, 0, 0
    Add2 &H420, 10, &H10, &H7, 0, 0, 0
    Add2 &H421, 12, &H11, &H7, 0, 0, 0
    Add2 &H422, 15, &H12, &H7, 0, 0, 0
    Add2 &H423, 10, &H13, &H7, 0, 0, 0
    Add2 &H424, 11, &H14, &H7, 0, 0, 0
    Add2 &H425, 11, &H15, &H7, 0, 0, 0
    Add2 &H426, 13, &H16, &H7, 0, 0, 0
    Add2 &H427, 11, &H17, &H7, 0, 0, 0
    Add2 &H428, 15, &H18, &H7, 0, 0, 0
    Add2 &H429, 18, &H19, &H7, 0, 0, 0
    Add2 &H42A, 12, &H1A, &H7, 0, 0, 0
    Add2 &H42B, 15, &H1B, &H7, 0, 0, 0
    Add2 &H42C, 10, &H1C, &H7, 0, 0, 0
    Add2 &H42D, 8, &H1D, &H7, 0, 0, 0
    Add2 &H42E, 13, &H1E, &H7, 0, 0, 0
    Add2 &H42F, 10, &H1F, &H7, 0, 0, 0
    Add2 &H430, 9, &H0, &H8, 0, 0, 0
    Add2 &H431, 7, &H1, &H8, 0, 0, 0
    Add2 &H432, 7, &H2, &H8, 0, 0, 0
    Add2 &H433, 7, &H3, &H8, 0, 0, 0
    Add2 &H434, 8, &H4, &H8, 0, 0, 0
    Add2 &H435, 8, &H5, &H8, 0, 0, 0
    Add2 &H436, 11, &H6, &H8, 0, 0, 0
    Add2 &H437, 6, &H7, &H8, 0, 0, 0
    Add2 &H438, 9, &H8, &H8, 0, 0, 0
    Add2 &H439, 9, &H9, &H8, 0, 0, 0
    Add2 &H43A, 8, &HA, &H8, 0, 0, 0
    Add2 &H43B, 8, &HB, &H8, 0, 0, 0
    Add2 &H43C, 11, &HC, &H8, 0, 0, 0
    Add2 &H43D, 8, &HD, &H8, 0, 0, 0
    Add2 &H43E, 8, &HE, &H8, 0, 0, 0
    Add2 &H43F, 8, &HF, &H8, 0, 0, 0
    Add2 &H440, 9, &H10, &H8, 0, 0, 0
    Add2 &H441, 8, &H11, &H8, 0, 0, 0
    Add2 &H442, 12, &H12, &H8, 0, 0, 0
    Add2 &H443, 8, &H13, &H8, 0, 0, 0
    Add2 &H444, 10, &H14, &H8, 0, 0, 0
    Add2 &H445, 9, &H15, &H8, 0, 0, 0
    Add2 &H446, 11, &H16, &H8, 0, 0, 0
    Add2 &H447, 9, &H17, &H8, 0, 0, 0
    Add2 &H448, 12, &H18, &H8, 0, 0, 0
    Add2 &H449, 14, &H19, &H8, 0, 0, 0
    Add2 &H44A, 9, &H1A, &H8, 0, 0, 0
    Add2 &H44B, 11, &H1B, &H8, 0, 0, 0
    Add2 &H44C, 7, &H1C, &H8, 0, 0, 0
    Add2 &H44D, 7, &H1D, &H8, 0, 0, 0
    Add2 &H44E, 9, &H1E, &H8, 0, 0, 0
    Add2 &H44F, 8, &H1F, &H8, 0, 0, 0
    Add2 &H451, 8, &H20, &H8, 0, 0, 0

    ReDim Preserve gGlyphs2(0 To m_AddIndex - 1)
    gGlyphs2Count = m_AddIndex
    m_Ready = True
End Sub

' ------------------------------------------------------------------
' 标志：1=有ExtraUp，2=有ExtraDown
Private Sub Add2(ByVal Code As Long, ByVal W As Long, ByVal X As Long, _
                 ByVal Y As Long, ByVal ExtraUp As Long, ByVal ExtraDown As Long, _
                 ByVal Flags As Long)
    If m_AddIndex > UBound(gGlyphs2) Then ReDim Preserve gGlyphs2(0 To m_AddIndex + 127)
    With gGlyphs2(m_AddIndex)
        .Code = Code
        .Width = W
        .X = X
        .Y = Y
        .ExtraUp = ExtraUp
        .ExtraDown = ExtraDown
        .HasExtraUp = ((Flags And 1) <> 0)
        .HasExtraDown = ((Flags And 2) <> 0)
        .Drawable = (X >= 0 And W > 0)
    End With
    m_AddIndex = m_AddIndex + 1
End Sub

' ------------------------------------------------------------------
'将索引返回到gGlyphs2（）中，当字符未知时返回-1。
Public Function Glyph2_Find(ByVal Code As Long) As Long
    Dim i As Long
    For i = 0 To gGlyphs2Count - 1
        If gGlyphs2(i).Code = Code Then
            Glyph2_Find = i
            Exit Function
        End If
    Next i
    Glyph2_Find = -1
End Function
