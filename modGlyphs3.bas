Attribute VB_Name = "modGlyphs3"
' =====================================================================
'modGlyphs3.bas-HoMM3位图字体度量
'《魔法门之英雄无敌》风格的消息生成器 By 高保洪 qq127644712
'从js/Homm3MessageGenerator.js（Homm3MessageGenerater.letters）移植而来
'每个字形都由Unicode码位存储，因此源文件保持不变
'“纯ASCII，适用于ANSI和Unicode版本。
'
'Sprite sheet:img/homm3-prite.png（904 x 333）
' =====================================================================
Option Explicit

Public Type GLYPH3_INFO
    Code        As Long      ' Unicode代码点
    Width       As Long      ' 前进/字形宽度（像素）
    X           As Long      ' sprite工作表内的列
    Y           As Long      ' 子画面内的行
    MarginLeft   As Long
    MarginRight  As Long
    TranslateY   As Long
    Height       As Long
    HasMarginLeft As Boolean
    HasMarginRight As Boolean
    HasTranslateY As Boolean
    HasHeight     As Boolean
    Drawable    As Boolean   ' 有一个位图要blit
End Type

Public gGlyphs3()    As GLYPH3_INFO
Public gGlyphs3Count As Long

Private m_AddIndex  As Long
Private m_Ready     As Boolean

' ------------------------------------------------------------------
Public Sub Glyphs3_Init()
    If m_Ready Then Exit Sub
    m_AddIndex = 0
    ReDim gGlyphs3(0 To 511)
    Add3 &H20, 3, -1, -1, 0, 0, 0, 0, 0
    Add3 &H21, 2, &H83, &H109, 1, 1, 0, 10, 11
    Add3 &H22, 4, &HC3, &H109, 0, 0, -6, 5, 12
    Add3 &H23, 8, &H8C, &H109, 0, 0, 0, 10, 8
    Add3 &H24, 6, &H73, &H107, 0, 0, 0, 12, 8
    Add3 &H25, 10, &HA3, &H109, 0, 0, 0, 10, 8
    Add3 &H26, 10, &H46, &H109, 0, 0, 0, 10, 8
    Add3 &H27, 2, &HC7, &H109, 0, 0, -6, 5, 12
    Add3 &H28, 4, &H9B, &H108, 0, 0, 3, 13, 12
    Add3 &H29, 4, &H9F, &H108, 1, -1, 3, 13, 15
    Add3 &H2A, 6, &H50, &H109, 0, 0, -5, 5, 12
    Add3 &H2B, 8, &HAD, &H10A, 0, 0, -1, 8, 12
    Add3 &H2C, 4, &H6A, &H110, 0, -1, 2, 5, 14
    Add3 &H2D, 5, &H6E, &H10D, 0, -1, -3, 2, 14
    Add3 &H2E, 2, &H79, &H111, 1, 0, 0, 2, 9
    Add3 &H2F, 6, &HD9, &H109, 0, 0, 2, 13, 12
    Add3 &H30, 7, &H0, &H109, 0, 1, 0, 10, 10
    Add3 &H31, 7, &H7, &H109, -1, 0, 0, 10, 9
    Add3 &H32, 7, &HE, &H109, 0, 0, 0, 10, 8
    Add3 &H33, 7, &H15, &H109, 0, 0, 0, 10, 8
    Add3 &H34, 7, &H1C, &H109, 0, 0, 0, 10, 8
    Add3 &H35, 7, &H23, &H109, 0, 0, 0, 10, 8
    Add3 &H36, 7, &H2A, &H109, 0, 1, 0, 10, 10
    Add3 &H37, 7, &H31, &H109, 0, -1, 0, 10, 10
    Add3 &H38, 7, &H38, &H109, 0, 0, 0, 10, 8
    Add3 &H39, 7, &H3F, &H109, 0, 0, 0, 10, 8
    Add3 &H3A, 2, &H68, &H10C, 0, 0, 0, 7, 8
    Add3 &H3B, 4, &HD5, &H10C, 0, 0, 2, 9, 12
    Add3 &H3C, 7, &H85, &H10B, 0, 0, 0, 8, 8
    Add3 &H3D, 8, &H7B, &H10C, 0, 0, -5, 4, 12
    Add3 &H3E, 7, &H94, &H10B, 0, 0, 0, 8, 8
    Add3 &H3F, 5, &HBB, &H109, 0, 0, 0, 10, 8
    Add3 &H40, 10, &H56, &H109, 0, 0, 0, 10, 8
    Add3 &H41, 10, &HB5, &HF7, 0, 0, 0, 10, 8
    Add3 &H42, 8, &HBF, &HF7, 0, 0, 0, 10, 8
    Add3 &H43, 9, &HC7, &HF7, 0, 0, 0, 10, 8
    Add3 &H44, 10, &HD0, &HF7, 0, 0, 0, 10, 8
    Add3 &H45, 8, &HDA, &HF7, 0, 0, 0, 10, 8
    Add3 &H46, 7, &HE2, &HF7, 0, 0, 0, 10, 8
    Add3 &H47, 10, &HE9, &HF7, 0, 0, 0, 10, 8
    Add3 &H48, 11, &HF3, &HF7, 0, 0, 0, 10, 8
    Add3 &H49, 4, &HFE, &HF7, 0, 0, 0, 10, 8
    Add3 &H4A, 5, &H102, &HF7, -1, 0, 3, 13, 13
    Add3 &H4B, 10, &H107, &HF7, 0, -1, 0, 10, 10
    Add3 &H4C, 8, &H111, &HF7, 0, 0, 0, 10, 8
    Add3 &H4D, 12, &H119, &HF7, 0, 0, 0, 10, 8
    Add3 &H4E, 11, &H125, &HF7, 0, 0, 0, 10, 8
    Add3 &H4F, 10, &H130, &HF7, 0, 0, 0, 10, 8
    Add3 &H50, 8, &H13A, &HF7, 0, 0, 0, 10, 8
    Add3 &H51, 10, &H142, &HF7, 0, 0, 2, 12, 12
    Add3 &H52, 9, &H14C, &HF7, 0, 0, 0, 10, 8
    Add3 &H53, 7, &H155, &HF7, 0, 0, 0, 10, 8
    Add3 &H54, 8, &H15C, &HF7, 0, 0, 0, 10, 8
    Add3 &H55, 10, &H164, &HF7, 0, 0, 0, 10, 8
    Add3 &H56, 10, &H16E, &HF7, 0, -1, 0, 10, 10
    Add3 &H57, 14, &H178, &HF7, 0, -1, 0, 10, 10
    Add3 &H58, 9, &H186, &HF7, 0, 0, 0, 10, 8
    Add3 &H59, 9, &H18F, &HF7, 0, 0, 0, 10, 8
    Add3 &H5A, 9, &H198, &HF7, 0, 0, 0, 10, 8
    Add3 &H5B, 4, &H60, &H106, 1, 0, 3, 13, 13
    Add3 &H5C, 6, &H108, &H109, 0, 0, 2, 13, 12
    Add3 &H5D, 4, &H64, &H106, 0, 0, 3, 13, 12
    Add3 &H5E, 6, &HB5, &H108, 0, 0, -5, 6, 12
    Add3 &H5F, 8, &HED, &H111, 0, 0, 0, 2, 8
    Add3 &H60, 3, &HC0, &H109, 0, 0, -7, 3, 12
    Add3 &H61, 7, &H0, &HFA, 0, 0, 0, 7, 8
    Add3 &H62, 7, &H7, &HF7, 0, 0, 0, 10, 8
    Add3 &H63, 6, &HE, &HFA, 0, 0, 0, 7, 8
    Add3 &H64, 7, &H14, &HF7, 1, 0, 0, 10, 9
    Add3 &H65, 6, &H1B, &HFA, 0, 0, 0, 7, 8
    Add3 &H66, 5, &H21, &HF7, 0, -1, 0, 10, 10
    Add3 &H67, 7, &H26, &HFA, 0, 0, 4, 11, 12
    Add3 &H68, 8, &H2D, &HF7, 0, 0, 0, 10, 8
    Add3 &H69, 4, &H35, &HF8, 0, 0, 0, 9, 8
    Add3 &H6A, 4, &H39, &HF8, -1, 0, 4, 13, 13
    Add3 &H6B, 8, &H3D, &HF7, 0, -1, 0, 10, 10
    Add3 &H6C, 4, &H45, &HF7, 0, 0, 0, 10, 8
    Add3 &H6D, 12, &H49, &HFA, 0, -1, 0, 7, 10
    Add3 &H6E, 8, &H55, &HFA, 0, 0, 0, 7, 8
    Add3 &H6F, 7, &H5D, &HFA, 0, 0, 0, 7, 8
    Add3 &H70, 8, &H64, &HFA, 0, 0, 4, 11, 12
    Add3 &H71, 8, &H6C, &HFA, 0, -1, 4, 11, 14
    Add3 &H72, 6, &H74, &HFA, 0, -1, 0, 7, 10
    Add3 &H73, 5, &H7A, &HFA, 0, 1, 0, 7, 10
    Add3 &H74, 5, &H7F, &HF8, -1, 0, 0, 9, 9
    Add3 &H75, 8, &H84, &HFA, 0, 0, 0, 7, 8
    Add3 &H76, 8, &H8C, &HFA, 0, -1, 0, 7, 10
    Add3 &H77, 11, &H94, &HFA, 0, 0, 0, 7, 8
    Add3 &H78, 7, &H9F, &HFA, 0, 0, 0, 7, 8
    Add3 &H79, 8, &HA6, &HFA, 0, -1, 4, 11, 14
    Add3 &H7A, 7, &HAE, &HFA, 0, 0, 0, 7, 8
    Add3 &H7B, 3, &H106, &H108, 0, 0, 0, 11, 8
    Add3 &H7C, 2, &HF5, &H108, 0, 0, 2, 13, 12
    Add3 &H7D, 3, &H109, &H108, 0, 0, 0, 11, 8
    Add3 &H7E, 8, &HE5, &H10B, 0, 0, -5, 3, 12
    Add3 &HA4, 8, &H100, &H131, 0, 0, -1, 8, 12
    Add3 &HA6, 2, &HF5, &H108, 0, 0, 2, 13, 12
    Add3 &HA7, 7, &H10E, &H130, 0, 0, 2, 12, 12
    Add3 &HA9, 11, &H115, &H130, 0, 0, 0, 10, 8
    Add3 &HAB, 6, &H129, &H135, 0, 0, -1, 4, 12
    Add3 &HAC, 8, &H12F, &H133, 0, 0, -3, 4, 12
    Add3 &HAE, 11, &H137, &H130, 0, 0, 0, 10, 8
    Add3 &HB0, 5, &H147, &H12F, 0, 0, -5, 5, 12
    Add3 &HB1, 8, &H14C, &H12F, 0, 0, 0, 10, 8
    Add3 &HB5, 6, &H159, &H132, 0, 0, 2, 9, 12
    Add3 &HB6, 7, &H15F, &H12F, 0, 0, 3, 13, 12
    Add3 &HB7, 2, &H166, &H135, 0, 0, -2, 2, 12
    Add3 &HBB, 6, &H17D, &H134, 0, 0, -1, 4, 12
    Add3 &HC1, 10, &HC7, &H13C, 0, 0, 0, 13, 8
    Add3 &HC9, 8, &HD1, &H13C, 0, 0, 0, 13, 8
    Add3 &HCD, 4, &HD9, &H13C, 0, 0, 0, 13, 8
    Add3 &HD3, 10, &H6D, &H13C, 0, 0, 0, 13, 8
    Add3 &HD6, 10, &HDD, &H13D, 0, 0, 0, 12, 8
    Add3 &HDA, 10, &HFB, &H13C, 0, 0, 0, 13, 8
    Add3 &HDC, 10, &HF1, &H13D, 0, 0, 0, 12, 8
    Add3 &HE1, 7, &H90, &H13F, 0, 0, 0, 10, 8
    Add3 &HE9, 6, &H97, &H13F, 0, 0, 0, 10, 8
    Add3 &HED, 4, &H9D, &H13F, 0, 0, 0, 10, 8
    Add3 &HF3, 7, &H22, &H13F, 0, 0, 0, 10, 8
    Add3 &HF6, 7, &HA1, &H140, 0, 0, 0, 9, 8
    Add3 &HFA, 8, &HB7, &H13F, 0, 0, 0, 10, 8
    Add3 &HFC, 8, &HAF, &H140, 0, 0, 0, 9, 8
    Add3 &H104, 10, &H3C, &H13F, 0, 0, 4, 14, 12
    Add3 &H105, 8, &H0, &H142, 0, 0, 3, 10, 12
    Add3 &H106, 9, &H46, &H13C, 0, 0, 0, 13, 8
    Add3 &H107, 6, &H8, &H13F, 0, 0, 0, 10, 8
    Add3 &H118, 9, &H4F, &H13F, 0, 0, 4, 14, 12
    Add3 &H119, 6, &HE, &H142, 0, 0, 3, 10, 12
    Add3 &H141, 10, &H58, &H13F, 0, 0, 0, 10, 8
    Add3 &H142, 6, &H14, &H13F, 0, 0, 0, 10, 8
    Add3 &H143, 11, &H62, &H13C, 0, 0, 0, 13, 8
    Add3 &H144, 8, &H1A, &H13F, 0, 0, 0, 10, 8
    Add3 &H150, 10, &HE7, &H13C, 0, 0, 0, 13, 8
    Add3 &H151, 7, &HA8, &H13F, 0, 0, 0, 10, 8
    Add3 &H15A, 7, &H77, &H13C, 0, 0, 0, 13, 8
    Add3 &H15B, 5, &H29, &H13F, 0, 0, 0, 10, 8
    Add3 &H170, 10, &H105, &H13C, 0, 0, 0, 13, 8
    Add3 &H171, 8, &HBF, &H13F, 0, 0, 0, 10, 8
    Add3 &H179, 9, &H7E, &H13C, 0, 0, 0, 13, 8
    Add3 &H17A, 7, &H2E, &H13F, 0, 0, 0, 10, 8
    Add3 &H17B, 9, &H87, &H13C, 0, 0, 0, 13, 8
    Add3 &H17C, 7, &H35, &H13F, 0, 0, 0, 10, 8
    Add3 &H2DC, 8, &HA9, &H131, 0, 0, 0, 9, 8
    Add3 &H401, 8, &HF6, &H119, 0, 0, 0, 12, 8
    Add3 &H402, 10, &H0, &H130, 0, 0, 1, 11, 12
    Add3 &H403, 8, &HA, &H12D, 0, 0, 0, 13, 8
    Add3 &H404, 9, &H120, &H130, 0, 0, 0, 10, 8
    Add3 &H405, 7, &H155, &HF7, 0, 0, 0, 10, 8
    Add3 &H406, 4, &HFE, &HF7, 0, 0, 0, 10, 8
    Add3 &H407, 5, &H142, &H12E, 0, 0, 0, 12, 8
    Add3 &H408, 5, &H102, &HF7, -1, 0, 3, 13, 13
    Add3 &H409, 13, &H50, &H130, 0, 0, 0, 10, 8
    Add3 &H40A, 14, &H60, &H130, 0, 0, 0, 10, 8
    Add3 &H40B, 10, &H78, &H130, 0, 0, 0, 10, 8
    Add3 &H40C, 10, &H6E, &H12D, 0, 0, 0, 13, 8
    Add3 &H40F, 10, &H82, &H130, 0, 0, 2, 12, 12
    Add3 &H410, 10, &HB5, &HF7, 0, 0, 0, 10, 8
    Add3 &H411, 8, &HDC, &H11B, 0, 0, 0, 10, 8
    Add3 &H412, 8, &HBF, &HF7, 0, 0, 0, 10, 8
    Add3 &H413, 8, &HE4, &H11B, 0, -1, 0, 10, 10
    Add3 &H414, 10, &HEC, &H11B, 0, 0, 2, 12, 12
    Add3 &H415, 8, &HDA, &HF7, 0, 0, 0, 10, 8
    Add3 &H416, 14, &HFE, &H11B, 0, -1, 0, 10, 10
    Add3 &H417, 7, &H10C, &H11B, 0, 0, 0, 10, 8
    Add3 &H418, 11, &H113, &H11B, 0, 0, 0, 10, 8
    Add3 &H419, 11, &H11E, &H119, 0, 0, 0, 12, 8
    Add3 &H41A, 10, &H107, &HF7, 0, -1, 0, 10, 10
    Add3 &H41B, 10, &H129, &H11B, 0, 0, 0, 10, 8
    Add3 &H41C, 12, &H119, &HF7, 0, 0, 0, 10, 8
    Add3 &H41D, 11, &HF3, &HF7, 0, 0, 0, 10, 8
    Add3 &H41E, 10, &H130, &HF7, 0, 0, 0, 10, 8
    Add3 &H41F, 11, &H133, &H11B, 0, 0, 0, 10, 8
    Add3 &H420, 8, &H13A, &HF7, 0, 0, 0, 10, 8
    Add3 &H421, 9, &HC7, &HF7, 0, 0, 0, 10, 8
    Add3 &H422, 8, &H15C, &HF7, 0, 0, 0, 10, 8
    Add3 &H423, 10, &H13E, &H11B, 0, -1, 3, 13, 14
    Add3 &H424, 12, &H148, &H11B, 0, -1, 0, 10, 10
    Add3 &H425, 9, &H186, &HF7, 0, 0, 0, 10, 8
    Add3 &H426, 12, &H154, &H11B, 0, -1, 2, 12, 14
    Add3 &H427, 10, &H160, &H11B, 0, 0, 0, 10, 8
    Add3 &H428, 14, &H16A, &H11B, 0, 0, 0, 10, 8
    Add3 &H429, 15, &H178, &H11B, 0, -1, 2, 12, 14
    Add3 &H42A, 9, &H187, &H11B, 0, 0, 0, 10, 8
    Add3 &H42B, 11, &H190, &H11B, 0, 1, 0, 10, 10
    Add3 &H42C, 8, &H19B, &H11B, 1, 0, 0, 10, 9
    Add3 &H42D, 9, &H1A3, &H11B, 0, 0, 0, 10, 8
    Add3 &H42E, 15, &H1AC, &H11B, 0, 0, 0, 10, 8
    Add3 &H42F, 9, &H1BB, &H11B, 0, 0, 0, 10, 8
    Add3 &H430, 7, &H0, &HFA, 0, 0, 0, 7, 8
    Add3 &H431, 6, &H0, &H11B, 1, 0, 0, 10, 9
    Add3 &H432, 6, &H6, &H11E, 0, 0, 0, 7, 8
    Add3 &H433, 7, &HC, &H11E, 0, -1, 0, 7, 10
    Add3 &H434, 8, &H13, &H11E, 0, -1, 2, 9, 14
    Add3 &H435, 6, &H1B, &HFA, 0, 0, 0, 7, 8
    Add3 &H436, 10, &H21, &H11E, 0, 0, 0, 7, 8
    Add3 &H437, 5, &H2B, &H11E, 0, 0, 0, 7, 8
    Add3 &H438, 8, &H30, &H11E, 0, 0, 0, 7, 8
    Add3 &H439, 8, &H38, &H11C, 0, 0, 0, 9, 8
    Add3 &H43A, 8, &H40, &H11E, 0, -1, 0, 7, 10
    Add3 &H43B, 7, &H48, &H11E, 0, 0, 0, 7, 8
    Add3 &H43C, 10, &H4F, &H11E, 0, -1, 0, 7, 10
    Add3 &H43D, 8, &H59, &H11E, 0, 0, 0, 7, 8
    Add3 &H43E, 7, &H5D, &HFA, 0, 0, 0, 7, 8
    Add3 &H43F, 8, &H61, &H11E, 0, 0, 0, 7, 8
    Add3 &H440, 8, &H64, &HFA, 0, 0, 4, 11, 12
    Add3 &H441, 6, &HE, &HFA, 0, 0, 0, 7, 8
    Add3 &H442, 6, &H69, &H11E, 0, 0, 0, 7, 8
    Add3 &H443, 8, &HA6, &HFA, 0, -1, 4, 11, 14
    Add3 &H444, 11, &H79, &H11B, 0, 0, 4, 14, 12
    Add3 &H445, 7, &H9F, &HFA, 0, 0, 0, 7, 8
    Add3 &H446, 8, &H84, &H11E, 0, 0, 2, 9, 12
    Add3 &H447, 8, &H8C, &H11E, 0, -1, 0, 7, 10
    Add3 &H448, 12, &H94, &H11E, 0, -1, 0, 7, 10
    Add3 &H449, 12, &HA0, &H11E, -1, 0, 2, 9, 13
    Add3 &H44A, 7, &HAC, &H11E, 0, 0, 0, 7, 8
    Add3 &H44B, 11, &HB3, &H11E, 0, -1, 0, 7, 10
    Add3 &H44C, 6, &HBE, &H11E, 0, 0, 0, 7, 8
    Add3 &H44D, 6, &HC4, &H11E, 0, 0, 0, 7, 8
    Add3 &H44E, 10, &HCA, &H11E, 0, 0, 0, 7, 8
    Add3 &H44F, 8, &HD4, &H11E, 0, -1, 0, 7, 10
    Add3 &H451, 6, &H1B, &H11C, 0, 0, 0, 9, 8
    Add3 &H452, 7, &H8C, &H130, 0, 0, 1, 11, 12
    Add3 &H453, 7, &H16, &H130, 0, 0, 0, 10, 8
    Add3 &H454, 6, &H177, &H132, 0, 0, 0, 7, 8
    Add3 &H455, 5, &H7A, &HFA, 0, 1, 0, 7, 10
    Add3 &H456, 4, &H35, &HF8, 0, 0, 0, 9, 8
    Add3 &H457, 5, &H183, &H130, 0, 0, 0, 9, 8
    Add3 &H458, 4, &H39, &HF8, -1, 0, 4, 13, 13
    Add3 &H459, 10, &HBE, &H133, 0, 0, 0, 7, 8
    Add3 &H45A, 11, &HCB, &H133, 0, 0, 0, 7, 8
    Add3 &H45B, 8, &HDE, &H130, 0, 0, 0, 10, 8
    Add3 &H45C, 8, &HD6, &H130, 0, 0, 0, 10, 8
    Add3 &H45F, 8, &HE6, &H133, 0, 0, 2, 9, 12
    Add3 &H490, 6, &H108, &H12F, 0, 0, 0, 11, 8
    Add3 &H491, 5, &H154, &H131, 0, 0, 0, 8, 8
    Add3 &H4EE, 10, &HEE, &H12E, 0, 0, 3, 15, 12
    Add3 &H4EF, 7, &HF8, &H131, 0, 0, 4, 13, 12
    Add3 &H2013, 7, &HF7, &H10D, 1, -1, -3, 2, 15
    Add3 &H2014, 10, &H9F, &H135, 0, 0, -3, 2, 12
    Add3 &H2018, 3, &H1D, &H137, 0, 0, -6, 4, 12
    Add3 &H2019, 3, &H1D, &H137, 0, 0, -6, 4, 12
    Add3 &H201A, 3, &H1D, &H137, 0, 0, 1, 4, 12
    Add3 &H201C, 6, &H1D, &H137, 0, 0, -6, 4, 12
    Add3 &H201D, 6, &H1D, &H137, 0, 0, -6, 4, 12
    Add3 &H201E, 6, &H1D, &H137, 0, 0, 1, 4, 12
    Add3 &H2020, 6, &H2D, &H130, 0, 0, 2, 12, 12
    Add3 &H2021, 6, &H33, &H130, 0, 0, 2, 12, 12
    Add3 &H2022, 5, &H93, &H132, 0, 0, -3, 5, 12
    Add3 &H2026, 10, &H23, &H138, 0, 0, 0, 2, 8
    Add3 &H2030, 15, &H41, &H130, 0, 0, 0, 10, 8
    Add3 &H2039, 3, &H129, &H135, 0, 0, 0, 4, 8
    Add3 &H203A, 3, &H17D, &H134, 0, 0, -1, 4, 12
    Add3 &H20AC, 8, &H39, &H130, 0, 0, 0, 10, 8
    Add3 &H2116, 15, &H168, &H12F, 0, 0, 0, 10, 8
    Add3 &H2122, 13, &HB1, &H130, 0, 0, -4, 6, 12

    ReDim Preserve gGlyphs3(0 To m_AddIndex - 1)
    gGlyphs3Count = m_AddIndex
    m_Ready = True
End Sub

' ------------------------------------------------------------------
'标志：1=具有MarginLeft，2=具有Margin Right，4=具有TranslateY，8=具有Height
Private Sub Add3(ByVal Code As Long, ByVal W As Long, ByVal X As Long, _
                 ByVal Y As Long, ByVal MarginLeft As Long, ByVal MarginRight As Long, ByVal TranslateY As Long, ByVal Height As Long, _
                 ByVal Flags As Long)
    If m_AddIndex > UBound(gGlyphs3) Then ReDim Preserve gGlyphs3(0 To m_AddIndex + 127)
    With gGlyphs3(m_AddIndex)
        .Code = Code
        .Width = W
        .X = X
        .Y = Y
        .MarginLeft = MarginLeft
        .MarginRight = MarginRight
        .TranslateY = TranslateY
        .Height = Height
        .HasMarginLeft = ((Flags And 1) <> 0)
        .HasMarginRight = ((Flags And 2) <> 0)
        .HasTranslateY = ((Flags And 4) <> 0)
        .HasHeight = ((Flags And 8) <> 0)
        .Drawable = (X >= 0 And W > 0)
    End With
    m_AddIndex = m_AddIndex + 1
End Sub

' ------------------------------------------------------------------
' 将索引返回到gGlyphs3（）中，当字符未知时返回-1。
Public Function Glyph3_Find(ByVal Code As Long) As Long
    Dim i As Long
    For i = 0 To gGlyphs3Count - 1
        If gGlyphs3(i).Code = Code Then
            Glyph3_Find = i
            Exit Function
        End If
    Next i
    Glyph3_Find = -1
End Function
