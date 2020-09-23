VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_DrawDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//--------------------------------------------------------------------------------
'
'//--  You are free to use this source as long as this copyright message
'//--    appears on your program's "About" dialog:
'
'//--  Outlook Bar Project
'//--  Copyright (c) 2002 Vlad Vissoultchev (wqw@myrealbox.com)
'
'//--------------------------------------------------------------------------------

Option Explicit
'//--------------------------------------------------------------------------------
'//--Public enums
'//--------------------------------------------------------------------------------
Public Enum UcsBackStyles
    BS_TRANSPARENT = 1                                '//--Draws transparent background.
    BS_OPAQUE = 2                                     '//--Draws opaque background.
    BS_NEWTRANSPARENT = 3
    '//--NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private BS_TRANSPARENT, BS_OPAQUE, BS_NEWTRANSPARENT
#End If
Public Enum UcsDrawTextStyles
    DT_LEFT = &H0                                     '//--Aligns text to the left.
    DT_TOP = &H0                                      '//--Justifies the text to the top of the rectangle.
    DT_CENTER = &H1                                   '//--Centers text horizontally in the rectangle.
    DT_RIGHT = &H2                                    '//--Aligns text to the right.
    DT_VCENTER = &H4
    '//--Centers text vertically. This value is used only with the DT_SINGLELINE value.
    DT_BOTTOM = &H8
    '//--Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value.
    DT_WORDBREAK = &H10
    '//--Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line.<br>If this is not specified, output is on one line.
    DT_SINGLELINE = &H20
    '//--Displays text on a single line only. Carriage returns and line feeds do not break the line.
    DT_EXPANDTABS = &H40
    '//--Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
    DT_TABSTOP = &H80
    '//--Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
    DT_NOCLIP = &H100
    '//--Draws without clipping. DrawText is somewhat faster when DT_NOCLIP is used.
    DT_EXTERNALLEADING = &H200
    '//--Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
    DT_CALCRECT = &H400
    '//--Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text.
    DT_NOPREFIX = &H800
    '//--Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off
    DT_INTERNAL = &H1000                              '//--Uses the system font to calculate text metrics.
    DT_EDITCONTROL = &H2000
    '//--Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
    DT_PATH_ELLIPSIS = &H4000
    '//--For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_END_ELLIPSIS = &H8000
    '//--For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_MODIFYSTRING = &H10000
    '//--Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
    DT_RTLREADING = &H20000
    '//--Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
    DT_WORD_ELLIPSIS = &H40000
    '//--Truncates any word that does not fit in the rectangle and adds ellipses.<br>Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private DT_LEFT, DT_TOP, DT_CENTER, DT_RIGHT, DT_VCENTER, DT_BOTTOM, DT_WORDBREAK, DT_SINGLELINE, DT_EXPANDTABS, DT_TABSTOP, DT_NOCLIP
    Private DT_EXTERNALLEADING, DT_CALCRECT, DT_NOPREFIX, DT_INTERNAL, DT_EDITCONTROL, DT_PATH_ELLIPSIS, DT_END_ELLIPSIS, DT_MODIFYSTRING
    Private DT_RTLREADING, DT_WORD_ELLIPSIS
#End If
Public Enum UcsBorderStyles
    BDR_RAISEDOUTER = &H1                             '//--Raised outer edge.
    BDR_SUNKENOUTER = &H2                             '//--Sunken outer edge.
    BDR_RAISEDINNER = &H4                             '//--Raised inner edge.
    BDR_SUNKENINNER = &H8                             '//--Sunken inner edge.
    BDR_OUTER = &H3                                   '//--(BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = &HC                                   '//--(BDR_RAISEDINNER Or BDR_SUNKENINNER)
    EDGE_RAISED = &H5                                 '//--(BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = &HA                                 '//--(BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = &H6                                 '//--(BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = &H9                                   '//--(BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private BDR_RAISEDOUTER, BDR_SUNKENOUTER, BDR_RAISEDINNER, BDR_SUNKENINNER, BDR_OUTER, BDR_INNER, EDGE_RAISED, EDGE_SUNKEN
    Private EDGE_ETCHED, EDGE_BUMP
#End If
Public Enum UcsBorderFlags
    BF_LEFT = &H1                                     '//--Left side of border rectangle.
    BF_TOP = &H2                                      '//--Top of border rectangle.
    BF_RIGHT = &H4                                    '//--Right side of border rectangle.
    BF_BOTTOM = &H8                                   '//--Bottom of border rectangle.
    BF_TOPLEFT = &H3                                  '//--(BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = &H6                                 '//--(BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = &H9                               '//--(BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = &HC                              '//--(BF_BOTTOM Or BF_RIGHT)
    BF_RECT = &HF                                     '//--(BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10                                '//--Diagonal border.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private BF_LEFT, BF_TOP, BF_RIGHT, BF_BOTTOM, BF_TOPLEFT, BF_TOPRIGHT, BF_BOTTOMLEFT, BF_BOTTOMRIGHT, BF_RECT, BF_DIAGONAL
#End If
Public Enum UcsExtTextOutStyles
    ETO_GRAYED = 1
    ETO_OPAQUE = 2
    '//--The current background color should be used to fill the rectangle.
    ETO_CLIPPED = 4                                   '//--The text will be clipped to the rectangle.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private ETO_GRAYED, ETO_OPAQUE, ETO_CLIPPED
#End If
Public Enum UcsDrawFrameControlType
    DFC_CAPTION = 1                                   '//--Title bar.
    DFC_MENU = 2                                      '//--Menu bar.
    DFC_SCROLL = 3                                    '//--Scroll bar.
    DFC_BUTTON = 4                                    '//--Standard button.
    DFC_POPUPMENU = 5                                 '//--<b>Windows 98/Me, Windows 2000 or later:</b> Popup menu item.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private DFC_CAPTION, DFC_MENU, DFC_SCROLL, DFC_BUTTON, DFC_POPUPMENU
#End If
Public Enum UcsDrawFrameControlStyle
    DFCS_BUTTONCHECK = &H0                            '//--Check box.
    DFCS_BUTTONRADIOIMAGE = &H1                       '//--Image for radio button (nonsquare needs image).
    DFCS_BUTTONRADIOMASK = &H2                        '//--Mask for radio button (nonsquare needs mask).
    DFCS_BUTTONRADIO = &H4                            '//--Radio button.
    DFCS_BUTTON3STATE = &H8                           '//--Three-state button.
    DFCS_BUTTONPUSH = &H10                            '//--Push button.
    DFCS_CAPTIONCLOSE = &H0                           '//--<b>Close</b> button.
    DFCS_CAPTIONMIN = &H1                             '//--<b>Minimize</b> button.
    DFCS_CAPTIONMAX = &H2                             '//--<b>Maximize</b> button.
    DFCS_CAPTIONRESTORE = &H3                         '//--<b>Restore</b> button.
    DFCS_CAPTIONHELP = &H4                            '//--<b>Help</b> button.
    DFCS_MENUARROW = &H0                              '//--Submenu arrow.
    DFCS_MENUCHECK = &H1                              '//--Check mark.
    DFCS_MENUBULLET = &H2                             '//--Bullet.
    DFCS_MENUARROWRIGHT = &H4
    '//--Submenu arrow pointing left. This is used for the right-to-left cascading menus used with right-to-left languages such as Arabic or Hebrew.
    DFCS_SCROLLUP = &H0                               '//--Up arrow of scroll bar.
    DFCS_SCROLLDOWN = &H1                             '//--Down arrow of scroll bar.
    DFCS_SCROLLLEFT = &H2                             '//--Left arrow of scroll bar.
    DFCS_SCROLLRIGHT = &H3                            '//--Right arrow of scroll bar.
    DFCS_SCROLLCOMBOBOX = &H5                         '//--Combo box scroll bar.
    DFCS_SCROLLSIZEGRIP = &H8                         '//--Size grip in bottom-right corner of window.
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
    '//--Size grip in bottom-left corner of window. This is used with right-to-left languages such as Arabic or Hebrew.
    DFCS_INACTIVE = &H100                             '//--Button is inactive (grayed).
    DFCS_PUSHED = &H200                               '//--Button is pushed.
    DFCS_CHECKED = &H400                              '//--Button is checked.
    DFCS_TRANSPARENT = &H800
    '//--<b>Windows 98/Me, Windows 2000 or later:</b> The background remains untouched.
    DFCS_HOT = &H1000
    '//--<b>Windows 98/Me, Windows 2000 or later:</b> Button is hot-tracked.
    DFCS_ADJUSTRECT = &H2000
    '//--Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
    DFCS_FLAT = &H4000                                '//--Button has a flat border.
    DFCS_MONO = &H8000                                '//--Button has a monochrome border.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private DFCS_BUTTONCHECK, DFCS_BUTTONRADIOIMAGE, DFCS_BUTTONRADIOMASK, DFCS_BUTTONRADIO, DFCS_BUTTON3STATE, DFCS_BUTTONPUSH
    Private DFCS_CAPTIONCLOSE, DFCS_CAPTIONMIN, DFCS_CAPTIONMAX, DFCS_CAPTIONRESTORE, DFCS_CAPTIONHELP, DFCS_MENUARROW, DFCS_MENUCHECK
    Private DFCS_MENUBULLET, DFCS_MENUARROWRIGHT, DFCS_SCROLLUP, DFCS_SCROLLDOWN, DFCS_SCROLLLEFT, DFCS_SCROLLRIGHT, DFCS_SCROLLCOMBOBOX
    Private DFCS_SCROLLSIZEGRIP, DFCS_SCROLLSIZEGRIPRIGHT, DFCS_INACTIVE, DFCS_PUSHED, DFCS_CHECKED, DFCS_TRANSPARENT, DFCS_HOT
    Private DFCS_ADJUSTRECT, DFCS_FLAT, DFCS_MONO
#End If
Public Enum UcsBrushStyle
    BS_SOLID = 0                                      '//--Solid brush.
    BS_HOLLOW = 1                                     '//--Hollow brush.
    BS_NULL = 1                                       '//--Same as BS_HOLLOW.
    BS_HATCHED = 2                                    '//--Hatched brush.
    BS_PATTERN = 3                                    '//--Pattern brush defined by a memory bitmap.
    BS_INDEXED = 4
    BS_DIBPATTERN = 5
    '//--A pattern brush defined by a device-independent bitmap (DIB) specification.
    BS_DIBPATTERNPT = 6
    '//--A pattern brush defined by a device-independent bitmap (DIB) specification. If <b>lbStyle</b> is BS_DIBPATTERNPT, the <b>lbHatch</b> member contains a pointer to a packed DIB.
    BS_PATTERN8X8 = 7                                 '//--Same as BS_PATTERN.
    BS_DIBPATTERN8X8 = 8                              '//--Same as BS_DIBPATTERN.
    BS_MONOPATTERN = 9                                '//--The brush is a monochrome (black & white) bitmap.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private BS_SOLID, BS_HOLLOW, BS_NULL, BS_HATCHED, BS_PATTERN, BS_INDEXED, BS_DIBPATTERN, BS_DIBPATTERNPT, BS_PATTERN8X8, BS_DIBPATTERN8X8
    Private BS_MONOPATTERN
#End If
Public Enum UcsHatchStyles
    HS_HORIZONTAL = 0                                 '//--Horizontal hatch.
    HS_VERTICAL = 1                                   '//--Vertical hatch.
    HS_FDIAGONAL = 2                                  '//--A 45-degree downward, left-to-right hatch.
    HS_BDIAGONAL = 3                                  '//--A 45-degree upward, left-to-right hatch.
    HS_CROSS = 4                                      '//--Horizontal and vertical cross-hatch.
    HS_DIAGCROSS = 5                                  '//--A 45-degree crosshatch.
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private HS_HORIZONTAL, HS_VERTICAL, HS_FDIAGONAL, HS_BDIAGONAL, HS_CROSS, HS_DIAGCROSS
#End If
Public Enum UcsPenStyles
    PS_SOLID = 0                                      '//--The pen is solid.
    PS_DASH = 1                                       '//--The pen is dashed.
    PS_DOT = 2                                        '//--The pen is dotted.
    PS_DASHDOT = 3                                    '//--The pen has alternating dashes and dots.
    PS_DASHDOTDOT = 4                                 '//--The pen has dashes and double dots.
    PS_NULL = 5                                       '//--The pen is invisible.
    PS_INSIDEFRAME = 6
    '//--The pen is solid. When this pen is used in any GDI drawing function that takes a bounding rectangle, the dimensions of the figure are shrunk so that it fits entirely in the bounding rectangle, taking into account the width of the pen. This applies only to geometric pens.
    PS_USERSTYLE = 7
    '//--<b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8
    '//--<b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    PS_STYLE_MASK = &HF                               '//--Mask for previous PS_XXX values.
    PS_ENDCAP_ROUND = &H0                             '//--End caps are round.
    PS_ENDCAP_SQUARE = &H100                          '//--End caps are square.
    PS_ENDCAP_FLAT = &H200                            '//--End caps are flat.
    PS_ENDCAP_MASK = &HF00                            '//--Mask for previous PS_ENDCAP_XXX values.
    PS_JOIN_ROUND = &H0                               '//--Joins are beveled.
    PS_JOIN_BEVEL = &H1000
    '//--Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                            '//--Joins are round.
    PS_JOIN_MASK = &HF000                             '//--Mask for previous PS_JOIN_XXX values.
    PS_COSMETIC = &H0                                 '//--The pen is cosmetic.
    PS_GEOMETRIC = &H10000                            '//--The pen is geometric.
    PS_TYPE_MASK = &HF0000                            '//--Mask for previous PS_XXX (pen type).
End Enum
#If False Then                                        'Trick preserves Case of Enums when typing in IDE
    Private PS_SOLID, PS_DASH, PS_DOT, PS_DASHDOT, PS_DASHDOTDOT, PS_NULL, PS_INSIDEFRAME, PS_USERSTYLE, PS_ALTERNATE, PS_STYLE_MASK
    Private PS_ENDCAP_ROUND, PS_ENDCAP_SQUARE, PS_ENDCAP_FLAT, PS_ENDCAP_MASK, PS_JOIN_ROUND, PS_JOIN_BEVEL, PS_JOIN_MITER, PS_JOIN_MASK
    Private PS_COSMETIC, PS_GEOMETRIC, PS_TYPE_MASK
#End If
'//--------------------------------------------------------------------------------
'//--API
'//--------------------------------------------------------------------------------
'//--- GetTextMetrics constants
Private Const FW_NORMAL As Long = 400
Private Const LF_FACESIZE As Long = 32
'//--- GetDeviceCaps constants
Private Const RASTERCAPS As Long = 38
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const SIZEPALETTE As Long = 104
Private Const RC_PALETTE As Long = &H100
Private Const CAPS1 As Long = 94                      '//--other caps
Private Const C1_TRANSPARENT As Long = &H1            '//--new raster cap
'//--- GetStockObject constants
Private Const DKGRAY_BRUSH As Long = 3
'//--- DrawIconEx constants
Private Const DI_NORMAL As Long = &H3
'//--- DIB Section constants
Private Const BI_RGB As Long = 0
Private Const DIB_RGB_COLORS As Long = 0              '//-- color table in RGBs
Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_PAL_INDICES As Long = 2
'//--- Raster Operation Codes
Private Const DSna As Long = &H220326
'//--- gradient fill
Private Const GRADIENT_FILL_RECT_H As Long = 0
Private Const GRADIENT_FILL_RECT_V As Long = 1
'//--- for SystemParametersInfo
Private Const SPI_GETICONTITLELOGFONT As Long = 31
Private Const SPI_GETNONCLIENTMETRICS As Long = 41
'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type POINTAPI
    X          As Long
    Y          As Long
End Type
Private Type SIZEAPI
    cx         As Long
    cy         As Long
End Type
Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type
Private Type LOGFONT
    lfHeight   As Long
    lfWidth    As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight   As Long
    lfItalic   As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet  As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality  As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Private Type DRAWTEXTPARAMS
    cbSize     As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type TEXTMETRIC
    tmHeight   As Long
    tmAscent   As Long
    tmDescent  As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight   As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic   As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet  As Byte
End Type
Private Type PALETTEENTRY
    peRed      As Byte
    peGreen    As Byte
    peBlue     As Byte
    peFlags    As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
Private Type PictDesc
    Size       As Long
Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type
Private Type LOGBRUSH
    lbStyle    As Long
    lbColor    As Long
    lbHatch    As Long
End Type
Private Type ICONINFO
    fIcon      As Long
    xHotspot   As Long
    yHotspot   As Long
    hbmMask    As Long
    hbmColor   As Long
End Type
Private Type BITMAPINFOHEADER                         '40 bytes
    biSize     As Long
    biWidth    As Long
    biHeight   As Long
    biPlanes   As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed  As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue    As Byte
    rgbGreen   As Byte
    rgbRed     As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader  As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type
Private Type TRIVERTEX
    X          As Long
    Y          As Long
    Red        As Integer
    Green      As Integer
    Blue       As Integer
    Alpha      As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type
Private Type NONCLIENTMETRICS
    cbSize     As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
'//--------------------------------------------------------------------------------
'//--Member variables and constants
'//--------------------------------------------------------------------------------
Private Const MASK_COLOR As Long = &HFF00FF
Private m_MemoryDC As Long                            '//--DC handle of the created Device Context
Private m_MemoryWidth As Long                         '//--Width of the bitmap
Private m_MemoryHeight As Long                        '//--Height of the bitmap
Private m_MemoryBitmap As Long                        '//--Handle of the created bitmap
Private m_OrginalBitmap As Long                       '//--Used in Destroy
Private m_MemoryPal As Long                           '//--Handle of the created palette
Private m_OrginalPal As Long                          '//--Used in Destroy
Private m_MemoryFont As Long                          '//--Font handle in use (last set)
Private m_OrginalFont As Long                         '//--Previously selected font
Private m_ParentDC As Long                            '//--Used in Copy method as default value
Private m_MemoryBrush As Long
Private m_OriginalBrush As Long
Private m_MemoryPen As Long
Private m_OriginalPen As Long
Private m_hDotBrush As Long
Private m_hDotBitmap As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal HDC As Long, _
                                                 ByVal crColor As Long) As Long
Private Declare Function ApiDrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal HDC As Long, _
                                                                         ByVal lpsz As String, _
                                                                         ByVal n As Long, _
                                                                         lpRect As RECT, _
                                                                         ByVal un As Long, _
                                                                         lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function GetBkMode Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, _
                                                ByVal nBkMode As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, _
                                                    ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal HDC As Long, _
                                                                       ByVal nCount As Long, _
                                                                       ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal HDC As Long, _
                                                                             lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, _
                                                              ByVal wStartIndex As Long, _
                                                              ByVal wNumEntries As Long, _
                                                              lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, _
                                                    ByVal HPALETTE As Long, _
                                                    ByVal bForceBackground As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function APIDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function ApiFillRect Lib "user32" Alias "FillRect" (ByVal HDC As Long, _
                                                                    lpRect As RECT, _
                                                                    ByVal hBrush As Long) As Long
Private Declare Function ApiBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hdcDest As Long, _
                                                               ByVal X As Long, _
                                                               ByVal Y As Long, _
                                                               ByVal nWidth As Long, _
                                                               ByVal nHeight As Long, _
                                                               ByVal hdcSrc As Long, _
                                                               ByVal xSrc As Long, _
                                                               ByVal ySrc As Long, _
                                                               ByVal dwRop As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, _
                                                                      riid As Any, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      ipic As IPicture) As Long
Private Declare Function ApiStretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal HDC As Long, _
                                                                       ByVal X As Long, _
                                                                       ByVal Y As Long, _
                                                                       ByVal nWidth As Long, _
                                                                       ByVal nHeight As Long, _
                                                                       ByVal hdcSrc As Long, _
                                                                       ByVal xSrc As Long, _
                                                                       ByVal ySrc As Long, _
                                                                       ByVal nSrcWidth As Long, _
                                                                       ByVal nSrcHeight As Long, _
                                                                       ByVal dwRop As Long) As Long
Private Declare Function CreateEnhMetaFileLong Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, _
                                                                                       ByVal lpFileName As String, _
                                                                                       ByVal lpRect As Long, _
                                                                                       ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal HDC As Long, _
                                                   ByVal hMF As Long) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal HDC As Long, _
                                                      ByVal hEmf As Long, _
                                                      lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal HDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal nPlanes As Long, _
                                                   ByVal nBitCount As Long, _
                                                   lpBits As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal HDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal HDC As Long, _
                                                                                     ByVal lpszString As String, _
                                                                                     ByVal cbString As Long, _
                                                                                     lpSize As SIZEAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                                               ByVal lHPalette As Long, _
                                                               lColorRef As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, _
                                                       pBitmapInfo As BITMAPINFO, _
                                                       ByVal un As Long, _
                                                       ByVal lplpVoid As Long, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal HDC As Long, _
                                                       ByVal un1 As Long, _
                                                       ByVal un2 As Long, _
                                                       pcRGBQuad As RGBQUAD) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal HDC As Long, _
                                                 ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal HDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal HDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long
Private Declare Function ApiCreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, _
                                                                     ByVal nWidth As Long, _
                                                                     ByVal crColor As Long) As Long
Private Declare Function ApiDrawEdge Lib "user32" Alias "DrawEdge" (ByVal HDC As Long, _
                                                                    qrc As RECT, _
                                                                    ByVal edge As Long, _
                                                                    ByVal grfFlags As Long) As Long
Private Declare Function ApiExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal HDC As Long, _
                                                                        ByVal X As Long, _
                                                                        ByVal Y As Long, _
                                                                        ByVal wOptions As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal lpString As String, _
                                                                        ByVal nCount As Long, _
                                                                        ByVal lpDx As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal HDC As Long, _
                                                    ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long
Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal HDC As Long, _
                                                         pVertex As TRIVERTEX, _
                                                         ByVal dwNumVertex As Long, _
                                                         pMesh As GRADIENT_RECT, _
                                                         ByVal dwNumMesh As Long, _
                                                         ByVal dwMode As Long) As Long
Private Declare Function ApiFrameRect Lib "user32" Alias "FrameRect" (ByVal HDC As Long, _
                                                                      lpRect As RECT, _
                                                                      ByVal hBrush As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal HDC As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal dx As Long, _
                                                    ByVal dy As Long, _
                                                    ByVal srcX As Long, _
                                                    ByVal srcY As Long, _
                                                    ByVal wSrcWidth As Long, _
                                                    ByVal wSrcHeight As Long, _
                                                    lpBits As Any, _
                                                    lpBitsInfo As BITMAPINFO, _
                                                    ByVal wUsage As Long, _
                                                    ByVal dwRop As Long) As Long
Private Declare Function ApiArc Lib "gdi32" Alias "Arc" (ByVal HDC As Long, _
                                                         ByVal x1 As Long, _
                                                         ByVal y1 As Long, _
                                                         ByVal x2 As Long, _
                                                         ByVal y2 As Long, _
                                                         ByVal X3 As Long, _
                                                         ByVal Y3 As Long, _
                                                         ByVal X4 As Long, _
                                                         ByVal Y4 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal HDC As Long, _
                                              ByVal nDrawMode As Long) As Long
Private Declare Function GetROP2 Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function ApiEllipse Lib "gdi32" Alias "Ellipse" (ByVal HDC As Long, _
                                                                 ByVal x1 As Long, _
                                                                 ByVal y1 As Long, _
                                                                 ByVal x2 As Long, _
                                                                 ByVal y2 As Long) As Long
Private Declare Function ApiCreateBrushIndirect Lib "gdi32" Alias "CreateBrushIndirect" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function ApiGetPixel Lib "gdi32" Alias "GetPixel" (ByVal HDC As Long, _
                                                                   ByVal X As Long, _
                                                                   ByVal Y As Long) As Long
Private Declare Function ApiSetPixel Lib "gdi32" Alias "SetPixel" (ByVal HDC As Long, _
                                                                   ByVal X As Long, _
                                                                   ByVal Y As Long, _
                                                                   ByVal crColor As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Private Declare Function ApiDrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal HDC As Long, _
                                                                                    lpRect As RECT, _
                                                                                    ByVal un1 As Long, _
                                                                                    ByVal un2 As Long) As Long
Private Declare Function APIRectangle Lib "gdi32" Alias "Rectangle" (ByVal HDC As Long, _
                                                                     ByVal x1 As Long, _
                                                                     ByVal y1 As Long, _
                                                                     ByVal x2 As Long, _
                                                                     ByVal y2 As Long) As Long

' Type - GetObjectAPI.lpObject
Private Type BITMAP
    bmType     As Long                                'LONG   // Specifies the bitmap type. This member must be zero.
    bmWidth    As Long                                'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
    bmHeight   As Long                                'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
    bmWidthBytes As Long                              'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
    bmPlanes   As Integer                             'WORD   // Specifies the count of color planes.
    bmBitsPixel As Integer                            'WORD   // Specifies the number of bits required to indicate the color of a pixel.
    bmBits     As Long                                'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long

'Purpose: Returns or sets the background color for the device context of the <b>IAPP_DrawDC</b> object.

Public Property Get BackColor() As Long

    If IsCreated() Then
        BackColor = GetBkColor(m_MemoryDC)
    End If

End Property

Public Property Let BackColor(ByVal newValue As Long)

    If IsCreated() Then
        SetBkColor m_MemoryDC, TranslateColor(newValue)
    End If

End Property

'Purpose: Returns or sets the background style for the device context of the <b>IAPP_DrawDC</b> object.

Public Property Get BackStyle() As UcsBackStyles

    If IsCreated() Then
        BackStyle = GetBkMode(m_MemoryDC)
    End If

End Property

Public Property Let BackStyle(ByVal newValue As UcsBackStyles)

    If IsCreated() Then
        SetBkMode m_MemoryDC, newValue
    End If

End Property

'Purpose: Block-image-transfers bits from current device context into a specified destination device context using a specified raster operation.

Public Sub BitBlt(Optional ByVal hdcDest As Long, _
                  Optional ByVal xDest As Long, _
                  Optional ByVal yDest As Long, _
                  Optional ByVal nWidth As Long, _
                  Optional ByVal nHeight As Long, _
                  Optional ByVal xSrc As Long, _
                  Optional ByVal ySrc As Long, _
                  Optional ByVal dwRop As RasterOpConstants = vbSrcCopy)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- fix arguments
    If hdcDest = 0 Then
        hdcDest = m_ParentDC
    End If
    If nWidth = 0 Then
        nWidth = m_MemoryWidth
    End If
    If nHeight = 0 Then
        nHeight = m_MemoryHeight
    End If
    ApiBitBlt hdcDest, xDest, yDest, nWidth, nHeight, m_MemoryDC, xSrc, ySrc, dwRop
    On Error GoTo 0

End Sub

'Purpose: Creates a <b>StdPicture</b> object from HBITMAP handle.

Public Function BitmapToPicture(ByVal hBmp As Long, _
                                Optional ByVal hPal As Long = 0) As IPicture

'//--- Returns a VB picture object containing the specified bitmap.

    Dim oNewPic As Picture
    Dim lpPictDesc As PictDesc
    Dim aGuid(0 To 3) As Long

    '//--- fill struct
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hBmpOrIcon = hBmp
        .hPal = hPal
    End With                                          'LPPICTDESC
    '//--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '//--- Create picture from bitmap handle
    OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
    '//--- success
    Set BitmapToPicture = oNewPic

End Function

'Purpose: Returns or sets the currently selected GDI brush (HBRUSH) for the device context of the <b>IAPP_DrawDC</b> object.

Public Property Get Brush() As Long

    Brush = m_MemoryBrush

End Property

Public Property Let Brush(ByVal lValue As Long)

    Dim hPrevBrush As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    If lValue <> 0 Then
        '//--- set new brush (and save original)
        hPrevBrush = SelectObject(m_MemoryDC, lValue)
        If m_OriginalBrush = 0 Then
            m_OriginalBrush = hPrevBrush
        End If
    ElseIf m_OriginalBrush <> 0 Then                  'NOT LVALUE...
        '//--- restore original brush
        hPrevBrush = SelectObject(m_MemoryDC, m_OriginalBrush)
        m_OriginalBrush = 0
    End If
    '//--- delete if previously selected
    If m_MemoryBrush = hPrevBrush Then
        If hPrevBrush <> m_hDotBrush Then
            APIDeleteObject hPrevBrush
        End If
    End If
    '//--- success
    m_MemoryBrush = lValue
    On Error GoTo 0

End Property

'//--------------------------------------------------------------------------------
'//--Class events
'//--------------------------------------------------------------------------------

Private Sub Class_Initialize()

    m_MemoryDC = 0

End Sub

Private Sub Class_Terminate()

    Dim lErrNumber As Long
    Dim sErrSrc As String
    Dim sErrDesc As String

    '//--- preserve errorinfo as to try not to clear the err object
    '//---   in Destroy -- it's using On Error Resume Next. usually
    '//---   this kind of support is done with Push/PopError funcs

    With Err
        lErrNumber = .Number
        sErrSrc = .Source
        sErrDesc = .Description
    End With                                          'Err
    Destroy                                           '//--- installs On Error Resume Next
    With Err
        .Number = lErrNumber
        .Source = sErrSrc
        .Description = sErrDesc
    End With                                          'Err
    On Error GoTo 0

End Sub

'Purpose: Clears contents of the currently selected device context optionally filling it with a specified color or brush.

Public Sub Cls(Optional ByVal clrFill As Long = -1, _
               Optional ByVal hbrFill As Long)

    FillRect 0, 0, Width, Height, clrFill, hbrFill

End Sub

'Purpose: Cleans up allocated resources and prevents GDI leaks.

Public Sub Destroy()

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- deselect objects
    If m_OriginalPen <> 0 Then
        SelectObject m_MemoryDC, m_OriginalPen
        m_OriginalPen = 0
    End If
    If m_OriginalBrush <> 0 Then
        SelectObject m_MemoryDC, m_OriginalBrush
        m_OriginalBrush = 0
    End If
    If m_OrginalPal <> 0 Then
        SelectPalette m_MemoryDC, m_OrginalPal, 0&
        RealizePalette m_MemoryDC
        m_OrginalPal = 0
    End If
    If m_OrginalBitmap <> 0 Then
        SelectObject m_MemoryDC, m_OrginalBitmap
        m_OrginalBitmap = 0
    End If
    If m_OrginalFont <> 0 Then
        SelectObject m_MemoryDC, m_OrginalFont
        m_OrginalFont = 0
    End If
    '//--- delete objects
    If m_MemoryPen <> 0 Then
        APIDeleteObject m_MemoryPen
        m_MemoryPen = 0
    End If
    If m_MemoryBrush <> 0 Then
        APIDeleteObject m_MemoryBrush
        m_MemoryBrush = 0
    End If
    If m_hDotBrush <> 0 Then
        APIDeleteObject m_hDotBrush
        m_hDotBrush = 0
        If m_hDotBitmap <> 0 Then
            APIDeleteObject m_hDotBitmap
            m_hDotBitmap = 0
        End If
    End If
    If m_MemoryFont <> 0 Then
        APIDeleteObject m_MemoryFont
        m_MemoryFont = 0
    End If
    If m_MemoryPal <> 0 Then
        APIDeleteObject m_MemoryPal
        m_MemoryPal = 0
    End If
    If m_MemoryBitmap <> 0 Then
        APIDeleteObject m_MemoryBitmap
        m_MemoryBitmap = 0
        If m_MemoryDC <> 0 Then
            DeleteDC m_MemoryDC
            m_MemoryDC = 0
        End If
    End If
    On Error GoTo 0

End Sub

'Purpose: Returns a Long which represents a handle to a patterns brush (HBRUSH) with alternating black and white pixels.

Public Property Get DotBrush() As Long

    Dim waBits(0 To 3) As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    If m_hDotBrush = 0 Then
        '//--- fill pattern array
        waBits(0) = &H5555AAAA
        waBits(1) = &H5555AAAA
        waBits(2) = &H5555AAAA
        waBits(3) = &H5555AAAA
        '//--- create pattern brush
        m_hDotBitmap = CreateBitmap(8, 8, 1, 1, waBits(0))
        If m_hDotBitmap <> 0 Then
            m_hDotBrush = CreatePatternBrush(m_hDotBitmap)
            APIDeleteObject m_hDotBitmap
            m_hDotBitmap = 0
        End If
        '//--- default to darkgray brush
        If m_hDotBrush = 0 Then
            m_hDotBrush = GetStockObject(DKGRAY_BRUSH)
        End If
    End If
    '//--- success
    DotBrush = m_hDotBrush
    On Error GoTo 0

End Property

'Purpose: Draws an edge with specified borders around a rectangle in the device context.

Public Sub DrawEdge(Optional ByVal LeftX As Long, _
                    Optional ByVal TopY As Long, _
                    Optional ByVal RightX As Long = -1, _
                    Optional ByVal BottomY As Long = -1, _
                    Optional ByVal edge As UcsBorderStyles = BDR_SUNKENOUTER, _
                    Optional ByVal Flags As UcsBorderFlags = BF_RECT)

    Dim rc     As RECT

    With rc
        .Left = LeftX
        .Top = TopY
        .Right = IIf(RightX < LeftX, Width, RightX)
        .Bottom = IIf(BottomY < TopY, Height, BottomY)
    End With                                          'RC
    ApiDrawEdge m_MemoryDC, rc, edge, Flags

End Sub

'Purpose: Draws a straight line between two points in the device context using a specified color or pen.

Public Sub DrawLine(ByVal x1 As Long, _
                    ByVal y1 As Long, _
                    ByVal x2 As Long, _
                    ByVal y2 As Long, _
                    Optional ByVal clrLine As Long = -1, _
                    Optional ByVal hpnLine As Long)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- create pen if neccessary
    If clrLine <> -1 Then
        hpnLine = ApiCreatePen(PS_SOLID, 1, TranslateColor(clrLine))
    End If
    Pen = hpnLine
    '//--- draw line
    If x1 >= 0 Then
        MoveToEx m_MemoryDC, x1, y1, 0
    End If
    LineTo m_MemoryDC, x2, y2
    On Error GoTo 0

End Sub

'Purpose: Returns or sets the ROP2 drawing mode for the device context of the <b>IAPP_DrawDC</b> object.

Private Property Get DrawMode() As DrawModeConstants

    If IsCreated() Then
        DrawMode = GetROP2(m_MemoryDC)
    End If

End Property

Private Property Let DrawMode(ByVal dwValue As DrawModeConstants)

    If IsCreated() Then
        SetROP2 m_MemoryDC, dwValue
    End If

End Property

'Purpose: Draws a text using specified drawing flags in a rectangle in the device context.

Public Sub DrawText(ByVal Text As String, _
                    ByRef LeftX As Long, _
                    ByRef TopY As Long, _
                    ByRef RightX As Long, _
                    ByRef BottomY As Long, _
                    Optional ByVal dtFlags As UcsDrawTextStyles)

    Dim wTextParams As DRAWTEXTPARAMS
    Dim rc     As RECT

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    With rc
        .Left = LeftX
        .Top = TopY
        .Right = RightX
        .Bottom = BottomY
    End With                                          'RC
    wTextParams.cbSize = Len(wTextParams)
    ApiDrawTextEx m_MemoryDC, Text, -1, rc, dtFlags, wTextParams
    With rc
        LeftX = .Left
        TopY = .Top
        RightX = .Right
        BottomY = .Bottom
    End With                                          'RC
    On Error GoTo 0

End Sub

'Purpose: Fills a rectangle in the device context with a horizontal or vertical gradient.

Public Sub FillGradient(ByVal LeftX As Long, _
                        ByVal TopY As Long, _
                        ByVal RightX As Long, _
                        ByVal BottomY As Long, _
                        ByVal clrFirst As OLE_COLOR, _
                        ByVal clrSecond As OLE_COLOR, _
                        Optional ByVal bVertical As Boolean)

    Dim pVert(0 To 1) As TRIVERTEX
    Dim clr    As OLE_COLOR
    Dim pGradRect As GRADIENT_RECT

    clr = TranslateColor(clrFirst)
    With pVert(0)
        .X = LeftX
        .Y = TopY
        .Red = pvRed(clr)
        .Green = pvGreen(clr)
        .Blue = pvBlue(clr)
    End With                                          'PVERT(0)
    clr = TranslateColor(clrSecond)
    With pVert(1)
        .X = RightX
        .Y = BottomY
        .Red = pvRed(clr)
        .Green = pvGreen(clr)
        .Blue = pvBlue(clr)
    End With                                          'PVERT(1)
    With pGradRect
        .UpperLeft = 0
        .LowerRight = 1
    End With                                          'PGRADRECT
    GradientFill m_MemoryDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

'Purpose: Fills a rectangle in the device context with a specified color or brush.

Public Sub FillRect(Optional ByVal LeftX As Long, _
                    Optional ByVal TopY As Long, _
                    Optional ByVal RightX As Long = -1, _
                    Optional ByVal BottomY As Long = -1, _
                    Optional ByVal clrFill As Long = -1, _
                    Optional ByVal hbrFill As Long)

    Dim rc     As RECT

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- create brush if neccessary
    If clrFill <> -1 Then
        hbrFill = CreateSolidBrush(TranslateColor(clrFill))
    End If
    With rc
        .Left = LeftX
        .Top = TopY
        .Right = IIf(RightX < LeftX, Width, RightX)
        .Bottom = IIf(BottomY < TopY, Height, BottomY)
    End With                                          'RC
    ApiFillRect m_MemoryDC, rc, hbrFill
    '//--- cleanup the brush (if neccessary)
    If clrFill <> -1 Then
        APIDeleteObject hbrFill
    End If
    On Error GoTo 0

End Sub

'Purpose: Returns or sets the font currently selected in the <b>IAPP_DrawDC</b> object.

Public Property Get Font() As StdFont

    Dim tm     As TEXTMETRIC
    Dim sFaceName As String * 80

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    GetTextMetrics m_MemoryDC, tm
    GetTextFace m_MemoryDC, 79, sFaceName
    Set Font = New StdFont
    With Font
        .Name = sFaceName                             'StrConv(sFaceName, vbUnicode)
        .Bold = (tm.tmWeight >= FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
    End With                                          'FONT
    On Error GoTo 0

End Property

Public Property Set Font(ByVal oValue As StdFont)

'Dim nName() As Byte
'Dim I As Byte
'Dim nSize As Byte

    Dim tFont  As LOGFONT

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    With tFont
        CopyMemory .lfFaceName(1), ByVal oValue.Name, Len(oValue.Name) + 1
        .lfCharSet = oValue.Charset
        .lfItalic = (-oValue.Italic)
        .lfStrikeOut = (-oValue.Strikethrough)
        .lfUnderline = (-oValue.Underline)
        .lfWeight = oValue.Weight
        .lfHeight = -(oValue.Size * GetDeviceCaps(m_MemoryDC, LOGPIXELSY) \ 72)
    End With                                          'TFONT
    If m_MemoryFont <> 0 Then
        SelectObject m_MemoryDC, m_OrginalFont
        APIDeleteObject m_MemoryFont
    End If
    m_MemoryFont = CreateFontIndirect(tFont)
    m_OrginalFont = SelectObject(m_MemoryDC, m_MemoryFont)
    On Error GoTo 0

End Property

'Purpose: Returns the height of the currently selected fron in a <b>IAPP_DrawDC</b> object in pixels.

Public Property Get FontHeight() As Long

    Dim tm     As TEXTMETRIC

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    GetTextMetrics m_MemoryDC, tm
    FontHeight = tm.tmHeight
    On Error GoTo 0

End Property

'Purpose: Returns or sets the foreground color for the device context of the <b>IAPP_DrawDC</b> object.

Public Property Get ForeColor() As Long

    If IsCreated() Then
        ForeColor = GetTextColor(m_MemoryDC)
    End If

End Property

Public Property Let ForeColor(ByVal newValue As Long)

    If IsCreated() Then
        SetTextColor m_MemoryDC, TranslateColor(newValue)
    End If

End Property

'Purpose: Gets the value of the color of a specifies pixel.

Public Function GetPixel(ByVal LeftX As Long, _
                         ByVal TopY As Long) As Long

    GetPixel = ApiGetPixel(m_MemoryDC, LeftX, TopY)

End Function

'//--------------------------------------------------------------------------------
'//--Properties
'//--------------------------------------------------------------------------------
'Purpose: Returns a handle provided by the Microsoft Windows operating environment to the device context of a <b>IAPP_DrawDC</b> object.

Public Property Get HDC() As Long

    HDC = m_MemoryDC

End Property

'Purpose: Returns a Long value indicating the height of the <b>IAPP_DrawDC</b> object in pixels.

Public Property Get Height() As Long

    Height = m_MemoryHeight

End Property

Public Function HM2Pix(ByVal dblValue As Double) As Double

    HM2Pix = pvHM2Pix(dblValue)

End Function

'Purpose: Creates a <b>StdPicture</b> object from HICON handle.

Public Function IconToPicture(ByVal hIcon As Long) As IPicture

'//--- Returns a VB picture object containing the specified icon.

    Dim oNewPic As Picture
    Dim lpPictDesc As PictDesc
    Dim aGuid(0 To 3) As Long

    On Error Resume Next
    '//--- check argument
    If hIcon = 0 Then
        Exit Function                                 '---> Bottom
    End If
    '//--- fill struct
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeIcon
        .hBmpOrIcon = hIcon
    End With                                          'LPPICTDESC
    '//--- fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '//--- do convert
    OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
    '//--- success
    Set IconToPicture = oNewPic
    On Error GoTo 0

End Function

'Purpose: Returns a <b>StdPicture</b> objects which contains current image of the <b>IAPP_DrawDC</b> object.

Public Property Get Image() As StdPicture

    Dim hdcPaint As Long
    Dim hbmPaint As Long
    Dim hbmPaintOrig As Long
    Dim hpalPaintOrig As Long

    '//--- state check

    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    '//--- prepare
    hdcPaint = CreateCompatibleDC(m_MemoryDC)
    hbmPaint = CreateCompatibleBitmap(m_MemoryDC, m_MemoryWidth, m_MemoryHeight)
    hbmPaintOrig = SelectObject(hdcPaint, hbmPaint)
    If m_MemoryPal <> 0 Then
        hpalPaintOrig = SelectPalette(hdcPaint, m_MemoryPal, 0)
        RealizePalette hdcPaint
    End If
    '//--- bitblit
    ApiBitBlt hdcPaint, 0, 0, m_MemoryWidth, m_MemoryHeight, m_MemoryDC, 0, 0, vbSrcCopy
    '//--- deselect
    SelectObject hdcPaint, hbmPaintOrig
    If hpalPaintOrig <> 0 Then
        SelectPalette hdcPaint, hpalPaintOrig, 0
        RealizePalette hdcPaint
    End If
    DeleteDC hdcPaint
    '//--- get image
    Set Image = BitmapToPicture(hbmPaint, m_MemoryPal)

End Property

'//--------------------------------------------------------------------------------
'//--Methods
'//--------------------------------------------------------------------------------
'Purpose: Initializes <b>IAPP_DrawDC</b> object with certain device context. Optionally allocates a memory bitmap to cache drawing operations.

Public Function Init(Optional ByVal PixelWidth As Long, _
                     Optional ByVal PixelHeight As Long, _
                     Optional ByVal hParentDC As Long, _
                     Optional ByVal hMemoryDC As Long) As Long

    Dim nHasPalette As Long
    Dim nPaletteSize As Long
    Dim LogPal As LOGPALETTE
    Dim tm     As TEXTMETRIC
    Dim sFaceName As String * 80
    Dim oFont  As StdFont

    If IsCreated() Then
        Destroy
    End If
    '//--- local vars
    m_MemoryWidth = PixelWidth
    m_MemoryHeight = PixelHeight
    '//--- check if parent dc supplied
    If hParentDC = 0 Then
        m_ParentDC = GetDC(0)
    Else                                              'NOT HPARENTDC...
        m_ParentDC = hParentDC
    End If
    '//--- check if dc already supplied
    If hMemoryDC = 0 Then
        '//--Create a memory device context to use
        m_MemoryDC = CreateCompatibleDC(m_ParentDC)
        '//--Tell'em it's a picture (so drawings can be done on the DC)
        m_MemoryBitmap = CreateCompatibleBitmap(m_ParentDC, m_MemoryWidth, m_MemoryHeight)
        m_OrginalBitmap = SelectObject(m_MemoryDC, m_MemoryBitmap)
    Else                                              'NOT HMEMORYDC...
        m_MemoryDC = hMemoryDC
        m_MemoryBitmap = 0
    End If
    '//--Get screen properties
    nHasPalette = GetDeviceCaps(m_ParentDC, RASTERCAPS) And RC_PALETTE    '//--Palette support
    nPaletteSize = GetDeviceCaps(m_ParentDC, SIZEPALETTE)    '//--Size of palette
    '//--If the screen has a palette make a copy and realize it
    If nHasPalette And (nPaletteSize = 256) Then
        '//--Create a copy of the system palette
        With LogPal
            .palVersion = &H300
            .palNumEntries = 256
            GetSystemPaletteEntries m_ParentDC, 0&, 256, .palPalEntry(0)
        End With                                      'LogPal
        m_MemoryPal = CreatePalette(LogPal)
        '//--Select the new palette into the memory DC and realize it
        m_OrginalPal = SelectPalette(m_MemoryDC, m_MemoryPal, 0&)
        RealizePalette m_MemoryDC
    End If
    If hParentDC <> 0 Then
        '//--- Set attributes. Take from parent.
        SetBkColor m_MemoryDC, GetBkColor(m_ParentDC)
        SetTextColor m_MemoryDC, GetTextColor(m_ParentDC)
        SetBkMode m_MemoryDC, GetBkMode(m_ParentDC)
        GetTextMetrics m_ParentDC, tm
        GetTextFace m_ParentDC, 79, sFaceName
        Set oFont = New StdFont
        With oFont
            .Bold = (tm.tmWeight > FW_NORMAL)
            .Charset = tm.tmCharSet
            .Italic = (tm.tmItalic <> 0)
            .Name = sFaceName
            .Strikethrough = (tm.tmStruckOut <> 0)
            .Underline = (tm.tmUnderlined <> 0)
            .Weight = tm.tmWeight
            .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
        End With                                      'OFONT
        Set Font = oFont
        Set oFont = Nothing
    Else                                              '//--- If hParentDC = 0 Then 'NOT HPARENTDC...
        ReleaseDC 0, m_ParentDC
        m_ParentDC = 0
    End If
    '//--- success
    Init = m_MemoryDC

End Function

'Purpose: Tests if current <b>IAPP_DrawDC</b> is already initialized.

Public Function IsCreated() As Boolean

    IsCreated = (m_MemoryDC <> 0)

End Function

'Purpose: Returns a value indication whether current device context operations are cached in memory bitmap.

Private Property Get IsMemoryDC() As Boolean

    IsMemoryDC = (m_MemoryBitmap <> 0)

End Property

'Purpose: Normalizes a rectangle dimensions ensuring Left <= Right and Top <= Bottom.

Public Sub NormalizeRect(LeftX As Long, _
                         TopY As Long, _
                         RightX As Long, _
                         BottomY As Long)

    Dim lTmp   As Long

    If BottomY < TopY Then
        lTmp = BottomY
        BottomY = TopY
        TopY = lTmp
    End If
    If RightX < LeftX Then
        lTmp = RightX
        RightX = LeftX
        LeftX = lTmp
    End If

End Sub

'Purpose: Paints a <b>StdPicture</b> objects in a rectangle in the device context effectively creating a disabled copy of its image.

Public Sub PaintDisabledPicture(ByVal oPic As StdPicture, _
                                Optional ByVal xDest As Long, _
                                Optional ByVal yDest As Long, _
                                Optional ByVal nWidth As Long, _
                                Optional ByVal nHeight As Long, _
                                Optional ByVal xSrc As Long, _
                                Optional ByVal ySrc As Long, _
                                Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                                Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                                Optional ByVal clrShadow As OLE_COLOR = vb3DShadow)

    Dim oMemDC As IAPP_DrawDC

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Init pvHM2Pix(oPic.Width), pvHM2Pix(oPic.Height)
    End If
    '//--- fix arguments
    If nWidth = 0 Then
        nWidth = pvHM2Pix(oPic.Width)
    End If
    If nHeight = 0 Then
        nHeight = pvHM2Pix(oPic.Height)
    End If
    '//--- paint to temp dc
    Set oMemDC = New IAPP_DrawDC
    With oMemDC
        .Init nWidth, nHeight
        .Cls clrMask
        .PaintPicture oPic, 0, 0, nWidth, nHeight, xSrc, ySrc
    End With                                          'OMEMDC
    pvDisabledBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, oMemDC.HDC, 0, 0, clrMask, clrHighlight, clrShadow
    On Error GoTo 0

End Sub

'Purpose: Paints a <b>StdPicture</b> objects in a rectangle in the device context using specified raster operation or skipping mask color pixels.

Public Sub PaintPicture(ByVal oPic As StdPicture, _
                        Optional ByVal xDest As Long, _
                        Optional ByVal yDest As Long, _
                        Optional ByVal nWidth As Long, _
                        Optional ByVal nHeight As Long, _
                        Optional ByVal xSrc As Long, _
                        Optional ByVal ySrc As Long, _
                        Optional ByVal dwRop As RasterOpConstants = vbSrcCopy, _
                        Optional ByVal clrMask As OLE_COLOR = -1)

    Dim rc     As RECT
    Dim hdcPaint As Long
    Dim hbmOrig As Long
    Dim hEmf   As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Init pvHM2Pix(oPic.Width), pvHM2Pix(oPic.Height)
    End If
    '//--- fix arguments
    If nWidth = 0 Then
        nWidth = pvHM2Pix(oPic.Width)
    End If
    If nHeight = 0 Then
        nHeight = pvHM2Pix(oPic.Height)
    End If
    clrMask = TranslateColor(clrMask)
    '//--- draw
    Select Case oPic.Type
        Case vbPicTypeIcon
            DrawIconEx m_MemoryDC, xDest, yDest, oPic.handle, nWidth, nHeight, 0, 0, DI_NORMAL
        Case vbPicTypeBitmap
            hdcPaint = CreateCompatibleDC(m_MemoryDC)
            hbmOrig = SelectObject(hdcPaint, oPic.handle)
            If clrMask = -1 Then
                ApiBitBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, hdcPaint, xSrc, ySrc, dwRop
            Else                                      'NOT CLRMASK...
                pvTransBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, hdcPaint, xSrc, ySrc, clrMask
            End If
            SelectObject hdcPaint, hbmOrig
            DeleteDC hdcPaint
        Case vbPicTypeEMetafile, vbPicTypeMetafile
            With rc
                .Left = xDest
                .Top = yDest
                .Right = xDest + nWidth
                .Bottom = yDest + nHeight
            End With                                  'RC
            If oPic.Type = vbPicTypeMetafile Then
                hdcPaint = CreateEnhMetaFileLong(m_MemoryDC, vbNullString, 0, vbNullString)
                PlayMetaFile hdcPaint, oPic.handle
                hEmf = CloseEnhMetaFile(hdcPaint)
            Else                                      'NOT OPIC.TYPE...
                hEmf = oPic.handle
            End If
            PlayEnhMetaFile m_MemoryDC, hEmf, rc
            If oPic.Type = vbPicTypeMetafile Then
                DeleteEnhMetaFile hEmf
            End If
    End Select
    On Error GoTo 0

End Sub

'Purpose: Returns or sets the currently selected GDI pen (HPEN) for the device context of the <b>IAPP_DrawDC</b> object.

Public Property Get Pen() As Long

    Pen = m_MemoryPen

End Property

Public Property Let Pen(ByVal lValue As Long)

    Dim hPrevPen As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Property                                 '---> Bottom
    End If
    If lValue <> 0 Then
        '//--- set new pen (and save original)
        hPrevPen = SelectObject(m_MemoryDC, lValue)
        If m_OriginalPen = 0 Then
            m_OriginalPen = hPrevPen
        End If
    ElseIf m_OriginalPen <> 0 Then                    'NOT LVALUE...
        '//--- restore original pen
        hPrevPen = SelectObject(m_MemoryDC, m_OriginalPen)
        m_OriginalPen = 0
    End If
    '//--- delete if previously selected
    If m_MemoryPen = hPrevPen Then
        APIDeleteObject hPrevPen
    End If
    '//--- success
    m_MemoryPen = lValue
    On Error GoTo 0

End Property

Private Function pvBlue(ByVal clr As OLE_COLOR) As Long

    pvBlue = ((clr \ &H10000) And &HFF) * &H100&
    If pvBlue >= &H8000& Then
        pvBlue = pvBlue - &H10000
    End If

End Function

Private Sub pvDisabledBlt(ByVal hdcDest As Long, _
                          ByVal xDest As Long, _
                          ByVal yDest As Long, _
                          ByVal nWidth As Long, _
                          ByVal nHeight As Long, _
                          ByVal hdcSrc As Long, _
                          Optional ByVal xSrc As Long = 0, _
                          Optional ByVal ySrc As Long = 0, _
                          Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                          Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                          Optional ByVal clrShadow As OLE_COLOR = vb3DShadow, _
                          Optional ByVal hPal As Long = 0)

    Dim hdcScreen As Long
    Dim hbmMonoSection As Long
    Dim hbmMonoSectionSav As Long
    Dim hdcMonoSection As Long

    'Dim hdcColor As Long
    Dim hdcDisabled As Long
    Dim hbmDisabledSav As Long
    Dim lpBI   As BITMAPINFO
    Dim hbmMono As Long
    Dim hdcMono As Long
    Dim hbmMonoSav As Long
    Dim lMaskColor As Long
    Dim lMaskColorCompare As Long
    Dim hdcMaskedSource As Long
    Dim hbmMasked As Long
    Dim hbmMaskedOld As Long
    Dim hpalMaskedOld As Long
    Dim hpalDisabledOld As Long
    Dim hpalMonoOld As Long
    Dim rgbBlack As RGBQUAD
    Dim rgbWhite As RGBQUAD
    Dim dwSys3dShadow As Long
    Dim dwSys3dHighlight As Long
    Dim pvBits As Long
    Dim rgbnew(1) As RGBQUAD
    Dim hbmDisabled As Long
    Dim lMonoBkGrnd As Long
    Dim lMonoBkGrndChoices(2) As Long
    Dim lIndex As Long
    Dim hbrWhite As Long
    Dim udtRect As RECT
    Dim hpalHalftone As Long

    hdcScreen = GetDC(0&)
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    '//--Define some colors
    OleTranslateColor clrShadow, hPal, dwSys3dShadow
    dwSys3dShadow = dwSys3dShadow And &HFFFFFF
    OleTranslateColor clrHighlight, hPal, dwSys3dHighlight
    dwSys3dHighlight = dwSys3dHighlight And &HFFFFFF
    With rgbBlack
        .rgbBlue = 0
        .rgbGreen = 0
        .rgbRed = 0
        .rgbReserved = 0
    End With                                          'RGBBLACK
    With rgbWhite
        .rgbBlue = 255
        .rgbGreen = 255
        .rgbRed = 255
        .rgbReserved = 255
    End With                                          'RGBWHITE
    '//--The first step is to create a monochrome bitmap with two colors:
    '//--white where colors in the original are light, and black
    '//--where the original is dark.  We can't simply ApiBitBlt to a bitmap.
    '//--Instead, we create a monochrome (bichrome?) DIB section and ApiBitBlt
    '//--to that.  Windows will do the conversion automatically based on the
    '//--DIB section's palette.  (I.e. using a DIB section, Windows knows how
    '//--to map "light" colors and "dark" colors to white/black, respectively.
    With lpBI.bmiHeader
        .biSize = LenB(lpBI.bmiHeader)
        .biWidth = nWidth
        .biHeight = -nHeight
        .biPlanes = 1
        .biBitCount = 1                               '//--monochrome
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0                                '//--max colors used (2^1 = 2)
        .biClrImportant = 0                           '//--all (both :-]) colors are important
    End With                                          'LPBI.BMIHEADER
    With lpBI
        .bmiColors(0) = rgbBlack
        .bmiColors(1) = rgbWhite
    End With                                          'LPBI
    hbmMonoSection = CreateDIBSection(hdcScreen, lpBI, DIB_RGB_COLORS, pvBits, 0&, 0)
    hdcMonoSection = CreateCompatibleDC(hdcScreen)
    hbmMonoSectionSav = SelectObject(hdcMonoSection, hbmMonoSection)
    '//--ApiBitBlt to the Monochrome DIB section
    '//--If a mask color is provided, create a new bitmap and copy the source
    '//--to it transparently.  If we don't do this, a dark mask color will be
    '//--turned into the outline part of the monochrome DIB section
    '//--Convert mask color and white before comparing
    '//--because the Mask color might be a system color that would be evaluated
    '//--to white.
    OleTranslateColor vbWhite, hPal, lMaskColorCompare
    lMaskColorCompare = lMaskColorCompare And &HFFFFFF
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    If lMaskColor = lMaskColorCompare Then
        ApiBitBlt hdcMonoSection, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy
    Else                                              'NOT LMASKCOLOR...
        hbmMasked = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
        hdcMaskedSource = CreateCompatibleDC(hdcScreen)
        hbmMaskedOld = SelectObject(hdcMaskedSource, hbmMasked)
        hpalMaskedOld = SelectPalette(hdcMaskedSource, hPal, True)
        RealizePalette hdcMaskedSource
        '//--Fill the bitmap with white
        With udtRect
            .Left = 0
            .Top = 0
            .Right = nWidth
            .Bottom = nHeight
        End With                                      'UDTRECT
        hbrWhite = CreateSolidBrush(vbWhite)
        ApiFillRect hdcMaskedSource, udtRect, hbrWhite
        APIDeleteObject hbrWhite
        '//--Do the transparent paint
        pvTransBlt hdcMaskedSource, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, lMaskColor, hPal
        '//--ApiBitBlt to the Mono DIB section.  The mask color has been turned to white.
        ApiBitBlt hdcMonoSection, 0, 0, nWidth, nHeight, hdcMaskedSource, 0, 0, vbSrcCopy
        '//--Clean up
        SelectPalette hdcMaskedSource, hpalMaskedOld, 1
        RealizePalette hdcMaskedSource
        APIDeleteObject SelectObject(hdcMaskedSource, hbmMaskedOld)
        DeleteDC hdcMaskedSource
    End If
    '//--Okay, we've got our B&W DIB section.
    '//--Now that we have our monochrome bitmap, the final appearance that we
    '//--want is this:  First, think of the black portion of the monochrome
    '//--bitmap as our new version of the original bitmap.  We want to have a dark
    '//--gray version of this with a light version underneath it, shifted down and
    '//--to the right.  The light acts as a highlight, and it looks like the original
    '//--image is a gray inset.
    '
    '//--First, create a copy of the destination.  Draw the light gray transparently,
    '//--and then draw the dark gray transparently
    hbmDisabled = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    hdcDisabled = CreateCompatibleDC(hdcScreen)
    hbmDisabledSav = SelectObject(hdcDisabled, hbmDisabled)
    hpalDisabledOld = SelectPalette(hdcDisabled, hPal, 1)
    RealizePalette hdcDisabled
    '//--We used to fill the background with gray, instead copy the
    '//--destination to memory DC.  This will allow a disabled image
    '//--to be drawn over a background image.
    ApiBitBlt hdcDisabled, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    '//--When painting the monochrome bitmaps transparently onto the background
    '//--we need a background color that is not the light color of the dark color
    '//--Provide three choices to ensure a unique color is picked.
    OleTranslateColor vbBlack, hPal, lMonoBkGrndChoices(0)
    lMonoBkGrndChoices(0) = lMonoBkGrndChoices(0) And &HFFFFFF
    OleTranslateColor vbRed, hPal, lMonoBkGrndChoices(1)
    lMonoBkGrndChoices(1) = lMonoBkGrndChoices(1) And &HFFFFFF
    OleTranslateColor vbBlue, hPal, lMonoBkGrndChoices(2)
    lMonoBkGrndChoices(2) = lMonoBkGrndChoices(2) And &HFFFFFF
    '//--Pick a background color choice that doesn't match
    '//--the shadow or highlight color
    For lIndex = 0 To 2
        If lMonoBkGrndChoices(lIndex) <> dwSys3dHighlight Then
            If lMonoBkGrndChoices(lIndex) <> dwSys3dShadow Then
                'This color can be used for a mask
                lMonoBkGrnd = lMonoBkGrndChoices(lIndex)
                Exit For                              'loop varying lindex
            End If
        End If
    Next lIndex
    '//--Now paint a the light color shifted and transparent over the background
    '//--It is not necessary to change the DIB section's color table
    '//--to equal the highlight color and mask color.  In fact, setting
    '//--the color table to anything besides black and white causes unpredictable
    '//--results (seen in win95 with IE4, using 256 colors).
    '//--Setting the Back and Text colors of the Monochrome bitmap, ensure
    '//--that the desired colors are produced.
    With rgbnew(0)
        .rgbRed = (vbWhite \ 2 ^ 16) And &HFF
        .rgbGreen = (vbWhite \ 2 ^ 8) And &HFF
        .rgbBlue = vbWhite And &HFF
    End With                                          'RGBNEW(0)
    With rgbnew(1)
        .rgbRed = (vbBlack \ 2 ^ 16) And &HFF
        .rgbGreen = (vbBlack \ 2 ^ 8) And &HFF
        .rgbBlue = vbBlack And &HFF
    End With                                          'RGBNEW(1)
    SetDIBColorTable hdcMonoSection, 0, 2, rgbnew(0)
    '//--...We can't pass a DIBSection to pvTransBlt(), so we need to
    '//--make a copy of our mono DIBSection.  Notice that we only need a monochrome
    '//--bitmap, but we must set its back/fore colors to the monochrome colors we
    '//--want (light gray and black), and pvTransBlt() will honor them.
    hbmMono = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    hdcMono = CreateCompatibleDC(hdcScreen)
    hbmMonoSav = SelectObject(hdcMono, hbmMono)
    SetMapMode hdcMono, GetMapMode(hdcSrc)
    SetBkColor hdcMono, dwSys3dHighlight
    SetTextColor hdcMono, lMonoBkGrnd
    hpalMonoOld = SelectPalette(hdcMono, hPal, True)
    RealizePalette hdcMono
    ApiBitBlt hdcMono, 0, 0, nWidth, nHeight, hdcMonoSection, 0, 0, vbSrcCopy
    '//--...We can go ahead and call pvTransBlt with our monochrome
    '//--copy
    '//--Draw this transparently over the disabled bitmap
    '//--...Don't forget to shift right and left....
    pvTransBlt hdcDisabled, 1, 1, nWidth, nHeight, hdcMono, 0, 0, lMonoBkGrnd, hPal
    '//--Now draw a transparent copy, using dark gray where the monochrome had
    '//--black, and transparent elsewhere.  We'll use a transparent color of black.
    '
    '//--...We can't pass a DIBSection to pvTransBlt(), so we need to
    '//--make a copy of our mono DIBSection.  Notice that we only need a monochrome
    '//--bitmap, but we must set its back/fore colors to the monochrome colors we
    '//--want (dark gray and black), and pvTransBlt() will honor them.
    '//--Use hbmMono and hdcMono; already created for first color
    SetBkColor hdcMono, dwSys3dShadow
    SetTextColor hdcMono, lMonoBkGrnd
    ApiBitBlt hdcMono, 0, 0, nWidth, nHeight, hdcMonoSection, 0, 0, vbSrcCopy
    '//--...We can go ahead and call pvTransBlt with our monochrome
    '//--copy
    '//--Draw this transparently over the disabled bitmap
    pvTransBlt hdcDisabled, 0, 0, nWidth, nHeight, hdcMono, 0, 0, lMonoBkGrnd, hPal
    ApiBitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcDisabled, 0, 0, vbSrcCopy
    '//--Okay, we're done!
    SelectPalette hdcDisabled, hpalDisabledOld, 1
    RealizePalette hdcDisabled
    APIDeleteObject SelectObject(hdcMonoSection, hbmMonoSectionSav)
    DeleteDC hdcMonoSection
    APIDeleteObject SelectObject(hdcDisabled, hbmDisabledSav)
    DeleteDC hdcDisabled
    APIDeleteObject SelectObject(hdcMono, hbmMonoSav)
    SelectPalette hdcMono, hpalMonoOld, 1
    RealizePalette hdcMono
    DeleteDC hdcMono
    ReleaseDC 0&, hdcScreen
    If hpalHalftone <> 0 Then
        APIDeleteObject hpalHalftone
    End If

End Sub

Private Function pvGreen(ByVal clr As OLE_COLOR) As Long

    pvGreen = ((clr \ &H100) And &HFF) * &H100&
    If pvGreen >= &H8000& Then
        pvGreen = pvGreen - &H10000
    End If

End Function

Private Function pvHM2Pix(ByVal dblValue As Double) As Double

    pvHM2Pix = dblValue * 1440 / 2540 / Screen.TwipsPerPixelX

End Function

Private Function pvRed(ByVal clr As OLE_COLOR) As Long

    pvRed = ((clr \ &H1) And &HFF) * &H100&
    If pvRed >= &H8000& Then
        pvRed = pvRed - &H10000
    End If

End Function

Private Sub pvTransBlt(ByVal hdcDest As Long, _
                       ByVal xDest As Long, _
                       ByVal yDest As Long, _
                       ByVal nWidth As Long, _
                       ByVal nHeight As Long, _
                       ByVal hdcSrc As Long, _
                       Optional ByVal xSrc As Long = 0, _
                       Optional ByVal ySrc As Long = 0, _
                       Optional ByVal clrMask As OLE_COLOR = MASK_COLOR, _
                       Optional ByVal hPal As Long = 0)

    Dim hdcMask As Long                               '//--hDC of the created mask image
    Dim hdcColor As Long                              '//--hDC of the created color image
    Dim hbmMask As Long                               '//--Bitmap handle to the mask image
    Dim hbmColor As Long                              '//--Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hpalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long                          '//--Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    Dim hpalHalftone As Long

    hdcScreen = GetDC(0&)
    '//--Validate palette
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    '//--Create a color bitmap to server as a copy of the destination
    '//--Do all work on this bitmap and then copy it back over the destination
    '//--when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    '//--Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    '//--Copy the destination to the screen buffer
    ApiBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    '//--Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    '//--hdcSrc, because this will create a DIB section if the original bitmap
    '//--is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    '//--Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    '//--First, blt the source bitmap onto the cover.  We do this first
    '//--and then use it instead of the source bitmap
    '//--because the source bitmap may be
    '//--a DIB section, which behaves differently than a bitmap.
    '//--(Specifically, copying from a DIB section to a monochrome bitmap
    '//--does a nearest-color selection rather than painting based on the
    '//--backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    '//--In case hdcSrc contains a monochrome bitmap, we must set the destination
    '//--foreground/background colors according to those currently set in hdcSrc
    '//--(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    ApiBitBlt hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy
    '//--Paint the mask.  What we want is white at the transparent color
    '//--from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)
    '//--When ApiBitBlt'ing from color to monochrome, Windows sets to 1
    '//--all pixels that match the background color of the source DC.  All
    '//--other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    ApiBitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
    '//--Paint the rest of the cover bitmap.
    '
    '//--What we want here is black at the transparent color, and
    '//--the original colors everywhere else.  To do this, we first
    '//--paint the original onto the cover (which we already did), then we
    '//--AND the inverse of the mask onto that using the DSna ternary raster
    '//--operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    '//--Operation Codes", "Ternary Raster Operations", or search in MSDN
    '//--for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    '//--When ApiBitBlt'ing from monochrome to color, Windows transforms all white
    '//--bits (1) to the background color of the destination hDC.  All black (0)
    '//--bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    ApiBitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna
    '//--Paint the Mask to the Screen buffer
    ApiBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
    '//--Paint the Color to the Screen buffer
    ApiBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
    '//--Copy the screen buffer to the screen
    ApiBitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
    '//--All done!
    APIDeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hpalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    APIDeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, 0
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    APIDeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    If hpalHalftone <> 0 Then
        APIDeleteObject hpalHalftone
    End If

End Sub

'Purpose: Fills a rectangle in the device context optionally using a specified color for pen and brush.

Public Sub Rectangle(Optional ByVal LeftX As Long, _
                     Optional ByVal TopY As Long, _
                     Optional ByVal RightX As Long = -1, _
                     Optional ByVal BottomY As Long = -1, _
                     Optional ByVal clrFill As OLE_COLOR = -1, _
                     Optional ByVal hbrFill As Long, _
                     Optional ByVal clrOutline As OLE_COLOR = -1, _
                     Optional ByVal hpnOutline As Long)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- create brush if neccessary
    If clrFill <> -1 Then
        hbrFill = CreateSolidBrush(TranslateColor(clrFill))
    End If
    If hbrFill <> 0 Then
        Brush = hbrFill
    End If
    '//--- create pen if neccessary
    If clrOutline <> -1 Then
        hpnOutline = ApiCreatePen(PS_SOLID, 1, TranslateColor(clrOutline))
    End If
    If hpnOutline <> 0 Then
        Pen = hpnOutline
    End If
    APIRectangle m_MemoryDC, LeftX, TopY, IIf(RightX < LeftX, Width, RightX), IIf(BottomY < TopY, Height, BottomY)
    On Error GoTo 0

End Sub

'Purpose: Set the color of a pixel in the device context to a specified value.

Public Sub SetPixel(ByVal LeftX As Long, _
                    ByVal TopY As Long, _
                    Optional clrColor As Long = vbWhite)

    ApiSetPixel m_MemoryDC, LeftX, TopY, TranslateColor(clrColor)

End Sub

'Purpose: Returns a <b>StdFont</b> object representing current font used by Windows to draw icons captions.

Private Property Get SystemIconFont() As StdFont

    Dim lF     As LOGFONT
    Dim sBuffer As String

    'Dim hr As Long
    Dim hTempDC As Long

    SystemParametersInfo SPI_GETICONTITLELOGFONT, LenB(lF), lF, 0
    Set SystemIconFont = New StdFont
    With SystemIconFont
        sBuffer = Space$(lstrlen(lF.lfFaceName(1)))
        CopyMemory ByVal sBuffer, lF.lfFaceName(1), Len(sBuffer)
        .Name = sBuffer
        .Bold = (lF.lfWeight >= FW_NORMAL)
        .Charset = lF.lfCharSet
        .Italic = (lF.lfItalic <> 0)
        .Strikethrough = (lF.lfStrikeOut <> 0)
        .Underline = (lF.lfUnderline <> 0)
        .Weight = lF.lfWeight
        hTempDC = GetDC(0)
        .Size = -(lF.lfHeight * 72) / GetDeviceCaps(hTempDC, LOGPIXELSY)
        ReleaseDC 0, hTempDC
    End With                                          'SYSTEMICONFONT

End Property

Public Property Get SystemMenuFont() As StdFont

    Dim lF     As LOGFONT
    Dim sBuffer As String

    'Dim hr As Long
    Dim hTempDC As Long
    Dim ncm    As NONCLIENTMETRICS

    ncm.cbSize = Len(ncm)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, LenB(ncm), ncm, 0
    lF = ncm.lfMenuFont
    Set SystemMenuFont = New StdFont
    With SystemMenuFont
        sBuffer = Space$(lstrlen(lF.lfFaceName(1)))
        CopyMemory ByVal sBuffer, lF.lfFaceName(1), Len(sBuffer)
        .Name = sBuffer
        .Bold = (lF.lfWeight >= FW_NORMAL)
        .Charset = lF.lfCharSet
        .Italic = (lF.lfItalic <> 0)
        .Strikethrough = (lF.lfStrikeOut <> 0)
        .Underline = (lF.lfUnderline <> 0)
        .Weight = lF.lfWeight
        hTempDC = GetDC(0)
        .Size = -(lF.lfHeight * 72) / GetDeviceCaps(hTempDC, LOGPIXELSY)
        ReleaseDC 0, hTempDC
    End With                                          'SYSTEMMENUFONT

End Property

'Purpose: Calculates the height of a text string using currently selected font settings.

Public Function TextHeight(ByVal sText As String) As Long

    Dim sz     As SIZEAPI

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Function                                 '---> Bottom
    End If
    GetTextExtentPoint m_MemoryDC, sText, Len(sText), sz
    '//--- success
    TextHeight = sz.cy
    On Error GoTo 0

End Function

'Purpose: Calculates the width of a text string using currently selected font settings.

Public Function TextWidth(ByVal sText As String) As Long

    Dim sz     As SIZEAPI

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Function                                 '---> Bottom
    End If
    GetTextExtentPoint m_MemoryDC, sText, Len(sText), sz
    '//--- success
    TextWidth = sz.cx
    On Error GoTo 0

End Function

'Private Sub pvTransBltOld(
'//--'//--           ByVal hdcDest As Long,
'//--'//--           ByVal xDest As Long,
'//--'//--           ByVal yDest As Long,
'//--'//--           ByVal nWidth As Long,
'//--'//--           ByVal nHeight As Long,
'//--'//--           ByVal hdcSrc As Long,
'//--'//--           Optional ByVal xSrc As Long = 0,
'//--'//--           Optional ByVal ySrc As Long = 0,
'//--'//--           Optional ByVal clrMask As Long = MASK_COLOR)
'//--   '
'//--   '//--32-Bit Transparent ApiBitBlt Function
'//--   '//--Written by Karl E. Peterson, 9/20/96.
'//--   '//--Portions borrowed and modified from KB.
'//--   '//--Other portions modified following input from users. <g>
'//--   '
'//--   '//--Parameters ************************************************************
'//--   '//--hdcDest: Destination device context
'//--   '//--x, y: Upper-left destination coordinates (pixels)
'//--   '//--nWidth: Width of destination
'//--   '//--nHeight: Height of destination
'//--   '//--hdcSrc: Source device context
'//--   '//--xSrc, ySrc: Upper-left source coordinates (pixels)
'//--   '//--clrMask: RGB value for transparent pixels, typically &HC0C0C0.
'//--   '//--***********************************************************************
'//--   '
'//--   Dim OrigColor As Long '//--Holds original background color
'//--   Dim OrigMode As Long '//--Holds original background drawing mode
'
'//--   If (GetDeviceCaps(hdcDest, CAPS1) And C1_TRANSPARENT) Then
'//--       '
'//--       '//--Some NT machines support this *super* simple method!
'//--       '//--Save original settings, Blt, restore settings.
'//--       '
'//--       OrigMode = SetBkMode(hdcDest, BS_NEWTRANSPARENT)
'//--       OrigColor = SetBkColor(hdcDest, clrMask)
'//--       Call ApiBitBlt(hdcDest, xDest, yDest, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
'//--       Call SetBkColor(hdcDest, OrigColor)
'//--       Call SetBkMode(hdcDest, OrigMode)
'//--   Else
'//--       Dim saveDC          As Long '//--Backup copy of source bitmap
'//--       Dim maskDC          As Long '//--Mask bitmap (monochrome)
'//--       Dim invDC           As Long '//--Inverse of mask bitmap (monochrome)
'//--       Dim resultDC        As Long '//--Combination of source bitmap & background
'//--       Dim screenDC        As Long
'//--       Dim hSaveBmp        As Long '//--Bitmap stores backup copy of source bitmap
'//--       Dim hMaskBmp        As Long '//--Bitmap stores mask (monochrome)
'//--       Dim hInvBmp         As Long '//--Bitmap holds inverse of mask (monochrome)
'//--       Dim hResultBmp      As Long '//--Bitmap combination of source & background
'//--       Dim hSavePrevBmp    As Long '//--Holds previous bitmap in saved DC
'//--       Dim hMaskPrevBmp    As Long '//--Holds previous bitmap in the mask DC
'//--       Dim hInvPrevBmp     As Long '//--Holds previous bitmap in inverted mask DC
'//--       Dim hDestPrevBmp    As Long '//--Holds previous bitmap in destination DC
'//--       '
'//--       '//--Create DCs to hold various stages of transformation.
'//--       '
'//--       screenDC = GetDC(0)
'//--       saveDC = CreateCompatibleDC(screenDC)
'//--       maskDC = CreateCompatibleDC(screenDC)
'//--       invDC = CreateCompatibleDC(screenDC)
'//--       resultDC = CreateCompatibleDC(screenDC)
'//--       '
'//--       '//--Create monochrome bitmaps for the mask-related bitmaps.
'//--       '
'//--       hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
'//--       hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
'//--       '
'//--       '//--Create color bitmaps for final result & stored copy of source.
'//--       '
'//--       hResultBmp = CreateCompatibleBitmap(hdcDest, nWidth, nHeight)
'//--       hSaveBmp = CreateCompatibleBitmap(hdcDest, nWidth, nHeight)
'//--       '
'//--       '//--Select bitmaps into DCs.
'//--       '
'//--       hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
'//--       hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
'//--       hInvPrevBmp = SelectObject(invDC, hInvBmp)
'//--       hDestPrevBmp = SelectObject(resultDC, hResultBmp)
'//--       '
'//--       '//--Create mask: set background color of source to transparent color.
'//--       '
'//--       OrigColor = SetBkColor(hdcSrc, clrMask)
'//--       Call ApiBitBlt(maskDC, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
'//--       clrMask = SetBkColor(hdcSrc, OrigColor)
'//--       '
'//--       '//--Create inverse of mask to AND w/ source & combine w/ background.
'//--       '
'//--       Call ApiBitBlt(invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy)
'//--       '
'//--       '//--Copy background bitmap to result.
'//--       '
'//--       Call ApiBitBlt(resultDC, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy)
'//--       '
'//--       '//--AND mask bitmap w/ result DC to punch hole in the background by
'//--       '//--painting black area for non-transparent portion of source bitmap.
'//--       '
'//--       Call ApiBitBlt(resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd)
'//--       '
'//--       '//--get overlapper
'//--       '
'//--       SetBkColor saveDC, GetBkColor(hdcSrc)
'//--       SetTextColor saveDC, GetTextColor(hdcSrc)
'//--       Call ApiBitBlt(saveDC, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
'//--       '
'//--       '//--AND with inverse monochrome mask
'//--       '
'//--       Call ApiBitBlt(saveDC, 0, 0, nWidth, nHeight, invDC, 0, 0, vbSrcAnd)
'//--       '
'//--       '//--XOR these two
'//--       '
'//--       Call ApiBitBlt(resultDC, 0, 0, nWidth, nHeight, saveDC, 0, 0, vbSrcInvert)
'//--       '
'//--       '//--Display transparent bitmap on background.
'//--       '
'//--       Call ApiBitBlt(hdcDest, xDest, yDest, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy)
'//--       '
'//--       '//--Select original objects back.
'//--       '
'//--       Call SelectObject(saveDC, hSavePrevBmp)
'//--       Call SelectObject(resultDC, hDestPrevBmp)
'//--       Call SelectObject(maskDC, hMaskPrevBmp)
'//--       Call SelectObject(invDC, hInvPrevBmp)
'//--       '
'//--       '//--Deallocate system resources.
'//--       '
'//--       Call ApiDeleteObject(hSaveBmp)
'//--       Call ApiDeleteObject(hMaskBmp)
'//--       Call ApiDeleteObject(hInvBmp)
'//--       Call ApiDeleteObject(hResultBmp)
'//--       Call DeleteDC(saveDC)
'//--       Call DeleteDC(invDC)
'//--       Call DeleteDC(maskDC)
'//--       Call DeleteDC(resultDC)
'//--       Call ReleaseDC(0, screenDC)
'//--   End If
'End Sub
'Purpose: Translates an OLE_COLOR value to a Long value which represents an RGB tripple.

Public Function TranslateColor(ByVal clrColor As OLE_COLOR) As Long

'//--- handle invalid (none) color

    If clrColor = -1 Then
        TranslateColor = -1
    Else                                              'NOT CLRCOLOR...
        OleTranslateColor clrColor, m_MemoryPal, TranslateColor
        TranslateColor = TranslateColor And &HFFFFFF
    End If

End Function

'Purpose: Returns a Long value indicating the width of the <b>IAPP_DrawDC</b> object in pixels.

Public Property Get Width() As Long

    Width = m_MemoryWidth

End Property

'Purpose: Creates a GDI brush object (HBRUSH) with the specified properties.

Public Function CreateBrush(ByVal clrFill As Long, _
                            Optional ByVal lStyle As UcsBrushStyle = BS_SOLID, _
                            Optional ByVal lHatch As UcsHatchStyles) As Long

    Dim lb     As LOGBRUSH

    On Error Resume Next
    With lb
        .lbStyle = lStyle
        .lbColor = TranslateColor(clrFill)
        .lbHatch = lHatch
    End With                                          'LB
    CreateBrush = ApiCreateBrushIndirect(lb)
    On Error GoTo 0

End Function

'Purpose: Creates a GDI pen object (HPEN) with the specified properties.

Public Function CreatePen(ByVal clrOutline As Long, _
                          Optional ByVal lStyle As UcsPenStyles = PS_SOLID, _
                          Optional ByVal lWidth As Long = 1) As Long

    On Error Resume Next
    CreatePen = ApiCreatePen(lStyle, lWidth, TranslateColor(clrOutline))
    On Error GoTo 0

End Function

Public Function DetachBitmap(Optional hNewBitmap As Long) As Long

    If m_MemoryBitmap <> 0 Then
        DetachBitmap = SelectObject(m_MemoryDC, m_OrginalBitmap)
    End If
    If hNewBitmap <> 0 Then
        m_OrginalBitmap = SelectObject(m_MemoryDC, hNewBitmap)
    End If
    m_MemoryBitmap = hNewBitmap

End Function

'Purpose: Block-image-transfers bits from current device context into a specified destination device context effectively creating a disabled copy of current image.

Public Sub DisabledBlt(Optional ByVal hdcDest As Long, _
                       Optional ByVal xDest As Long, _
                       Optional ByVal yDest As Long, _
                       Optional ByVal nWidth As Long, _
                       Optional ByVal nHeight As Long, _
                       Optional ByVal xSrc As Long, _
                       Optional ByVal ySrc As Long, _
                       Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                       Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                       Optional ByVal clrShadow As OLE_COLOR = vb3DShadow)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- fix arguments
    If hdcDest = 0 Then
        hdcDest = m_ParentDC
    End If
    If nWidth = 0 Then
        nWidth = m_MemoryWidth
    End If
    If nHeight = 0 Then
        nHeight = m_MemoryHeight
    End If
    pvDisabledBlt hdcDest, xDest, yDest, nWidth, nHeight, m_MemoryDC, xSrc, ySrc, clrMask, clrHighlight, clrShadow
    On Error GoTo 0

End Sub

'Purpose: Draws an ellipse in a rectangle in the device context using a specified color or pen.

Public Sub DrawEllipse(Optional ByVal LeftX As Long, _
                       Optional ByVal TopY As Long, _
                       Optional ByVal RightX As Long = -1, _
                       Optional ByVal BottomY As Long = -1, _
                       Optional ByVal clrOutline As Long = -1, _
                       Optional ByVal hpnOutline As Long)

'//--   Dim lb              As LOGBRUSH
'//--   Dim hbrOrig         As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- create pen if neccessary
    If clrOutline <> -1 Then
        hpnOutline = ApiCreatePen(PS_SOLID, 1, TranslateColor(clrOutline))
    End If
    If hpnOutline <> 0 Then
        Pen = hpnOutline
    End If
    '//--- draw ellipse
    ApiArc m_MemoryDC, LeftX, TopY, IIf(RightX < LeftX, Width, RightX), IIf(BottomY < TopY, Height, BottomY), LeftX, TopY, LeftX, TopY
    '//--   lb.lbStyle = BS_HOLLOW
    '//--   hbrOrig = SelectObject(m_MemoryDC, ApiCreateBrushIndirect(lb))
    '//--   ApiEllipse m_MemoryDC, LeftX,
    '//--TopY,
    '//--IIf(RightX < LeftX, Width, RightX),
    '//--IIf(BottomY < TopY, Height, BottomY)
    '//--   Call ApiDeleteObject(SelectObject(m_MemoryDC, hbrOrig))
    On Error GoTo 0

End Sub

'Purpose: Draws a windows control with a specific location and style in the device context.

Public Sub DrawFrameControl(Optional ByVal LeftX As Long, _
                            Optional ByVal TopY As Long, _
                            Optional ByVal RightX As Long = -1, _
                            Optional ByVal BottomY As Long = -1, _
                            Optional ByVal Type_ As UcsDrawFrameControlType = DFC_BUTTON, _
                            Optional ByVal ucsStyle As UcsDrawFrameControlStyle = DFCS_BUTTONPUSH)

    Dim rc     As RECT

    With rc
        .Left = LeftX
        .Top = TopY
        .Right = IIf(RightX < LeftX, Width, RightX)
        .Bottom = IIf(BottomY < TopY, Height, BottomY)
    End With                                          'RC
    ApiDrawFrameControl m_MemoryDC, rc, Type_, ucsStyle

End Sub

'Purpose: Creates a <b>StdPicture</b> object that represents an icon containing current image with pixels of specified mask color being transparent.

Public Function ExtractIcon(ByVal clrMask As OLE_COLOR) As StdPicture

    Dim ii     As ICONINFO
    Dim hMaskDC As Long
    Dim hImgDC As Long
    Dim hImgBmp As Long
    Dim hMaskBmp As Long
    Dim hPrevMaskBmp As Long
    Dim hPrevImgBmp As Long
    Dim hIcon  As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Function                                 '---> Bottom
    End If
    '//--- alloc resources
    hMaskDC = CreateCompatibleDC(0)
    hImgDC = CreateCompatibleDC(m_MemoryDC)
    hMaskBmp = CreateCompatibleBitmap(hMaskDC, Width, Height)
    hImgBmp = CreateCompatibleBitmap(m_MemoryDC, Width, Height)
    '//--- select bitmaps
    hPrevMaskBmp = SelectObject(hMaskDC, hMaskBmp)
    hPrevImgBmp = SelectObject(hImgDC, hImgBmp)
    '//--- copy image
    ApiBitBlt hImgDC, 0, 0, Width, Height, m_MemoryDC, 0, 0, vbSrcCopy
    '//--- create mask
    SetBkColor hImgDC, clrMask
    SetTextColor hImgDC, vbWhite
    ApiBitBlt hMaskDC, 0, 0, Width, Height, hImgDC, 0, 0, vbSrcCopy
    '//--- do mask image
    SetBkColor hImgDC, vbBlack
    ApiBitBlt hImgDC, 0, 0, Width, Height, hMaskDC, 0, 0, vbSrcAnd
    '//--- deselect bitmaps
    SelectObject hMaskDC, hPrevMaskBmp
    SelectObject hImgDC, hPrevImgBmp
    '//--- create icon
    With ii
        .fIcon = 1
        .hbmColor = hImgBmp
        .hbmMask = hMaskBmp
    End With                                          'II
    hIcon = CreateIconIndirect(ii)
    '//--- cleanup
    APIDeleteObject hMaskBmp
    APIDeleteObject hImgBmp
    DeleteDC hMaskDC
    DeleteDC hImgDC
    '//--- return StdPicture
    Set ExtractIcon = IconToPicture(hIcon)
    On Error GoTo 0

End Function

'Purpose: Outputs a text in a rectangle in the device context using specified output flags.

Public Sub ExtTextOut(ByVal Text As String, _
                      ByVal LeftX As Long, _
                      ByVal TopY As Long, _
                      ByVal RightX As Long, _
                      ByVal BottomY As Long, _
                      Optional ByVal xOffset As Long, _
                      Optional ByVal yOffset As Long, _
                      Optional ByVal etoFlags As UcsExtTextOutStyles)

    Dim rc     As RECT

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    With rc
        .Left = LeftX
        .Top = TopY
        .Right = RightX
        .Bottom = BottomY
    End With                                          'RC
    ApiExtTextOut m_MemoryDC, xOffset, yOffset, etoFlags, rc, Text, Len(Text), 0
    On Error GoTo 0

End Sub

'Purpose: Draws a frame around a rectangle in the device context with a specified color or brush.

Public Sub FrameRect(ByVal LeftX As Long, _
                     ByVal TopY As Long, _
                     ByVal RightX As Long, _
                     ByVal BottomY As Long, _
                     Optional ByVal clrFill As Long = -1, _
                     Optional ByVal hbrFill As Long)

    Dim rc     As RECT

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- create brush if neccessary
    If clrFill <> -1 Then
        hbrFill = CreateSolidBrush(TranslateColor(clrFill))
    End If
    With rc
        .Left = LeftX
        .Top = TopY
        .Right = RightX
        .Bottom = BottomY
    End With                                          'RC
    ApiFrameRect m_MemoryDC, rc, hbrFill
    '//--- cleanup the brush (if neccessary)
    If clrFill <> -1 Then
        APIDeleteObject hbrFill
    End If
    On Error GoTo 0

End Sub

'Purpose: Block-image-transfers bits to current device context from a specified destination device context using a specified raster operation.

Public Sub LoadBlt(Optional ByVal hdcSrc As Long, _
                   Optional ByVal xSrc As Long, _
                   Optional ByVal ySrc As Long, _
                   Optional ByVal xDest As Long, _
                   Optional ByVal yDest As Long, _
                   Optional ByVal nWidth As Long, _
                   Optional ByVal nHeight As Long, _
                   Optional ByVal dwRop As RasterOpConstants = vbSrcCopy)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- fix arguments
    If hdcSrc = 0 Then
        hdcSrc = m_ParentDC
    End If
    If nWidth = 0 Then
        nWidth = m_MemoryWidth
    End If
    If nHeight = 0 Then
        nHeight = m_MemoryHeight
    End If
    ApiBitBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, hdcSrc, xSrc, ySrc, dwRop
    On Error GoTo 0

End Sub

Public Sub PaintBitmap(ByVal hBmp As Long, _
                       Optional ByVal xDest As Long, _
                       Optional ByVal yDest As Long, _
                       Optional ByVal nWidth As Long, _
                       Optional ByVal nHeight As Long, _
                       Optional ByVal xSrc As Long, _
                       Optional ByVal ySrc As Long, _
                       Optional ByVal dwRop As RasterOpConstants = vbSrcCopy, _
                       Optional ByVal clrMask As OLE_COLOR = -1)

'Dim RC As RECT

    Dim hdcPaint As Long
    Dim hbmOrig As Long

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- draw
    hdcPaint = CreateCompatibleDC(m_MemoryDC)
    hbmOrig = SelectObject(hdcPaint, hBmp)
    If clrMask = -1 Then
        ApiBitBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, hdcPaint, xSrc, ySrc, dwRop
    Else                                              'NOT CLRMASK...
        pvTransBlt m_MemoryDC, xDest, yDest, nWidth, nHeight, hdcPaint, xSrc, ySrc, clrMask
    End If
    SelectObject hdcPaint, hbmOrig
    DeleteDC hdcPaint
    On Error GoTo 0

End Sub

'Purpose: Sets a clipping rectangle for the device context. Pixels outside clipping rectangle are not modified by drawing operations.

Public Sub SetClipRect(ByVal LeftX As Long, _
                       Optional ByVal TopY As Long, _
                       Optional ByVal RightX As Long, _
                       Optional ByVal BottomY As Long)

    Dim hRgn   As Long

    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    If LeftX >= 0 Then
        NormalizeRect LeftX, TopY, RightX, BottomY
        hRgn = CreateRectRgn(LeftX, TopY, RightX, BottomY)
    End If
    SelectClipRgn m_MemoryDC, hRgn
    If hRgn <> 0 Then
        APIDeleteObject hRgn
    End If

End Sub

'Purpose: Paints a DIB section in a rectangle in the device context using specified raster operation.

Public Sub SetDIBits(ByVal xDest As Long, _
                     ByVal yDest As Long, _
                     ByVal nWidth As Long, _
                     ByVal nHeight As Long, _
                     aBits() As Byte, _
                     Optional ByVal dwRop As RasterOpConstants = vbSrcCopy)

    Dim bmi    As BITMAPINFO

    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = nWidth
        .biHeight = nHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
    End With                                          'BMI.BMIHEADER
    StretchDIBits m_MemoryDC, xDest, yDest, nWidth, nHeight, 0, 0, nWidth, nHeight, aBits(0), bmi, DIB_RGB_COLORS, dwRop

End Sub

'Purpose: Block-image-transfers bits from current device context into a specified destination device context, stretching image and using a specified raster operation.

Public Sub StretchBlt(Optional ByVal hdcDest As Long, _
                      Optional ByVal xDest As Long, _
                      Optional ByVal yDest As Long, _
                      Optional ByVal nWidth As Long, _
                      Optional ByVal nHeight As Long, _
                      Optional ByVal xSrc As Long, _
                      Optional ByVal ySrc As Long, _
                      Optional ByVal nSrcWidth As Long, _
                      Optional ByVal nSrcHeight As Long, _
                      Optional ByVal dwRop As RasterOpConstants = vbSrcCopy)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- fix arguments
    If hdcDest = 0 Then
        hdcDest = m_ParentDC
    End If
    If nWidth = 0 Then
        nWidth = m_MemoryWidth
    End If
    If nHeight = 0 Then
        nHeight = m_MemoryHeight
    End If
    If nSrcWidth = 0 Then
        nSrcWidth = m_MemoryWidth
    End If
    If nSrcHeight = 0 Then
        nSrcHeight = m_MemoryHeight
    End If
    ApiStretchBlt hdcDest, xDest, yDest, nWidth, nHeight, m_MemoryDC, xSrc, ySrc, nSrcWidth, nSrcHeight, dwRop
    On Error GoTo 0

End Sub

'Purpose: Block-image-transfers bits from current device context into a specified destination device context skipping pixels with a specified transparent mask color.

Public Sub TransBlt(Optional ByVal hdcDest As Long, _
                    Optional ByVal xDest As Long, _
                    Optional ByVal yDest As Long, _
                    Optional ByVal nWidth As Long, _
                    Optional ByVal nHeight As Long, _
                    Optional ByVal xSrc As Long, _
                    Optional ByVal ySrc As Long, _
                    Optional ByVal clrMask As OLE_COLOR = MASK_COLOR)

    On Error Resume Next
    '//--- state check
    If Not IsCreated() Then
        Exit Sub                                      '---> Bottom
    End If
    '//--- fix arguments
    If hdcDest = 0 Then
        hdcDest = m_ParentDC
    End If
    If nWidth = 0 Then
        nWidth = m_MemoryWidth
    End If
    If nHeight = 0 Then
        nHeight = m_MemoryHeight
    End If
    pvTransBlt hdcDest, xDest, yDest, nWidth, nHeight, m_MemoryDC, xSrc, ySrc, TranslateColor(clrMask)
    On Error GoTo 0

End Sub

Public Function PaintIconGrayscale(ByVal Dest_hDC As Long, _
                                   ByVal hIcon As Long, _
                                   Optional ByVal Dest_X As Long, _
                                   Optional ByVal Dest_Y As Long, _
                                   Optional ByVal Dest_Height As Long, _
                                   Optional ByVal Dest_Width As Long) As Boolean

    Dim hBMP_Mask As Long
    Dim hBMP_Image As Long
    Dim hBMP_Prev As Long
    Dim hIcon_Temp As Long
    Dim hDC_Temp As Long

    ' Make sure parameters passed are valid

    If Dest_hDC = 0 Or hIcon = 0 Then Exit Function

    ' Extract the bitmaps from the icon
    If pvGetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function

    ' Create a memory DC to work with
    hDC_Temp = CreateCompatibleDC(0)

    If hDC_Temp = 0 Then GoTo CleanUp

    ' Make the image bitmap gradient
    If pvRenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0) = False Then GoTo CleanUp

    ' Extract the gradient bitmap out of the DC
    SelectObject hDC_Temp, hBMP_Prev

    ' Take the newly gradient bitmap and make a gradient icon from it
    hIcon_Temp = pvCreateIconFromBMP(hBMP_Mask, hBMP_Image)
    If hIcon_Temp = 0 Then GoTo CleanUp

    ' Draw the newly created gradient icon onto the specified DC
    If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0 Then
        PaintIconGrayscale = True
    End If

CleanUp:

    DestroyIcon hIcon_Temp
    hIcon_Temp = 0
    DeleteDC hDC_Temp
    hDC_Temp = 0
    APIDeleteObject hBMP_Mask
    hBMP_Mask = 0
    APIDeleteObject hBMP_Image
    hBMP_Image = 0

End Function

Private Function pvGetIconBitmaps(ByVal hIcon As Long, _
                                  ByRef Return_hBmpMask As Long, _
                                  ByRef Return_hBmpImage As Long) As Boolean

    Dim TempICONINFO As ICONINFO

    If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
    Return_hBmpMask = TempICONINFO.hbmMask
    Return_hBmpImage = TempICONINFO.hbmColor
    pvGetIconBitmaps = True

End Function

Private Function pvRenderBitmapGrayscale(ByVal Dest_hDC As Long, _
                                         ByVal hBitmap As Long, _
                                         Optional ByVal Dest_X As Long, _
                                         Optional ByVal Dest_Y As Long, _
                                         Optional ByVal Srce_X As Long, _
                                         Optional ByVal Srce_Y As Long _
                                         ) As Boolean

    Dim TempBITMAP As BITMAP
    Dim hScreen As Long
    Dim hDC_Temp As Long
    Dim hBMP_Prev As Long
    Dim MyCounterX As Long
    Dim MyCounterY As Long
    Dim NewColor As Long
    Dim hNewPicture As Long
    Dim DeletePic As Boolean

    ' Make sure parameters passed are valid

    If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function

    ' Get the handle to the screen DC
    hScreen = GetDC(0)
    If hScreen = 0 Then Exit Function

    ' Create a memory DC to work with the picture
    hDC_Temp = CreateCompatibleDC(hScreen)
    If hDC_Temp = 0 Then GoTo CleanUp

    ' If the user specifies NOT to alter the original, then make a copy of it to use
    DeletePic = False
    hNewPicture = hBitmap

    ' Select the bitmap into the DC
    hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)

    ' Get the height / width of the bitmap in pixels
    If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
    If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp

    ' Loop through each pixel and conver it to it's grayscale equivelant
    For MyCounterX = 0 To TempBITMAP.bmWidth - 1
        For MyCounterY = 0 To TempBITMAP.bmHeight - 1
            NewColor = ApiGetPixel(hDC_Temp, MyCounterX, MyCounterY)
            If NewColor <> -1 Then
                Select Case NewColor
                        ' If the color is already a grey shade, no need to convert it
                    Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
                        NewColor = NewColor

                    Case Else
                        NewColor = 0.33 * (NewColor Mod 256) + _
                                   0.59 * ((NewColor \ 256) Mod 256) + _
                                   0.11 * ((NewColor \ 65536) Mod 256)
                        NewColor = RGB(NewColor, NewColor, NewColor)

                End Select
                ApiSetPixel hDC_Temp, MyCounterX, MyCounterY, NewColor
            End If
        Next MyCounterY
    Next MyCounterX

    ' Display the picture on the specified hDC
    ApiBitBlt Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy

    pvRenderBitmapGrayscale = True

CleanUp:

    ReleaseDC 0, hScreen
    hScreen = 0
    SelectObject hDC_Temp, hBMP_Prev
    DeleteDC hDC_Temp
    hDC_Temp = 0
    If DeletePic Then
        APIDeleteObject hNewPicture
        hNewPicture = 0
    End If

End Function

Private Function pvCreateIconFromBMP(ByVal hBMP_Mask As Long, _
                                     ByVal hBMP_Image As Long) As Long

    Dim TempICONINFO As ICONINFO

    If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function

    TempICONINFO.fIcon = 1
    TempICONINFO.hbmMask = hBMP_Mask
    TempICONINFO.hbmColor = hBMP_Image

    pvCreateIconFromBMP = CreateIconIndirect(TempICONINFO)

End Function

