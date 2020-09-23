VERSION 5.00
Begin VB.UserControl JCToolbar 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   ControlContainer=   -1  'True
   PropertyPages   =   "JCToolbar.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ToolboxBitmap   =   "JCToolbar.ctx":0013
   Begin VB.PictureBox PicRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   6510
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox PicTB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   180
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   423
      TabIndex        =   2
      Top             =   0
      Width           =   6345
   End
   Begin VB.Timer TmrBtns 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2310
      Top             =   420
   End
   Begin VB.PictureBox PicLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      MousePointer    =   15  'Size All
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Timer tmrRight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1770
      Top             =   450
   End
End
Attribute VB_Name = "JCToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'=========================================================================
'
'   jcToolbars v 2.0.1
'   Copyright © 2005 Juan Carlos San Román Arias (sanroman2004@yahoo.com)
'
'=========================================================================

'   ------------------------------
'   Version 2.0 Data: 25-Dec-2005
'   ------------------------------
'   - It is just one control (it includes toolbar button (improved JCF_Toolbutton created by João Fortes) and vertical 3d line)
'   - add buttons and separators at runtime
'   - popup menu that shows hidden buttons and five theme colors
'   - auto hide toolbar buttons when toolbar is sizing
'   - initial toolbar autosizing (height and width) taking into account
'     icon and font size used in toolbar buttons.
'   - added two customs theme colors (norton 2004 and visual studio 2005 themes)
'   - It´s possible to change at runtime:
'     - button caption
'     - button state
'     - button tag
'     - button tooltiptext
'     - button value
'     - menu language selection (english and spanish)
'     - show or hide menu to change toolbar theme color
'   - improving of JCF_Toolbutton created by João Fortes
'     - icon size can be changed
'     - font can be changed (type, size, bold, italic)
'     - font color can be changed
'     - added four type of icon and caption aligments  color can be changed (IconLeftTextRight, IconRightTextLeft, IconTopTextBotton and IconBottonTextTop)

'   ------------------------------
'   Version 1.0  Data: 23-Nov-2005
'   ------------------------------
'   This is an Office 2003 toolbar for VB. You can built a nice Toolbar.
'   The initial idea taken from JCF_Toolbutton created by João Fortes.
'   I have made a compilation of different jobs published on Planet-Source-Code.com
'   I want to thank to
'   - Everyday Panos for your Office 2003 Button AND MOVING TOOLBAR project
'   - Fred cpp for api functions used in his isbutton control
'   - Carles P.V. for 3d UcVertical line
'   - All control is drawn using api functions (no images, no other controls)
'
'   ---------------------------------
'   Version 2.0.1  Data: 15-Feb-2006
'   ---------------------------------
'   - Control structure was completely reorganized (just one control)
'   - Chevrons have been added to show hidden buttons and 5 theme colors
'   - Buttons and separators can be added, moved or deleted at design and runtime
'   - Toolbar autosizing taking into account icon and font size used in toolbar buttons
'   - Windows XP theme auto detection or selection  (blue, silver and olive)
'   - Two customs theme colors (norton 2004 and visual studio 2005 themes) have been added
'   - You can select language for theme color Menu (spanish and english)
'   - You can determine if theme color Menu is shown
'   - A property page have been added for toolbar design
'
'=======================================================================================
'   I want specially thanks Jim Jose for his excellent McToolbar, I have used in my
'   jcToolbars some ideas from his usercontrol, such as chevrons and the way of loading
'   chevron picture as a separated window
'=======================================================================================

'=======================================================================================
'   There are still some unresolved problems (any help is wellcome):
'   - To modify used subclassing method to self subclassing in order to eliminate 2 class modules
'   - When button appears in picchevron (it is not visible) and you assign to this button
'     the function of unloading your program an error will occur.
'=======================================================================================


Option Explicit

'*************************************************************
'   Required Type Definitions
'*************************************************************
Private Type POINT
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type ToolbItem
    Caption As String
    Enabled As Boolean
    Key As String
    icon As StdPicture
    Iconsize As Integer
    BtnAlignment As AlignCont
    Tooltip As String
    BtnForeColor As OLE_COLOR
    Font As StdFont
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
    R_Height As Integer
    Type As jcBtnType
    State As jcBtnState
    Style As jcBtnStyle
    maskColor As OLE_COLOR
    UseMaskColor As Boolean
    Value As Boolean
End Type

Private Type TmpTBItem
    State As jcBtnState
    Value As Boolean
    icon As StdPicture
    Visible As Boolean
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

'types for cmnDialog
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

'xp theme
Public Enum ThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    Autodetect = 5
End Enum

'Menu language
Public Enum LangConst
    Spanish = 1
    English = 2
End Enum

'moving direction
Public Enum MoveConst
    ToLeft = 0
    ToRight = 1
End Enum

'Aligment icon and text
Public Enum AlignCont
    IconLeftTextRight = 0
    IconRightTextLeft = 1
    IconTopTextBottom = 2
    IconBottomTextTop = 3
End Enum

'type of toolbar item
Public Enum jcBtnType
    Button = 0
    Separator = 1
End Enum

'button style
Public Enum jcBtnStyle
    [Normal button] = 0
    [Check button] = 1
    [Dropdown button] = 2
End Enum

'state constants
Public Enum jcBtnState
    STA_NORMAL = 0
    STA_OVER = 1
    STA_PRESSED = 2
    STA_OVERDOWN = 3
    STA_SELECTED = 4
    STA_DISABLED = 5
End Enum

'button property
Public Enum jcBtnChangeProp
    jcCaption = 1
    jcEnabled = 2
    jckey = 3
    jcIcon = 4
    jcIconSize = 5
    jcTooltip = 6
    jcBtnForeColor = 7
    jcStyle = 8
    jcState = 9
    jcValue = 10
    jcFont = 11
    jcAlignment = 12
    jcType = 13
    jcUseMaskColor = 14
    jcMaskColor = 15
End Enum

'gradient type
Public Enum jcGradConst
    VerticalGradient = 0
    HorizontalGradient = 1
    VCilinderGradient = 2
    HCilinderGradient = 3
End Enum

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

'for bitmap conversion
Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

'for bitmap conversion
Private Type pictDesc
   cbSizeofStruct As Long
   picType As Long
   hImage As Long
End Type

'*************************************************************
'   Constants
'*************************************************************
Private Const ALTERNATE = 1      ' ALTERNATE and WINDING are
Private Const WINDING = 2        ' constants for FillMode.
Private Const BLACKBRUSH = 4     ' Constant for brush type.
Private Const WHITE_BRUSH = 0    ' Constant for brush type.
Private Const lWidth As Long = 24
Private Const lHeight As Long = 24
Private Const m_EmptyCaption As Integer = 16
Private Const m_MnuItems As Integer = 6
Private Const m_MnuItemHeight As Integer = 23
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

'constants for cmnDialog
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const REGULAR_FONTTYPE = &H400
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const PD_PRINTSETUP = &H40
Private Const PD_DISABLEPRINTTOFILE = &H80000

'state constants
Private Const RightBtn_NORMAL = 0
Private Const RightBtn_OVER = 1
Private Const RightBtn_PRESSED = 2
Private Const RightBtn_OVERDOWN = 3
Private Const m_OffSet = 4

'alignment constants
Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0
Private Const TEXT_INACTIVE = &H808080

'constants for subclassing
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_WINDOWPOSCHANGING  As Long = &H46
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_SIZE               As Long = &H5
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_MOVE               As Long = &H3
Private Const WM_TIMER              As Long = &H113
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MOUSEHOVER         As Long = &H2A1
Private Const WM_ACTIVATE           As Long = &H6
Private Const WM_ACTIVATEAPP        As Long = &H1C
Private Const WM_MOUSEACTIVATE      As Long = &H21
Private Const WM_CANCELMODE         As Long = &H1F
Private Const WM_NCACTIVATE         As Long = &H86
'*************************************************************
'events
'*************************************************************
Public Event ThemeColorClick(themeIndex As Integer)
Public Event ButtonClick(btnIndex As Integer, sKey As String, iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer, blnVisible As Boolean)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************
'members
'*************************************************************
Private m_IsStrech As Boolean
Private m_ThemeColor As ThemeConst
Private m_BtnAlignment As AlignCont
Private m_MenuLanguage As LangConst
Private m_ShowMenuColor As Boolean
Private m_ButtonCount As Long
Private ToolbarItem() As ToolbItem
Private TmpTBarItem() As TmpTBItem
Private m_BtnIndex As Integer
Private m_PrevBtnIndex As Integer
Private m_PrevBtn As Integer
Private m_TmpState As Integer
Private m_State As jcBtnState
Private m_StateAux As jcBtnState
Private R_Caption As RECT, R_Button As RECT
Private ChevronHeight As Integer
Private ChevronWidth As Integer
Private R_Chevron As RECT
Private m_MinWidth As Integer
Private m_MinHeight As Integer
Private m_MenuBtn() As String
Private m_MnuItemSelected As Integer
Private m_ChevronDropdown As Boolean
Private ColorFrom As OLE_COLOR, ColorTo As OLE_COLOR
Private ColorFromOver As OLE_COLOR, ColorToOver As OLE_COLOR
Private ColorFromDown As OLE_COLOR, ColorToDown As OLE_COLOR
Private ColorToolbar As OLE_COLOR, ColorBorderPic As OLE_COLOR
Private ColorToRight As OLE_COLOR, ColorFromRight As OLE_COLOR
Private ColorToRightPress As OLE_COLOR, ColorFromRightPress As OLE_COLOR
Private ColorToRightOver As OLE_COLOR, ColorFromRightOver As OLE_COLOR
Private ColorChevronOver As OLE_COLOR, ColorChevronPress As OLE_COLOR, ColorChevronSel As OLE_COLOR
Private useMask As Boolean
Private m_sCurrentSystemThemename As String 'Current Theme Name

Dim OFName As OPENFILENAME
Public mHwnd As Long
Dim CustomColors() As Byte
Public mFontName As String
Public mFontsize As Integer
Public mBold As Boolean
Public mItalic As Boolean
Public mUnderline As Boolean
Public mStrikethru As Boolean
Public mFontColor As Long

'*************************************************************
'   Required API Declarations
'*************************************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long

'api for cmndialog
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

'api for xp theme detection
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Private m_SubClassA As cSuperClass
Private m_SubClassB As cSuperClass
''Private m_SubClassC As cSuperClass
'
Implements ISuperClass
Private WithEvents picChevron As PictureBox
Attribute picChevron.VB_VarHelpID = -1

Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    If PropertyName = "BackColor" Then
        UserControl.BackColor = Ambient.BackColor
        UserControl_Resize
        DrawLeft
    End If
End Sub

'==========================================================================
' Init, Read & Write UserControl
'==========================================================================

Private Sub UserControl_Initialize()
    'Color selection according to window setup
    m_ThemeColor = Autodetect
    m_MenuLanguage = English
    m_ShowMenuColor = True
    Call SetThemeColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
    PicTB.BackColor = UserControl.BackColor
End Sub

Private Sub UserControl_InitProperties()
    UserControl.BackColor = Ambient.BackColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
    PicTB.BackColor = UserControl.BackColor
    'Extender.ToolTipText = "Toolbar Options"
    Calculate_Size
    InitialGradToolbar
    Width = 400
    m_MinHeight = 435
End Sub

Private Sub UserControl_Resize()
    Dim MinHeight As Long
    
    If UserControl.Width < 400 Then UserControl.Width = 400
'    If m_ButtonCount > 0 Then
'        If UserControl.Height < m_MinHeight Then UserControl.Height = m_MinHeight
'    Else
        MinHeight = MinimalHeight
'        If UserControl.Height < MinHeight Then
        UserControl.Height = MinHeight
'    End If
    
    PicLeft.Height = UserControl.ScaleHeight
    PicTB.Move PicLeft.Left + PicLeft.Width, 0, 1200, UserControl.ScaleHeight
    PicRight.Move UserControl.ScaleWidth - PicRight.Width, PicRight.Top, PicRight.Width, UserControl.ScaleHeight
    If m_ButtonCount > 0 Then
        Dim I As Integer
        For I = 1 To m_ButtonCount
            If ToolbarItem(I).Type = Button Then
                ToolbarItem(I).R_Height = PicTB.ScaleHeight - 2 * ToolbarItem(I).Top - 1
            Else
                ToolbarItem(I).R_Height = PicTB.ScaleHeight - 9
            End If
        Next I
    End If
    m_IsStrech = CheckIsVisible
    DrawRight ColorFromRight, ColorToRight
    If Not Ambient.UserMode Then
        Calculate_Size
        UserControl.Height = MinHeight
        InitialGradToolbar
        DrawTBtns
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim I As Integer
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    m_ThemeColor = PropBag.ReadProperty("ThemeColor", Autodetect)
    m_IsStrech = PropBag.ReadProperty("IsStrech", False)
    m_MenuLanguage = PropBag.ReadProperty("MenuLanguage", English)
    m_ShowMenuColor = PropBag.ReadProperty("ShowMenuColor", True)
    m_ButtonCount = PropBag.ReadProperty("ButtonCount", 0)
    PicRight.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    PicLeft.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    PicTB.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    
    'Load toolButtons
    ReDim ToolbarItem(m_ButtonCount)
    ReDim TmpTBarItem(m_ButtonCount)
    For I = 1 To m_ButtonCount
        With ToolbarItem(I)
            .Caption = PropBag.ReadProperty("BtnCaption" & I, Empty)
            .Enabled = PropBag.ReadProperty("BtnEnabled" & I, True)
            Set .icon = PropBag.ReadProperty("BtnIcon" & I, Nothing)
            ConvertToIcon I
            .Iconsize = PropBag.ReadProperty("BtnIconSize" & I, 16)
            .Tooltip = PropBag.ReadProperty("BtnToolTipText" & I, vbNullString)
            .Key = PropBag.ReadProperty("BtnKey" & I, Empty)
            .BtnAlignment = PropBag.ReadProperty("BtnAlignment" & I, IconLeftTextRight)
            .Type = PropBag.ReadProperty("BtnType" & I, 0)
            .Style = PropBag.ReadProperty("BtnStyle" & I, 0)
            .BtnForeColor = PropBag.ReadProperty("BtnForeColor" & I, &H80000008)
            Set .Font = PropBag.ReadProperty("BtnFont" & I, Ambient.Font)
            .State = PropBag.ReadProperty("BtnState" & I, 0)
            .Value = PropBag.ReadProperty("BtnValue" & I, False)
            .UseMaskColor = PropBag.ReadProperty("BtnUseMaskColor" & I, True)
            .maskColor = PropBag.ReadProperty("BtnMaskColor" & I, QBColor(13))
            TmpTBarItem(I).State = .State
            TmpTBarItem(I).Value = .Value
            TmpTBarItem(I).Visible = True
        End With
    Next I
    If Ambient.UserMode Then CreateChevron
    Call SetThemeColor
    PicTB.Move PicLeft.Left + PicLeft.Width, 0, 1200, UserControl.ScaleHeight
    InitialGradToolbar
    Calculate_Size
    DrawTBtns
    m_BtnIndex = -1
    m_ChevronDropdown = False
    
    Set m_SubClassA = New cSuperClass
    Set m_SubClassB = New cSuperClass
''    Set m_SubClassC = New cSuperClass
    pvSubClass
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim I As Long
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ThemeColor", m_ThemeColor, Autodetect)
    Call PropBag.WriteProperty("IsStrech", m_IsStrech, False)
    Call PropBag.WriteProperty("MenuLanguage", m_MenuLanguage, English)
    Call PropBag.WriteProperty("ShowMenuColor", m_ShowMenuColor, True)
    Call PropBag.WriteProperty("ButtonCount", m_ButtonCount, 0)
    
    'Load toolButtons
    For I = 1 To m_ButtonCount
        With ToolbarItem(I)
            Call PropBag.WriteProperty("BtnCaption" & I, .Caption, Empty)
            Call PropBag.WriteProperty("BtnEnabled" & I, .Enabled, True)
            Call PropBag.WriteProperty("BtnIcon" & I, .icon, Nothing)
            Call PropBag.WriteProperty("BtnIconSize" & I, .Iconsize, 16)
            Call PropBag.WriteProperty("BtnToolTipText" & I, .Tooltip, vbNullString)
            Call PropBag.WriteProperty("BtnKey" & I, .Key, Empty)
            Call PropBag.WriteProperty("BtnAlignment" & I, .BtnAlignment, IconLeftTextRight)
            Call PropBag.WriteProperty("BtnType" & I, .Type, 0)
            Call PropBag.WriteProperty("BtnStyle" & I, .Style, 0)
            Call PropBag.WriteProperty("BtnForeColor" & I, .BtnForeColor, &H80000008)
            Call PropBag.WriteProperty("BtnFont" & I, .Font, Ambient.Font)
            Call PropBag.WriteProperty("BtnState" & I, .State, 0)
            Call PropBag.WriteProperty("BtnValue" & I, .Value, False)
            Call PropBag.WriteProperty("BtnLeft" & I, .Left, 0)
            Call PropBag.WriteProperty("BtnTop" & I, .Top, 0)
            Call PropBag.WriteProperty("BtnWidth" & I, .Width, 0)
            Call PropBag.WriteProperty("BtnHeight" & I, .Height, 0)
            Call PropBag.WriteProperty("BtnUseMaskColor" & I, .UseMaskColor, True)
            Call PropBag.WriteProperty("BtnMaskColor" & I, .maskColor, QBColor(13))
        End With
    Next I
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Catch
    Set m_SubClassA = Nothing
    Set m_SubClassB = Nothing
'    Set m_SubClassC = Nothing
'    DestroyWindow picChevron.hWnd
    ' kill image used for transparencies when selected button pic is a bitmap
Catch:
End Sub

Private Sub picChevron_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As RECT
    If m_State = STA_PRESSED Then Exit Sub
    If Y > (ChevronHeight - m_MnuItemHeight) And Y < (ChevronHeight) And m_StateAux = STA_PRESSED Then Exit Sub
    picChevron.Cls
    PicTB.Cls
    m_BtnIndex = GetWhatButton(True, X, Y)
    If m_BtnIndex > -1 Then
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
            If TmpTBarItem(m_BtnIndex).Value Then
                m_State = STA_OVERDOWN
            Else
                m_State = STA_OVER
            End If
        Else
            m_State = STA_OVER
        End If
        SetRect R_Chevron, TmpTBarItem(m_BtnIndex).Left, TmpTBarItem(m_BtnIndex).Top, TmpTBarItem(m_BtnIndex).Width, TmpTBarItem(m_BtnIndex).Height
        DrawBtnInChevron picChevron, m_State, R_Chevron, m_BtnIndex
        picChevron.ToolTipText = ToolbarItem(m_BtnIndex).Tooltip
    Else
        If m_ShowMenuColor = False Then Exit Sub
        picChevron.ToolTipText = ""
        If Y < (ChevronHeight) And m_StateAux = STA_PRESSED Then Exit Sub
        If Y > (ChevronHeight - m_MnuItemHeight) Then
            Dim I As Integer
            For I = 0 To m_MnuItems
                If Y > (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight And Y < (ChevronHeight - m_MnuItemHeight) + (I + 1) * m_MnuItemHeight Then
                    DrawMenu STA_OVER, I
                    m_MnuItemSelected = I
                    Exit For
                Else
                    m_MnuItemSelected = -1
                End If
            Next I
        End If
    End If
End Sub

'==========================================================================
'  ThemeColorClick
'==========================================================================
Private Sub picChevron_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As RECT
    If Button = vbLeftButton Then
        If m_BtnIndex <> -1 Then
            If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
            m_State = STA_PRESSED
            DrawBtnInChevron picChevron, m_State, R_Chevron, m_BtnIndex
            m_PrevBtn = Button
            m_StateAux = STA_NORMAL
        Else
            If m_ShowMenuColor = False Then Exit Sub
            If picChevron.Height > (ChevronHeight) * 15 Then
                If Y > (ChevronHeight) Then
                    picChevron.Visible = False
                    m_TmpState = RightBtn_NORMAL
                    DrawRight ColorFromRight, ColorToRight
                    m_ThemeColor = m_MnuItemSelected - 1
                    Call SetThemeColor
                    InitialGradToolbar
                    DrawTBtns
                    RaiseEvent ThemeColorClick(CInt(m_ThemeColor))
                End If
                picChevron.Height = picChevron.Height - (m_MnuItemHeight * m_MnuItems + 3) * 15
                m_ChevronDropdown = True
                m_StateAux = STA_NORMAL
                DrawMenu STA_NORMAL, 0
            Else
                picChevron.Height = picChevron.Height + (m_MnuItemHeight * m_MnuItems + 3) * 15
                picChevron.Refresh
                ApiRectangle picChevron.hdc, 0, ChevronHeight - 1, picChevron.ScaleWidth - 1, m_MnuItemHeight * m_MnuItems + 3, ColorToolbar 'ColorBorderPic
                Dim I As Integer
                SetRect r, 1, ChevronHeight + 1, m_MnuItemHeight + 2, m_MnuItemHeight * m_MnuItems '- 1
                DrawGradientInRectangle picChevron.hdc, ColorFrom, ColorTo, r, HorizontalGradient, False, vbBlack
                For I = 0 To m_MnuItems
                    DrawMenu STA_PRESSED, I
                Next I
                m_ChevronDropdown = False

            End If
            picChevron.Picture = picChevron.Image
        End If
    End If
End Sub

'==========================================================================
'  ButtonClick
'==========================================================================
Private Sub picChevron_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnDrop As Boolean
    blnDrop = False
    If Button = vbLeftButton Then
        If m_BtnIndex <> -1 Then
            If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
            If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
                If TmpTBarItem(m_BtnIndex).Value Then
                    m_State = STA_NORMAL
                Else
                    m_State = STA_SELECTED
                    blnDrop = True
                End If
                TmpTBarItem(m_BtnIndex).Value = Not TmpTBarItem(m_BtnIndex).Value
                TmpTBarItem(m_BtnIndex).State = m_State
                UpdateCheckValue picChevron, m_BtnIndex
            Else
                m_State = STA_NORMAL
            End If
            DrawBtnInChevron picChevron, m_State, R_Chevron, m_BtnIndex
            m_PrevBtn = Button
            If ToolbarItem(m_BtnIndex).Style = [Dropdown button] And blnDrop = True Then
                m_State = STA_NORMAL
                TmpTBarItem(m_BtnIndex).Value = Not TmpTBarItem(m_BtnIndex).Value
                TmpTBarItem(m_BtnIndex).State = m_State
                UpdateCheckValue PicTB, m_BtnIndex
                m_PrevBtnIndex = -1
            End If
            picChevron.Visible = False
            m_TmpState = RightBtn_NORMAL
            DrawRight ColorFromRight, ColorToRight
            RaiseEvent ButtonClick(m_BtnIndex, ToolbarItem(m_BtnIndex).Key, (UserControl.ScaleWidth) * 15, (TmpTBarItem(m_BtnIndex).Top) * 15, (TmpTBarItem(m_BtnIndex).Width) * 15, (TmpTBarItem(m_BtnIndex).Height) * 15, False)
        End If
    End If
End Sub

Private Sub PicTB_DblClick()
    'If lPrevButton = vbLeftButton Then
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        PicTB_MouseDown 1, 0, 1, 1
    'End If
End Sub

Private Sub PicTB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If m_State = STA_PRESSED Then Exit Sub
    m_BtnIndex = GetWhatButton(False, X, Y)
    TmrBtns.Enabled = CheckMouseOverPicTB
    If m_BtnIndex > -1 Then
        If TmpTBarItem(m_BtnIndex).Visible = False Then Exit Sub
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        If ToolbarItem(m_BtnIndex).Type = Button Then
            If m_PrevBtnIndex <> m_BtnIndex Then
                m_PrevBtnIndex = m_BtnIndex
                PicTB.Cls
                If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
                    If TmpTBarItem(m_BtnIndex).Value Then
                        m_State = STA_OVERDOWN
                    Else
                        m_State = STA_OVER
                    End If
                Else
                    m_State = STA_OVER
                End If
                SetRect R_Button, ToolbarItem(m_BtnIndex).Left, ToolbarItem(m_BtnIndex).Top, ToolbarItem(m_BtnIndex).Width, ToolbarItem(m_BtnIndex).R_Height
                DrawBtn PicTB, m_State, R_Button, m_BtnIndex
                PicTB.ToolTipText = ToolbarItem(m_BtnIndex).Tooltip
            Else
                PicTB.ToolTipText = ToolbarItem(m_BtnIndex).Tooltip
            End If
        End If
    Else
        PicTB.ToolTipText = ""
        m_PrevBtnIndex = m_BtnIndex
        PicTB.Cls
    End If
    
End Sub

Private Sub PicTB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And m_BtnIndex <> -1 Then
        If ToolbarItem(m_BtnIndex).Type = Separator Then Exit Sub
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        m_State = STA_PRESSED
        DrawBtn PicTB, m_State, R_Button, m_BtnIndex
        m_PrevBtn = Button
    End If
End Sub

'==========================================================================
'  ButtonClick
'==========================================================================
Private Sub PicTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnDrop As Boolean
    blnDrop = False
    If Button = vbLeftButton And m_BtnIndex <> -1 Then
        If ToolbarItem(m_BtnIndex).Type = Separator Then Exit Sub
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
            If TmpTBarItem(m_BtnIndex).Value Then
                m_State = STA_NORMAL
            Else
                m_State = STA_SELECTED
                blnDrop = True
            End If
            TmpTBarItem(m_BtnIndex).Value = Not TmpTBarItem(m_BtnIndex).Value
            TmpTBarItem(m_BtnIndex).State = m_State
            UpdateCheckValue PicTB, m_BtnIndex
            m_PrevBtnIndex = -1
        Else
            m_State = STA_OVER
            DrawBtn PicTB, m_State, R_Button, m_BtnIndex
        End If
        m_PrevBtn = Button
        RaiseEvent ButtonClick(m_BtnIndex, ToolbarItem(m_BtnIndex).Key, (ToolbarItem(m_BtnIndex).Left + PicTB.Left) * 15, (ToolbarItem(m_BtnIndex).Top) * 15, (ToolbarItem(m_BtnIndex).Width) * 15, (ToolbarItem(m_BtnIndex).R_Height) * 15, True)
    End If
End Sub

Private Sub TmrBtns_Timer()
    If CheckMouseOverPicTB Then
    Else
        PicTB.Cls
        PicTB.Refresh
        If UserControl.BackColor <> Ambient.BackColor Then
            UserControl.BackColor = Ambient.BackColor
        End If
    End If
End Sub

Private Sub tmrRight_Timer()
    If CheckMouseOver Then
        If m_TmpState = RightBtn_NORMAL Then
            m_TmpState = RightBtn_OVER
            DrawRight ColorFromRightOver, ColorToRightOver
        End If
    Else
        If m_TmpState = RightBtn_OVER Then
            m_TmpState = RightBtn_NORMAL
            DrawRight ColorFromRight, ColorToRight
        End If
    End If
'
End Sub

'==========================================================================
' MouseDown
'==========================================================================
Private Sub picleft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picleft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'==========================================================================
' MouseMove
'==========================================================================
Private Sub picleft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicRight.ToolTipText = ""
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PicRight_DblClick()
    PicRight_MouseDown 1, 0, 1, 1
End Sub

Private Sub PicRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngI As Integer, X1 As Integer, Y1 As Integer
    
    If Button = vbLeftButton And picChevron.Visible = False Then
        m_TmpState = RightBtn_NORMAL
        DrawRight ColorFromRight, ColorToRight
    End If
End Sub

Private Sub PicRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Exit Sub
    m_TmpState = RightBtn_PRESSED
    DrawRight ColorFromRightPress, ColorToRightPress
    If picChevron.Visible = True Then
        picChevron.Visible = False
        m_TmpState = RightBtn_NORMAL
        DrawRight ColorFromRight, ColorToRight
        m_StateAux = STA_NORMAL
    Else
        If m_ShowMenuColor = False Then
            If m_IsStrech Then
                ShowChevron
                m_StateAux = STA_NORMAL
            End If
        Else
            ShowChevron
            m_StateAux = STA_NORMAL
        End If
    End If
End Sub

Private Sub PicRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Exit Sub
    If tmrRight.Enabled = False Then tmrRight.Enabled = True
End Sub

'==========================================================================
' Properties
'==========================================================================
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PicRight.BackColor = New_BackColor
    PicLeft.BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawTBtns
End Property

Public Property Get ButtonCount() As Integer
    ButtonCount = m_ButtonCount
End Property

Public Property Get IsStrech() As Boolean
    IsStrech = m_IsStrech
End Property

Public Property Let IsStrech(ByVal New_Value As Boolean)
    m_IsStrech = New_Value
    DrawRight ColorFromRight, ColorToRight
    PropertyChanged "IsStrech"
End Property

Public Property Get ThemeColor() As ThemeConst
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As ThemeConst)
    If m_ThemeColor <> vData Then
        m_ThemeColor = vData
        Call SetThemeColor
        InitialGradToolbar
        DrawTBtns
        PropertyChanged "ThemeColor"
    End If
End Property

Public Property Get MenuLanguage() As LangConst
    MenuLanguage = m_MenuLanguage
End Property

Public Property Let MenuLanguage(ByVal New_Value As LangConst)
    m_MenuLanguage = New_Value
    PropertyChanged "MenuLanguage"
End Property

Public Property Get BtnCaption(ByVal Index As Integer) As String
    BtnCaption = ToolbarItem(Index).Caption
End Property

Public Property Let BtnCaption(ByVal Index As Integer, ByVal New_Value As String)
    If ToolbarItem(Index).Caption = New_Value Then Exit Property
    ToolbarItem(Index).Caption = New_Value
    ChangeBtnProperty jcCaption, Index, New_Value
    PropertyChanged "BtnCaption"
End Property

Public Property Get BtnEnabled(ByVal Index As Integer) As Boolean
    BtnEnabled = ToolbarItem(Index).Enabled
End Property

Public Property Let BtnEnabled(ByVal Index As Integer, ByVal New_Value As Boolean)
    If ToolbarItem(Index).Enabled = New_Value Then Exit Property
    ToolbarItem(Index).Enabled = New_Value
    ChangeBtnProperty jcEnabled, Index, New_Value
    PropertyChanged "BtnEnabled"
End Property

Public Property Get BtnIcon(ByVal Index As Integer) As StdPicture
    Set BtnIcon = ToolbarItem(Index).icon
End Property

Public Property Set BtnIcon(ByVal Index As Integer, ByVal New_Picture As StdPicture)
    Set ToolbarItem(Index).icon = New_Picture
    ChangeBtnProperty jcIcon, Index, New_Picture
    PropertyChanged "BtnIcon"
End Property

Public Property Get BtnIconSize(ByVal Index As Integer) As Integer
    BtnIconSize = ToolbarItem(Index).Iconsize
End Property

Public Property Let BtnIconSize(ByVal Index As Integer, ByVal New_Value As Integer)
    If ToolbarItem(Index).Iconsize = New_Value Then Exit Property
    ToolbarItem(Index).Iconsize = New_Value
    ChangeBtnProperty jcIconSize, Index, New_Value
    PropertyChanged "BtnIconSize"
End Property

Public Property Get BtnToolTipText(ByVal Index As Integer) As String
    BtnToolTipText = ToolbarItem(Index).Tooltip
End Property

Public Property Let BtnToolTipText(ByVal Index As Integer, ByVal New_Value As String)
    If ToolbarItem(Index).Tooltip = New_Value Then Exit Property
    ToolbarItem(Index).Tooltip = New_Value
    ChangeBtnProperty jcTooltip, Index, New_Value
    PropertyChanged "BtnToolTipText"
End Property

Public Property Get BtnKey(ByVal Index As Integer) As String
    BtnKey = ToolbarItem(Index).Key
End Property

Public Property Let BtnKey(ByVal Index As Integer, ByVal New_Value As String)
    If ToolbarItem(Index).Key = New_Value Then Exit Property
    ToolbarItem(Index).Key = New_Value
    ChangeBtnProperty jckey, Index, New_Value
    PropertyChanged "BtnKey"
End Property

Public Property Get BtnAlignment(ByVal Index As Integer) As AlignCont
    BtnAlignment = ToolbarItem(Index).BtnAlignment
End Property

Public Property Let BtnAlignment(ByVal Index As Integer, ByVal New_Value As AlignCont)
    If ToolbarItem(Index).BtnAlignment = New_Value Then Exit Property
    ToolbarItem(Index).BtnAlignment = New_Value
    ChangeBtnProperty jcAlignment, Index, New_Value
    PropertyChanged "BtnAlignment"
End Property

Public Property Get BtnType(ByVal Index As Integer) As jcBtnType
    BtnType = ToolbarItem(Index).Type
End Property

Public Property Let BtnType(ByVal Index As Integer, ByVal New_Value As jcBtnType)
    If ToolbarItem(Index).Type = New_Value Then Exit Property
    ToolbarItem(Index).Type = New_Value
    ChangeBtnProperty jcType, Index, New_Value
    PropertyChanged "BtnType"
End Property

Public Property Get BtnStyle(ByVal Index As Integer) As jcBtnStyle
    BtnStyle = ToolbarItem(Index).Style
End Property

Public Property Let BtnStyle(ByVal Index As Integer, ByVal New_Value As jcBtnStyle)
    If ToolbarItem(Index).Style = New_Value Then Exit Property
    ToolbarItem(Index).Style = New_Value
    ChangeBtnProperty jcStyle, Index, New_Value
    PropertyChanged "BtnStyle"
End Property

Public Property Get BtnForeColor(ByVal Index As Integer) As OLE_COLOR
    BtnForeColor = ToolbarItem(Index).BtnForeColor
End Property

Public Property Let BtnForeColor(ByVal Index As Integer, ByVal New_Value As OLE_COLOR)
    If ToolbarItem(Index).BtnForeColor = New_Value Then Exit Property
    ToolbarItem(Index).BtnForeColor = New_Value
    ChangeBtnProperty jcBtnForeColor, Index, New_Value
    PropertyChanged "BtnForeColor"
End Property

Public Property Get BtnFont(ByVal Index As Integer) As Font
    Set BtnFont = ToolbarItem(Index).Font
End Property

Public Property Let BtnFont(ByVal Index As Integer, ByVal New_Value As Font)
    If ToolbarItem(Index).Font.Name = New_Value.Name And _
    ToolbarItem(Index).Font.Bold = New_Value.Bold And _
    ToolbarItem(Index).Font.Italic = New_Value.Italic And _
    ToolbarItem(Index).Font.Size = New_Value.Size And _
    ToolbarItem(Index).Font.Strikethrough = New_Value.Strikethrough And _
    ToolbarItem(Index).Font.Underline = New_Value.Underline And _
    ToolbarItem(Index).Font.Weight = New_Value.Weight Then Exit Property
    Set ToolbarItem(Index).Font = New_Value
    ChangeBtnProperty jcFont, Index, New_Value
    PropertyChanged "BtnFont"
End Property

Public Property Get btnState(ByVal Index As Integer) As Integer
    btnState = ToolbarItem(Index).State
End Property

Public Property Let btnState(ByVal Index As Integer, ByVal New_Value As Integer)
    If ToolbarItem(Index).State = New_Value Then Exit Property
    ToolbarItem(Index).State = New_Value
    TmpTBarItem(Index).State = New_Value
    ChangeBtnProperty jcState, Index, New_Value
    PropertyChanged "BtnState"
    'UpdateCheckValue PicTB, Index
End Property

Public Property Get BtnValue(ByVal Index As Integer) As Boolean
    BtnValue = ToolbarItem(Index).Value
End Property

Public Property Let BtnValue(ByVal Index As Integer, ByVal New_Value As Boolean)
    If ToolbarItem(Index).Value = New_Value Then Exit Property
    ToolbarItem(Index).Value = New_Value
    TmpTBarItem(Index).Value = New_Value
    ChangeBtnProperty jcValue, Index, New_Value
    PropertyChanged "BtnValue"
    'UpdateCheckValue PicTB, Index
End Property

Public Property Get BtnLeft(ByVal Index As Integer) As Integer
    BtnLeft = ToolbarItem(Index).Left
End Property

Public Property Let BtnLeft(ByVal Index As Integer, ByVal New_Value As Integer)
    ToolbarItem(Index).Left = New_Value
    PropertyChanged "BtnLeft"
End Property

Public Property Get BtnTop(ByVal Index As Integer) As Integer
    BtnTop = ToolbarItem(Index).Top
End Property

Public Property Let BtnTop(ByVal Index As Integer, ByVal New_Value As Integer)
    ToolbarItem(Index).Top = New_Value
    PropertyChanged "BtnTop"
End Property

Public Property Get BtnWidth(ByVal Index As Integer) As Integer
    BtnWidth = ToolbarItem(Index).Width
End Property

Public Property Let BtnWidth(ByVal Index As Integer, ByVal New_Value As Integer)
    ToolbarItem(Index).Width = New_Value
    PropertyChanged "BtnWidth"
End Property

Public Property Get btnHeight(ByVal Index As Integer) As Integer
    btnHeight = ToolbarItem(Index).Height
End Property

Public Property Let btnHeight(ByVal Index As Integer, ByVal New_Value As Integer)
    ToolbarItem(Index).Height = New_Value
    PropertyChanged "BtnHeight"
End Property

Public Property Get btnRHeight(ByVal Index As Integer) As Integer
    btnRHeight = ToolbarItem(Index).R_Height
End Property

Public Property Let btnRHeight(ByVal Index As Integer, ByVal New_Value As Integer)
    ToolbarItem(Index).R_Height = New_Value
    PropertyChanged "BtnRHeight"
End Property

Public Property Get BtnUseMaskColor(ByVal Index As Integer) As Boolean
    BtnUseMaskColor = ToolbarItem(Index).UseMaskColor
End Property

Public Property Let BtnUseMaskColor(ByVal Index As Integer, ByVal New_Value As Boolean)
    If ToolbarItem(Index).UseMaskColor = New_Value Then Exit Property
    ToolbarItem(Index).UseMaskColor = New_Value
    ChangeBtnProperty jcUseMaskColor, Index, New_Value
    PropertyChanged "BtnUseMaskColor"
End Property

Public Property Get BtnMaskColor(ByVal Index As Integer) As OLE_COLOR
    BtnMaskColor = ToolbarItem(Index).maskColor
End Property

Public Property Let BtnMaskColor(ByVal Index As Integer, ByVal New_Value As OLE_COLOR)
    If ToolbarItem(Index).maskColor = New_Value Then Exit Property
    ToolbarItem(Index).maskColor = New_Value
    ChangeBtnProperty jcMaskColor, Index, New_Value
    PropertyChanged "BtnMaskColor"
End Property

Public Property Get ShowMenuColor() As Boolean
    ShowMenuColor = m_ShowMenuColor
End Property

Public Property Let ShowMenuColor(ByVal New_Value As Boolean)
    m_ShowMenuColor = New_Value
    PropertyChanged "ShowMenuColor"
End Property

'==========================================================================
' Functions
'==========================================================================
Private Function CheckMouseOver() As Boolean
    Dim pt As POINT
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = PicRight.hWnd)
    tmrRight.Enabled = CheckMouseOver
End Function

Private Function CheckMouseOverPicTB() As Boolean
    Dim pt As POINT
    GetCursorPos pt
    CheckMouseOverPicTB = (WindowFromPoint(pt.X, pt.Y) = PicTB.hWnd)
End Function

Private Sub DrawRight(FromColor As Long, ToColor As Long)
    Dim r As RECT, lcolor As Long
    Dim poly(1 To 3) As POINT, I As Integer
    
    SetRect r, 0, 0, 2, PicRight.Height
    DrawVGradientEx PicRight.hdc, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom
    
    SetRect r, 2, 0, PicRight.Width, PicRight.Height
    DrawVGradientEx PicRight.hdc, FromColor, ToColor, r.Left, r.Top, r.Right, r.Bottom
    
    lcolor = TranslateColor(Ambient.BackColor)
    SetPixel PicRight.hdc, PicRight.Width - 1, 0, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, PicRight.Height - 1, lcolor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), FromColor)
    SetPixel PicRight.hdc, PicRight.Width - 2, 0, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, 1, lcolor
    
    lcolor = BlendColors(TranslateColor(ColorTo), FromColor)
    SetPixel PicRight.hdc, 0, 0, lcolor
    SetPixel PicRight.hdc, 1, 1, lcolor
    SetPixel PicRight.hdc, 1, 0, FromColor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), ToColor)
    SetPixel PicRight.hdc, PicRight.Width - 2, PicRight.Height - 1, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, PicRight.Height - 2, lcolor
    
    lcolor = BlendColors(TranslateColor(ColorFrom), ToColor)
    SetPixel PicRight.hdc, 0, PicRight.Height - 1, lcolor
    SetPixel PicRight.hdc, 1, PicRight.Height - 2, lcolor
    SetPixel PicRight.hdc, 1, PicRight.Height - 1, ToColor
        
        'drawing big right arrow
        'white triangle
        poly(1).X = 5:  poly(1).Y = PicRight.Height - 7 ' 19
        poly(2).X = 5 + 4: poly(2).Y = PicRight.Height - 7 '19
        poly(3).X = 5 + 2: poly(3).Y = PicRight.Height - 7 + 2 '19 + 3
        DrawTriangle PicRight, vbWhite, WHITE_BRUSH, poly, 3
        
        'black triangle
        poly(1).X = 4:  poly(1).Y = PicRight.Height - 8 '18
        poly(2).X = 4 + 4: poly(2).Y = PicRight.Height - 8 '18
        poly(3).X = 4 + 2: poly(3).Y = PicRight.Height - 8 + 2 '18 + 3
        DrawTriangle PicRight, vbBlack, BLACKBRUSH, poly, 3
        
        'black line
        'SetRect R, 6, 15, 13, 15
        SetRect r, 4, PicRight.Height - 11, 9, PicRight.Height - 11
        APILineEx PicRight.hdc, r.Left, r.Top, r.Right, r.Bottom, vbBlack
        
        'white line
        'SetRect R, 7, 16, 14, 16
        SetRect r, 5, PicRight.Height - 10, 10, PicRight.Height - 10
        APILineEx PicRight.hdc, r.Left, r.Top, r.Right, r.Bottom, vbWhite
    
    If m_IsStrech Then
        'drawing small arrows
        For I = 0 To 1
            SetRect r, 4 + 4 * I, 5, 4 + 4 * I, 8
            APILineEx PicRight.hdc, r.Left, r.Top, r.Right, r.Bottom, vbBlack
            SetPixel PicRight.hdc, 5 + 4 * I, 6, vbBlack
            SetPixel PicRight.hdc, 5 + 4 * I, 7, vbWhite
            SetPixel PicRight.hdc, 6 + 4 * I, 7, vbWhite
            SetPixel PicRight.hdc, 5 + 4 * I, 8, vbWhite
        Next I
    End If
    
    PicRight.Refresh
End Sub

Private Sub DrawLeft()
    Dim r As RECT, lcolor As Long, I As Long, yTop As Long, NumRect As Integer
    
    SetRect r, 0, 0, PicLeft.Width, PicLeft.Height
    DrawVGradientEx PicLeft.hdc, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom

    SetRect r, 2, PicLeft.Height - 1, PicLeft.Width, PicLeft.Height - 1
    APILineEx PicLeft.hdc, r.Left, r.Top, r.Right, r.Bottom, ColorToolbar

    lcolor = TranslateColor(Ambient.BackColor)
    SetPixel PicLeft.hdc, 0, 0, lcolor
    SetPixel PicLeft.hdc, 0, PicRight.Height - 1, lcolor
    SetPixel PicLeft.hdc, 0, PicRight.Height - 2, lcolor
    SetPixel PicLeft.hdc, 1, PicRight.Height - 1, lcolor

    lcolor = BlendColors(vbWhite, ColorTo)
    SetPixel PicLeft.hdc, 1, 0, lcolor
    SetPixel PicLeft.hdc, 0, 1, lcolor

    lcolor = BlendColors(ColorBorderPic, ColorFrom)
    SetPixel PicLeft.hdc, 1, PicRight.Height - 3, lcolor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), ColorFrom)
    SetPixel PicLeft.hdc, 0, PicRight.Height - 3, lcolor
    SetPixel PicLeft.hdc, 1, PicRight.Height - 2, lcolor
    
    NumRect = (PicRight.ScaleHeight - PicRight.ScaleHeight * 0.4) / 2 / 2
    yTop = (PicRight.ScaleHeight - 4 * (NumRect - 1) - 1) / 2
    For I = 0 To NumRect - 1
        SetRect r, 5, yTop + 4 * I, 1, 1
        ApiRectangle PicLeft.hdc, r.Left, r.Top, r.Right, r.Bottom, vbWhite
        SetRect r, 4, (yTop - 1) + 4 * I, 1, 1
        ApiRectangle PicLeft.hdc, r.Left, r.Top, r.Right, r.Bottom, ColorToolbar 'ColorBorderPic
    Next I
    PicLeft.Refresh
End Sub

Private Sub SetThemeColor()
    
    Select Case m_ThemeColor
        Case Is = Autodetect
            Call GetGradientColor(UserControl.hWnd)
        Case Else
            SetDefaultThemeColor m_ThemeColor
    End Select
End Sub

Private Sub SetDefaultThemeColor(ThemeType As Long)
    
    ColorChevronOver = RGB(255, 238, 194)
    ColorChevronPress = RGB(255, 128, 62)
    ColorChevronSel = RGB(255, 192, 111)
    ColorFromOver = RGB(255, 207, 142)
    ColorToOver = RGB(255, 245, 206)
    ColorFromDown = RGB(254, 145, 78)
    ColorToDown = RGB(254, 211, 142)
    ColorFromRightOver = RGB(255, 244, 204)
    ColorToRightOver = RGB(255, 197, 125)
    ColorFromRightPress = RGB(255, 154, 87)
    ColorToRightPress = RGB(255, 212, 144)
    
    Select Case ThemeType
            Case 0 '"NormalColor"
                ColorFrom = RGB(129, 169, 226)
                ColorTo = RGB(221, 236, 254)
                ColorToolbar = RGB(59, 97, 156)
                ColorBorderPic = RGB(0, 45, 150)
                ColorFromRight = RGB(118, 167, 241)
                ColorToRight = RGB(0, 53, 145)
            Case 1 '"Metallic"
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
            Case 2 '"HomeStead"
                ColorFrom = RGB(181, 197, 143)
                ColorTo = RGB(247, 249, 225)
                ColorToolbar = RGB(96, 128, 88)
                ColorBorderPic = RGB(63, 93, 56)
                ColorFromRight = RGB(177, 195, 141)
                ColorToRight = RGB(96, 119, 107)
            Case 3 '"Visual2005"
                ColorFrom = RGB(194, 194, 171)
                ColorTo = RGB(248, 248, 242)
                ColorFromOver = RGB(191, 209, 239)
                ColorToOver = RGB(191, 209, 239)
                ColorFromDown = RGB(149, 179, 228)
                ColorToDown = RGB(149, 179, 228)
                ColorToolbar = RGB(145, 145, 115)
                ColorBorderPic = RGB(49, 106, 197)
                ColorFromRight = RGB(208, 208, 198)
                ColorToRight = RGB(148, 148, 119)
                ColorFromRightOver = RGB(191, 209, 239)
                ColorToRightOver = RGB(191, 209, 239)
                ColorFromRightPress = RGB(149, 179, 228)
                ColorToRightPress = RGB(149, 179, 228)
                ColorChevronOver = RGB(191, 209, 239)
                ColorChevronPress = RGB(149, 179, 228)
                ColorChevronSel = RGB(149, 179, 228)
            Case 4 '"Norton2004"
                ColorFrom = RGB(217, 172, 1)
                ColorTo = RGB(255, 239, 165)
                ColorToolbar = RGB(117, 91, 30)
                ColorBorderPic = RGB(117, 91, 30)
                ColorFromRight = RGB(245, 179, 90)
                ColorToRight = RGB(117, 91, 30)
            Case Else
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
        End Select
End Sub

Private Sub GetGradientColor(lhWnd As Long)

    GetThemeName lhWnd
    
    ColorChevronOver = RGB(255, 238, 194)
    ColorChevronPress = RGB(255, 128, 62)
    ColorChevronSel = RGB(255, 192, 111)
    ColorFromOver = RGB(255, 207, 142)
    ColorToOver = RGB(255, 245, 206)
    ColorFromDown = RGB(254, 145, 78)
    ColorToDown = RGB(254, 211, 142)
    ColorFromRightOver = RGB(255, 244, 204)
    ColorToRightOver = RGB(255, 197, 125)
    ColorFromRightPress = RGB(255, 154, 87)
    ColorToRightPress = RGB(255, 212, 144)
    
    If AppThemed Then   '/Check if themed.
        Select Case m_sCurrentSystemThemename
            Case "NormalColor"
                ColorFrom = RGB(129, 169, 226)
                ColorTo = RGB(221, 236, 254)
                ColorToolbar = RGB(59, 97, 156)
                ColorBorderPic = RGB(0, 45, 150)
                ColorFromRight = RGB(118, 167, 241)
                ColorToRight = RGB(0, 53, 145)
            Case "Metallic"
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
            Case "HomeStead"
                ColorFrom = RGB(181, 197, 143)
                ColorTo = RGB(247, 249, 225)
                ColorToolbar = RGB(96, 128, 88)
                ColorBorderPic = RGB(63, 93, 56)
                ColorFromRight = RGB(177, 195, 141)
                ColorToRight = RGB(96, 119, 107)
            Case Else
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
        End Select
    Else
        ColorFrom = RGB(153, 151, 180)
        ColorTo = RGB(244, 244, 251)
        ColorToolbar = RGB(124, 124, 148)
        ColorBorderPic = RGB(75, 75, 111)
        ColorFromRight = RGB(180, 179, 200)
        ColorToRight = RGB(118, 116, 146)
    End If
End Sub

Public Function AddButton(Optional ByVal m_Type As jcBtnType = Button, Optional ByVal m_Caption As String = Empty, Optional ByVal m_Key As String = Empty, Optional ByVal m_Icon As StdPicture = Nothing, Optional ByVal m_IconSize As Integer = 16, _
                          Optional ByVal m_BtnAlignment As AlignCont = IconLeftTextRight, Optional ByVal m_Tooltip As String = Empty, Optional ByVal m_ForeColor As OLE_COLOR = &H80000008, Optional ByVal m_Font As StdFont, _
                          Optional ByVal m_Style As jcBtnStyle = [Normal button], Optional ByVal m_State As jcBtnState = STA_NORMAL, Optional ByVal m_UseMaskColor As Boolean = True, Optional ByVal m_MaskColor As OLE_COLOR = vbMagenta) As Integer
    
    Dim I As Integer, ix As Integer, iy As Integer, iw As Integer, ih As Integer
  
    Dim X As New StdFont
    X.Size = 8
    X.Name = "MS Sans Serif"

    m_ButtonCount = m_ButtonCount + 1
    
    ReDim Preserve ToolbarItem(m_ButtonCount)
    ReDim Preserve TmpTBarItem(m_ButtonCount)
    
    ToolbarItem(m_ButtonCount).Type = m_Type
    
    If Not (m_Font Is Nothing) Then
        Set ToolbarItem(m_ButtonCount).Font = m_Font
    Else
        Set ToolbarItem(m_ButtonCount).Font = X
    End If
    
    ToolbarItem(m_ButtonCount).Caption = m_Caption
    ToolbarItem(m_ButtonCount).Key = m_Key
    
    If Not (m_Icon Is Nothing) Then
        Set ToolbarItem(m_ButtonCount).icon = m_Icon
        ConvertToIcon CInt(m_ButtonCount)
    Else
        Set ToolbarItem(m_ButtonCount).icon = Nothing
    End If
    
    ToolbarItem(m_ButtonCount).Iconsize = m_IconSize
    ToolbarItem(m_ButtonCount).BtnAlignment = m_BtnAlignment
    ToolbarItem(m_ButtonCount).Tooltip = m_Tooltip
    ToolbarItem(m_ButtonCount).BtnForeColor = m_ForeColor
    ToolbarItem(m_ButtonCount).Style = m_Style
    ToolbarItem(m_ButtonCount).State = m_State
    ToolbarItem(m_ButtonCount).UseMaskColor = m_UseMaskColor
    ToolbarItem(m_ButtonCount).maskColor = m_MaskColor

    AddButton = m_ButtonCount
    Calculate_Size
    If Ambient.UserMode Then
        'Height = MinimalHeight
        InitialGradToolbar
    End If
    Width = m_MinWidth
End Function

Private Sub DrawSeparator(ColorLine As OLE_COLOR, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer)
    APILineEx PicTB.hdc, CLng(lLeft) + 1, CLng(lTop), CLng(lLeft) + 1, CLng(lTop + lHeight), ColorLine
    APILineEx PicTB.hdc, CLng(lLeft) + 2, CLng(lTop) + 1, CLng(lLeft) + 2, CLng(lTop + lHeight) + 1, vbWhite 'BlendColors(ColorLine, vbWhite)
End Sub

Private Function CheckIsVisible() As Boolean
    CheckIsVisible = False
    Dim I As Integer, RA As RECT
    For I = m_ButtonCount To 1 Step -1
        If UserControl.ScaleWidth - PicLeft.Width - PicRight.Width < ToolbarItem(I).Left + ToolbarItem(I).Width Then
            TmpTBarItem(I).Visible = False
            SetRect RA, ToolbarItem(I).Left, 0, ToolbarItem(I).Width + 1, PicTB.ScaleHeight - 2
            DrawGradientInRectangle PicTB.hdc, ColorFrom, ColorTo, RA, VerticalGradient, False, vbBlack
            CheckIsVisible = True
        Else
            TmpTBarItem(I).Visible = True
            If ToolbarItem(I).Type = Button Then
                SetRect RA, ToolbarItem(I).Left, ToolbarItem(I).Top, ToolbarItem(I).Width, ToolbarItem(I).R_Height
                DrawBtn PicTB, TmpTBarItem(I).State, RA, I
            Else
                DrawSeparator ColorToolbar, ToolbarItem(I).Left, ToolbarItem(I).Top, ToolbarItem(I).Width, ToolbarItem(I).R_Height
            End If
        End If
    Next I
    PicTB.Refresh
    PicTB.Picture = PicTB.Image
End Function

Public Sub ChangeBtnProperty(intOption As jcBtnChangeProp, intI As Integer, NewValue As Variant)
    Select Case intOption
        Case jcAlignment
            ToolbarItem(intI).BtnAlignment = NewValue
            Calculate_Size 'intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcCaption
            ToolbarItem(intI).Caption = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            DrawTBtns True, CLng(intI)
        Case jcEnabled
            ToolbarItem(intI).Enabled = NewValue
            ToolbarItem(intI).State = STA_DISABLED
            TmpTBarItem(intI).State = STA_DISABLED
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jckey
            ToolbarItem(intI).Key = NewValue
        Case jcIcon
            Set ToolbarItem(intI).icon = NewValue
            ConvertToIcon intI
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcIconSize
            ToolbarItem(intI).Iconsize = NewValue
            Calculate_Size intI, True
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcTooltip
            ToolbarItem(intI).Tooltip = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcBtnForeColor
            ToolbarItem(intI).BtnForeColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcStyle
            ToolbarItem(intI).Style = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            DrawTBtns True, CLng(intI), m_ButtonCount
        Case jcState
            ToolbarItem(intI).State = NewValue
            TmpTBarItem(intI).State = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcValue
            If NewValue = True Then
                ToolbarItem(intI).State = STA_PRESSED
                TmpTBarItem(intI).State = STA_PRESSED
            Else
                ToolbarItem(intI).State = STA_NORMAL
                TmpTBarItem(intI).State = STA_NORMAL
            End If
            ToolbarItem(intI).Value = NewValue
            TmpTBarItem(intI).Value = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcFont
            Set ToolbarItem(intI).Font = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcType
            ToolbarItem(intI).Type = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcUseMaskColor
            ToolbarItem(intI).UseMaskColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcMaskColor
            ToolbarItem(intI).maskColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
    End Select
End Sub

Public Function DeleteButton(ByVal Index As Integer) As Integer
    'Deletes an existing button from the control.
    Dim I As Long, BtnWidth As Long

    If Index < 0 Or Index > m_ButtonCount Then
        DeleteButton = -1
        Exit Function
    End If
    
    If m_ButtonCount = 0 Then Exit Function
    BtnWidth = ToolbarItem(Index).Width + m_OffSet / 2
    
    For I = Index To m_ButtonCount - 1
        ToolbarItem(I) = ToolbarItem(I + 1)
        TmpTBarItem(I) = TmpTBarItem(I + 1)
        ToolbarItem(I).Left = ToolbarItem(I).Left - BtnWidth
    Next I
    
    m_ButtonCount = m_ButtonCount - 1
    ReDim Preserve ToolbarItem(m_ButtonCount)
    ReDim Preserve TmpTBarItem(m_ButtonCount)
    
    'initial toolbar width
    UserControl.Width = (ToolbarItem(m_ButtonCount).Left + ToolbarItem(m_ButtonCount).Width + m_OffSet / 2 + PicLeft.Width + PicRight.Width) * 15
    Calculate_Size
    Width = m_MinWidth
    Height = m_MinHeight
    InitialGradToolbar
    DrawTBtns True
    DeleteButton = m_ButtonCount
End Function

Public Function MoveButton(ByVal Index As Integer, lDirection As MoveConst) As Integer
    Dim ToolbarItem_Aux As ToolbItem, NewIndex As Integer
    Dim TmpTBarItem_Aux As TmpTBItem
    Dim I As Long, BtnWidth As Long, iFrom As Long, iTo As Long

    Select Case lDirection
        Case ToLeft
            If Index > 1 Then NewIndex = Index - 1 Else Exit Function
            BtnWidth = ToolbarItem(Index).Width + m_OffSet / 2
            ToolbarItem(Index).Left = ToolbarItem(NewIndex).Left
            ToolbarItem(NewIndex).Left = ToolbarItem(Index).Left + BtnWidth
            iFrom = NewIndex
            iTo = Index
        Case ToRight
            If Index < m_ButtonCount Then NewIndex = Index + 1 Else Exit Function
            BtnWidth = ToolbarItem(NewIndex).Width + m_OffSet / 2
            ToolbarItem(NewIndex).Left = ToolbarItem(Index).Left
            ToolbarItem(Index).Left = ToolbarItem(NewIndex).Left + BtnWidth
            iFrom = Index
            iTo = NewIndex
    End Select
    
    ToolbarItem_Aux = ToolbarItem(NewIndex)
    ToolbarItem(NewIndex) = ToolbarItem(Index)
    ToolbarItem(Index) = ToolbarItem_Aux
    
    TmpTBarItem_Aux = TmpTBarItem(NewIndex)
    TmpTBarItem(NewIndex) = TmpTBarItem(Index)
    TmpTBarItem(Index) = TmpTBarItem_Aux
    
    DrawTBtns True, iFrom, iTo
    MoveButton = NewIndex
End Function

Public Sub Calculate_Size(Optional j As Integer = 1, Optional blnSize As Boolean = False)
    Dim I As Integer, k As Integer, MaxHeight As Integer
    Dim ix As Long, iy As Long, iw As Long, ih As Long
    
    Dim X As New StdFont
    X.Size = 8
    X.Name = "MS Sans Serif"
    
    m_MinWidth = 400
    m_MinHeight = 435

    If m_ButtonCount = 0 Then Exit Sub
    
    For I = j To m_ButtonCount
        If I > 0 Then
            ix = ToolbarItem(I - 1).Left + ToolbarItem(I - 1).Width + m_OffSet / 2
        Else
            ix = ToolbarItem(I + 1).Left
        End If
        
        If ToolbarItem(I).Type = Button Then
            iy = 2
            
            If Not (ToolbarItem(I).Font Is Nothing) Then
                Set UserControl.Font = ToolbarItem(I).Font
            Else
                Set UserControl.Font = X
            End If
            
            If ToolbarItem(I).icon Is Nothing Then  'there is no icon
                If Not (ToolbarItem(I).Caption = "") Then   'there is no caption
                    iw = 2 * m_OffSet + TextWidth(ToolbarItem(I).Caption)
                    ih = TextHeight(ToolbarItem(I).Caption) + 2 * m_OffSet
                Else
                    iw = 2 * m_OffSet + m_EmptyCaption
                    ih = m_EmptyCaption + 2 * m_OffSet
                End If
            Else
                If Not (ToolbarItem(I).Caption = "") Then   'there is no caption
                    Select Case ToolbarItem(I).BtnAlignment
                        Case IconLeftTextRight, IconRightTextLeft
                            iw = 3 * m_OffSet + TextWidth(ToolbarItem(I).Caption) + ToolbarItem(I).Iconsize
                            If TextHeight(ToolbarItem(I).Caption) > ToolbarItem(I).Iconsize Then
                                ih = TextHeight(ToolbarItem(I).Caption) + 2 * m_OffSet
                            Else
                                ih = ToolbarItem(I).Iconsize + 2 * m_OffSet
                            End If
                        Case IconTopTextBottom, IconBottomTextTop
                            If TextWidth(ToolbarItem(I).Caption) > ToolbarItem(I).Iconsize Then
                                iw = TextWidth(ToolbarItem(I).Caption) + 2 * m_OffSet
                            Else
                                iw = ToolbarItem(I).Iconsize + 2 * m_OffSet
                            End If
                            ih = 3 * m_OffSet + TextHeight(ToolbarItem(I).Caption) + ToolbarItem(I).Iconsize + 1
                    End Select
                Else
                    iw = 2 * m_OffSet + ToolbarItem(I).Iconsize
                    ih = ToolbarItem(I).Iconsize + 2 * m_OffSet
                End If
            End If
            If ToolbarItem(I).Style = [Dropdown button] Then iw = iw + 13
        Else    'it is separator
            iy = 4
            iw = 2
            ih = PicTB.ScaleHeight - 9
            ToolbarItem(I).R_Height = ih
        End If
        
        ToolbarItem(I).Left = ix
        ToolbarItem(I).Top = iy
        ToolbarItem(I).Width = iw
        ToolbarItem(I).Height = ih
        If ToolbarItem(I).Type = Button Then
            ToolbarItem(I).R_Height = PicTB.ScaleHeight - 2 * ToolbarItem(I).Top - 1
        Else
            ToolbarItem(I).R_Height = ih
        End If
    Next I
    
    'initial toolbar width
    m_MinWidth = (ix + iw + m_OffSet / 2 + PicLeft.Width + PicRight.Width) * 15
    
    If blnSize Then
        Width = m_MinWidth
        UserControl_Resize
    End If

End Sub

Private Function GetWhatButton(InChevron As Boolean, ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim I As Integer
    For I = 0 To m_ButtonCount
        
        If InChevron Then   'find button in chevron
            If TmpTBarItem(I).Visible = False Then
                If X > TmpTBarItem(I).Left And X < TmpTBarItem(I).Left + TmpTBarItem(I).Width Then
                    If Y > TmpTBarItem(I).Top And Y < TmpTBarItem(I).Top + TmpTBarItem(I).Height Then
                        GetWhatButton = I
                        Exit Function
                    End If
                End If
            End If
        Else    'find button in picTB
            If TmpTBarItem(I).Visible = True And ToolbarItem(I).Type = Button Then
                If X > ToolbarItem(I).Left And X < ToolbarItem(I).Left + ToolbarItem(I).Width Then
                    If Y > ToolbarItem(I).Top And Y < ToolbarItem(I).Top + ToolbarItem(I).R_Height Then
                        GetWhatButton = I
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    GetWhatButton = -1
End Function

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)

    'Use the API LineTo for Fast Drawing
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, 0
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Private Function ApiRectangle(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional lcolor As OLE_COLOR = -1) As Long
    
    Dim hPen As Long, hPenOld As Long
    Dim r
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X, Y, 0
    LineTo hdc, X + w, Y
    LineTo hdc, X + w, Y + H
    LineTo hdc, X, Y + H
    LineTo hdc, X, Y
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Function

Private Sub DrawVGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)

    ''Draw a Vertical Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    For ni = 0 To Y2
        APILineEx lhdcEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
End Sub

Private Sub DrawGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional blnVertical = True)
    
    'Draw a Vertical or horizontal Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
        dR = (sR - eR) / Y2
        dG = (sG - eG) / Y2
        dB = (sB - eB) / Y2
        For ni = 1 To Y2 - 1
            APILineEx lhdcEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    Else
        dR = (sR - eR) / X2
        dG = (sG - eG) / X2
        dB = (sB - eB) / X2
        For ni = 1 To X2 - 1
            APILineEx lhdcEx, X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    End If
End Sub

Private Sub DrawGradBorderRect(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, r As RECT, Optional lcolor As OLE_COLOR = -1)
    'draw gradient rectangle with border
    DrawVGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top, r.Right, r.Bottom
    ApiRectangle lhdcEx, r.Left, r.Top, r.Right, r.Bottom, lcolor
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lcolor As Long) As Long

    If OleTranslateColor(lcolor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
    
End Function

Private Function DrawTriangle(Pic As Object, ColorFore As Long, BrushColor, poly() As POINT, NumCoords As Long)
    Dim hBrush As Long, hRgn As Long
  
    Pic.ForeColor = ColorFore
    ' Polygon function creates unfilled polygon on screen.
    Polygon Pic.hdc, poly(1), NumCoords
    ' Gets stock black brush.
    hBrush = GetStockObject(BrushColor) 'WHITE_BRUSH)
    ' Creates region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    ' If the creation of the region was successful then color.
    If hRgn Then FillRgn Pic.hdc, hRgn, hBrush
    DeleteObject hRgn

End Function

Private Sub DrawGradientInRectangle(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, r As RECT, GradientType As jcGradConst, Optional blnDrawBorder As Boolean = False, Optional lBorderColor As Long = vbBlack, Optional LightCenter As Double = 2.01)
    Select Case GradientType
        Case VerticalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top, r.Right + r.Left, r.Bottom, True
        Case HorizontalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top, r.Right, r.Bottom + r.Top, False
        Case VCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, r, True, LightCenter
        Case HCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, r, False, LightCenter
    End Select
    If blnDrawBorder Then ApiRectangle lhdcEx, r.Left, r.Top, r.Right, r.Bottom, lBorderColor
End Sub

Private Sub DrawGradCilinder(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, r As RECT, Optional ByVal blnVertical As Boolean = True, Optional ByVal LightCenter As Double = 2.01)
    If LightCenter <= 1# Then LightCenter = 1.01
    If blnVertical Then
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, r.Left, r.Top, r.Right + r.Left, r.Bottom / LightCenter, True
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top + r.Bottom / LightCenter - 1, r.Right + r.Left, (LightCenter - 1) * r.Bottom / LightCenter + 1, True
    Else
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, r.Left, r.Top, r.Right / LightCenter, r.Bottom + r.Top, False
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left + r.Right / LightCenter - 1, r.Top, (LightCenter - 1) * r.Right / LightCenter + 1, r.Bottom + r.Top, False
    End If
End Sub

Private Sub DrawCaption(Pic As PictureBox, strText As String, FntColor As Long, RCaption As RECT, MyFont As StdFont, Optional BlnCenter As Boolean = False)
    Dim textAligment As Long
    
    Pic.ForeColor = FntColor
    Set Pic.Font = MyFont
    
    'Set the rectangle's values
    If BlnCenter = False Then
        textAligment = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    Else
        textAligment = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    End If
    
    'Draw text in PicTB or picChevron
    DrawTextEx Pic.hdc, strText, Len(strText), RCaption, textAligment, ByVal 0&
End Sub

'Drawing buttons in picTB
Public Sub DrawTBtns(Optional blnClear As Boolean = False, Optional iFrom As Long = 1, Optional iTo As Long = -1)
    Dim I As Integer, R1 As RECT
    Dim r As RECT
    
    If m_ButtonCount = 0 Then Exit Sub
    DrawLeft
    DrawRight ColorFromRight, ColorToRight
    
    If iTo = -1 Then iTo = m_ButtonCount
    If iFrom < 1 Then iFrom = 1
    If iTo > m_ButtonCount Then iTo = m_ButtonCount
    
    'clearing picTB background
    If blnClear Then
        If iFrom = m_ButtonCount Then
            SetRect r, ToolbarItem(iFrom).Left, 0, PicTB.ScaleWidth - ToolbarItem(iFrom).Left - 2, PicTB.ScaleHeight - 2
        Else
            SetRect r, ToolbarItem(iFrom).Left - 1, 0, ToolbarItem(iTo).Left - ToolbarItem(iFrom).Left + ToolbarItem(iTo).Width + m_OffSet / 2 + 1, PicTB.ScaleHeight - 2
        End If
        If iFrom = 1 Then r.Left = 0
        DrawGradientInRectangle PicTB.hdc, ColorFrom, ColorTo, r, VerticalGradient, False, vbBlack
        PicTB.Refresh
    End If
    
    For I = iFrom To iTo
        If TmpTBarItem(I).Visible = True Then
            If ToolbarItem(I).Type = Button Then
                SetRect R1, ToolbarItem(I).Left, ToolbarItem(I).Top, ToolbarItem(I).Width, ToolbarItem(I).R_Height
                DrawBtn PicTB, TmpTBarItem(I).State, R1, I
            Else
                DrawSeparator ColorToolbar, ToolbarItem(I).Left, ToolbarItem(I).Top, ToolbarItem(I).Width, ToolbarItem(I).R_Height
            End If
        End If
    Next I
    PicTB.Picture = PicTB.Image
End Sub

'Drawing picChevron button
Private Sub DrawBtnInChevron(Pic As PictureBox, btnState As jcBtnState, RBTN As RECT, Index As Integer)
    Dim R2 As RECT, RA As RECT
    Dim xIcon As Integer, yIcon As Integer, m_IconSize As Integer
    m_IconSize = 16
    Select Case btnState
        Case STA_PRESSED, STA_OVERDOWN
            DrawGradientInRectangle picChevron.hdc, ColorChevronPress, ColorChevronPress, RBTN, VerticalGradient, True, vbBlack
        Case STA_OVER
            DrawGradientInRectangle picChevron.hdc, ColorChevronOver, ColorChevronOver, RBTN, VerticalGradient, True, vbBlack
        Case STA_SELECTED
            DrawGradientInRectangle picChevron.hdc, ColorChevronSel, ColorChevronSel, RBTN, VerticalGradient, True, vbBlack
    End Select
    
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        If Not (ToolbarItem(Index).Caption = "") Then
            xIcon = RBTN.Left + m_OffSet
            yIcon = RBTN.Top + (RBTN.Bottom - m_IconSize) / 2
            'define caption rectangle
            SetRect R_Caption, RBTN.Left + m_IconSize + 2 * m_OffSet, RBTN.Top + m_OffSet - 1, RBTN.Left + RBTN.Right - m_OffSet, RBTN.Top + RBTN.Bottom - m_OffSet
        Else    'only icon
            xIcon = RBTN.Left + (RBTN.Right - m_IconSize) / 2
            yIcon = RBTN.Top + (RBTN.Bottom - m_IconSize) / 2
        End If
    Else 'no icon
        SetRect R_Caption, RBTN.Left + m_OffSet, RBTN.Top + m_OffSet - 1, RBTN.Left + RBTN.Right - m_OffSet, RBTN.Top + RBTN.Bottom - m_OffSet
    End If
    
    If btnState = STA_DISABLED Then
        DrawCaption Pic, ToolbarItem(Index).Caption, TEXT_INACTIVE, R_Caption, picChevron.Font
    Else
        DrawCaption Pic, ToolbarItem(Index).Caption, vbBlack, R_Caption, picChevron.Font
    End If
    
    'Drawing icon picture
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        useMask = ToolbarItem(Index).UseMaskColor
        If btnState = STA_DISABLED Then
            TransBlt Pic.hdc, xIcon, yIcon, m_IconSize, m_IconSize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , True, False
        Else
            TransBlt Pic.hdc, xIcon, yIcon, m_IconSize, m_IconSize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , False, False
        End If

    End If
    'drawing arrow
    If ToolbarItem(Index).Style = [Dropdown button] Then
        Const arrowWidth = 7
        Const arrowHeight = 4
        DrawArrow Pic, RBTN.Left + RBTN.Right - arrowWidth - m_OffSet, RBTN.Top + (RBTN.Bottom + 2 - arrowHeight) / 2
    End If
    picChevron.Refresh
End Sub

'Drawing picTB button
Private Sub DrawBtn(Pic As PictureBox, btnState As jcBtnState, RBTN As RECT, Index As Integer)
    Dim RA As RECT
    Dim xIcon As Integer, yIcon As Integer

    Select Case btnState
        Case STA_PRESSED
                DrawGradientInRectangle Pic.hdc, ColorFromDown, ColorToDown, RBTN, VerticalGradient, True, vbBlack
        Case STA_OVERDOWN
            If ToolbarItem(Index).Style = [Dropdown button] Then
                DrawGradientInRectangle Pic.hdc, BlendColors(ColorTo, vbWhite), BlendColors(ColorFrom, vbWhite), RBTN, VerticalGradient, True, vbBlack
            ElseIf ToolbarItem(Index).Style = [Check button] Then
                DrawGradientInRectangle Pic.hdc, ColorToDown, ColorFromDown, RBTN, VerticalGradient, True, vbBlack
            End If
        Case STA_OVER
                DrawGradientInRectangle Pic.hdc, ColorFromOver, ColorToOver, RBTN, VerticalGradient, True, ColorBorderPic 'vbBlack
                TmpTBarItem(Index).Value = False
        Case STA_SELECTED
            If ToolbarItem(Index).Style = [Dropdown button] Then
                DrawGradientInRectangle Pic.hdc, ColorTo, ColorFrom, RBTN, VerticalGradient, True, vbBlack
            ElseIf ToolbarItem(Index).Style = [Check button] Then
                DrawGradientInRectangle Pic.hdc, ColorFromDown, ColorToDown, RBTN, VerticalGradient, True, vbBlack
            End If
        Case STA_NORMAL, STA_DISABLED
            If ToolbarItem(Index).Style = [Dropdown button] Or ToolbarItem(Index).Style = [Check button] Then
                SetRect RA, RBTN.Left, 0, RBTN.Right + 1, Pic.ScaleHeight - 2
                DrawGradientInRectangle Pic.hdc, ColorFrom, ColorTo, RA, VerticalGradient, False, vbBlack
                TmpTBarItem(Index).Value = False
            End If
    End Select
    Set_CaptionAndIcon_Rect Index, RBTN, R_Caption, xIcon, yIcon

    If btnState = STA_DISABLED Then
        DrawCaption Pic, ToolbarItem(Index).Caption, TEXT_INACTIVE, R_Caption, ToolbarItem(Index).Font
    Else
        DrawCaption Pic, ToolbarItem(Index).Caption, ToolbarItem(Index).BtnForeColor, R_Caption, ToolbarItem(Index).Font
    End If
    
    'Drawing icon picture
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        useMask = ToolbarItem(Index).UseMaskColor
        If btnState = STA_DISABLED Then
            TransBlt Pic.hdc, xIcon, yIcon, ToolbarItem(Index).Iconsize, ToolbarItem(Index).Iconsize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , True, False
        Else
            TransBlt Pic.hdc, xIcon, yIcon, ToolbarItem(Index).Iconsize, ToolbarItem(Index).Iconsize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , False, False
        End If
    End If
    
    'drawing arrow
    If ToolbarItem(Index).Style = [Dropdown button] Then
        Const arrowWidth = 7
        Const arrowHeight = 4
        DrawArrow Pic, RBTN.Left + RBTN.Right - arrowWidth - m_OffSet, RBTN.Top + (RBTN.Bottom + 2 - arrowHeight) / 2
    End If
    Pic.Refresh
End Sub

Private Sub UpdateCheckValue(Pic As PictureBox, Index As Integer)
    Dim R1 As RECT
    SetRect R1, ToolbarItem(Index).Left, ToolbarItem(Index).Top, ToolbarItem(Index).Width, ToolbarItem(Index).R_Height
    Pic.Cls
    Pic.Refresh
    DrawBtn Pic, TmpTBarItem(Index).State, R1, Index
    Pic.Refresh
    Pic.Picture = Pic.Image
End Sub

Private Sub DrawArrow(Pic As PictureBox, ix As Single, iy As Single)
    Dim poly(1 To 3) As POINT, I As Integer
    
    'drawing big right arrow
    'black triangle
    poly(1).X = ix:  poly(1).Y = iy
    poly(2).X = ix + 6: poly(2).Y = iy
    poly(3).X = ix + 3: poly(3).Y = iy + 3
    DrawTriangle Pic, vbBlack, BLACKBRUSH, poly, 3
End Sub

Private Sub Set_CaptionAndIcon_Rect(Index As Integer, R_Button As RECT, R_Caption As RECT, xIcon As Integer, yIcon As Integer)
    'drawing image
    Set PicTB.Font = ToolbarItem(Index).Font
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        If Not (ToolbarItem(Index).Caption = "") Then
            Select Case ToolbarItem(Index).BtnAlignment
                Case IconLeftTextRight
                    xIcon = R_Button.Left + m_OffSet
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + ToolbarItem(Index).Iconsize + 2 * m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconRightTextLeft
                    xIcon = R_Button.Left + 2 * m_OffSet + PicTB.TextWidth(ToolbarItem(Index).Caption)
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - 2 * m_OffSet - ToolbarItem(Index).Iconsize, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconTopTextBottom
                    xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) / 2
                    'yIcon = R_Button.Top + m_OffSet + 1
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + ToolbarItem(Index).Iconsize + 2 * m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconBottomTextTop
                    xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) / 2
                    'yIcon = R_Button.Top + 2 * m_OffSet + PicTB.TextHeight(ToolbarItem(Index).Caption) - 1
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2 + PicTB.TextHeight(ToolbarItem(Index).Caption) + m_OffSet
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2 + PicTB.TextHeight(ToolbarItem(Index).Caption)
              End Select
        Else    'only icon
            xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) / 2
            yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) / 2
        End If
    Else 'no icon
        SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
    End If
End Sub

Private Sub CreateChevron()
    ChevronWidth = lWidth * 6
    ChevronHeight = lWidth * 7
    Set picChevron = UserControl.Controls.Add("vb.PictureBox", "picChevron")
    
    With picChevron
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Appearance = 0
        .BorderStyle = 0
    End With
    
    ' Hide the chevron from taskbar
    Dim lstyle As Long
    lstyle = GetWindowLong(picChevron.hWnd, GWL_EXSTYLE)
    lstyle = lstyle Or WS_EX_TOOLWINDOW
    'apply the tool window extended style
    SetWindowLongA picChevron.hWnd, GWL_EXSTYLE, lstyle
    SetParent picChevron.hWnd, 0
    
End Sub

Private Sub ShowChevron()
    Dim Rct As RECT
    GetWindowRect UserControl.hWnd, Rct
    Calculate_Size_InChevron
    SetWindowPos picChevron.hWnd, HWND_TOPMOST, Rct.Right - ChevronWidth, Rct.Bottom + 1, ChevronWidth, ChevronHeight, SWP_SHOWWINDOW
    DrawTBtnsInChevron
End Sub

'Drawing picChevron buttons
Public Sub DrawTBtnsInChevron()
    Dim I As Integer, R1 As RECT, m_BtnIndexA As Integer, R2 As RECT
    picChevron.Cls
    Set picChevron.Picture = Nothing
    ApiRectangle picChevron.hdc, 0, 0, picChevron.ScaleWidth - 1, picChevron.ScaleHeight - 1, ColorToolbar ' ColorBorderPic
    
    For I = 1 To m_ButtonCount
        If TmpTBarItem(I).Visible = False And ToolbarItem(I).Type = Button Then
            SetRect R1, TmpTBarItem(I).Left, TmpTBarItem(I).Top, TmpTBarItem(I).Width, TmpTBarItem(I).Height
            DrawBtnInChevron picChevron, TmpTBarItem(I).State, R1, I
        End If
    Next I
    
    If m_ShowMenuColor = True Then
        ApiRectangle picChevron.hdc, 0, picChevron.ScaleHeight - m_MnuItemHeight, picChevron.ScaleWidth - 1, picChevron.ScaleHeight - 1, ColorToolbar ' ColorBorderPic
    Else
        ApiRectangle picChevron.hdc, 0, picChevron.ScaleHeight, picChevron.ScaleWidth - 1, picChevron.ScaleHeight - 1, ColorToolbar ' ColorBorderPic
    End If
    
    If m_ShowMenuColor = True Then DrawMenu STA_NORMAL, 0
    
    picChevron.Picture = picChevron.Image
End Sub

Private Sub Calculate_Size_InChevron()
    Dim I As Integer, NumRow As Integer, LastBtn As Integer, tmpLeft As Integer
    Dim btnHeight As Integer, MaxLenght As Integer
    Dim iw As Integer, ih As Integer, m_IconSize As Integer, m_OffSet1 As String
    Dim BtnsInChevron As Integer
    
    NumRow = 1
    ChevronWidth = lWidth * 7
    LastBtn = -1
    m_IconSize = 16
    m_OffSet1 = m_OffSet / 2
    btnHeight = m_IconSize + 2 * m_OffSet1
    MaxLenght = btnHeight * 2
    BtnsInChevron = 0
    
    For I = 1 To m_ButtonCount
        If TmpTBarItem(I).Visible = False And ToolbarItem(I).Type = Button Then
            BtnsInChevron = BtnsInChevron + 1
            If ToolbarItem(I).icon Is Nothing Then
                If Not (ToolbarItem(I).Caption = "") Then
                    iw = 2 * m_OffSet + picChevron.TextWidth(ToolbarItem(I).Caption)
                Else
                    iw = 16
                End If
            Else
                If Not (ToolbarItem(I).Caption = "") Then
                    iw = 3 * m_OffSet + picChevron.TextWidth(ToolbarItem(I).Caption) + m_IconSize
                Else
                    iw = 2 * m_OffSet + m_IconSize
                End If
            End If
            If ToolbarItem(I).Style = [Dropdown button] Then iw = iw + 13
            TmpTBarItem(I).Width = iw
            TmpTBarItem(I).Height = btnHeight

Begining:
            tmpLeft = 0
            If LastBtn > 0 Then
                tmpLeft = TmpTBarItem(LastBtn).Left + TmpTBarItem(LastBtn).Width + m_OffSet1
                If tmpLeft + TmpTBarItem(I).Width + m_OffSet1 > ChevronWidth Then
                    NumRow = NumRow + 1
                    LastBtn = -1
                    If tmpLeft > MaxLenght Then MaxLenght = tmpLeft + 1
                    GoTo Begining
                Else
                    TmpTBarItem(I).Left = tmpLeft
                    TmpTBarItem(I).Top = m_OffSet1 + (NumRow - 1) * (btnHeight + m_OffSet1)
                    If tmpLeft + TmpTBarItem(I).Width + m_OffSet1 > MaxLenght Then
                        MaxLenght = tmpLeft + TmpTBarItem(I).Width + m_OffSet1 + 1
                    End If
                    LastBtn = I
                End If
            Else
                tmpLeft = m_OffSet1 + TmpTBarItem(I).Width + m_OffSet1
                If tmpLeft > ChevronWidth Then ChevronWidth = tmpLeft + 1
                If tmpLeft > MaxLenght Then MaxLenght = tmpLeft + 1
                TmpTBarItem(I).Left = m_OffSet1
                TmpTBarItem(I).Top = m_OffSet1 + (NumRow - 1) * (btnHeight + m_OffSet1)
                LastBtn = I
            End If
            
        End If
        ChevronHeight = (btnHeight + m_OffSet1) * NumRow + m_OffSet1 + 1
        
        If m_ShowMenuColor = True Then ChevronHeight = ChevronHeight + m_MnuItemHeight
    Next I
    
    ChevronWidth = MaxLenght
    If ChevronWidth < 155 Then ChevronWidth = 155
    If m_ShowMenuColor = True Then
        If BtnsInChevron = 0 Then ChevronHeight = m_MnuItemHeight
        ReDim m_MenuBtn(0 To m_MnuItems)
        m_MenuBtn(0) = LoadResString(200 + m_MenuLanguage)
        m_MenuBtn(1) = LoadResString(210 + m_MenuLanguage)
        m_MenuBtn(2) = LoadResString(220 + m_MenuLanguage)
        m_MenuBtn(3) = LoadResString(230 + m_MenuLanguage)
        m_MenuBtn(4) = LoadResString(240 + m_MenuLanguage)
        m_MenuBtn(5) = LoadResString(250 + m_MenuLanguage)
        m_MenuBtn(6) = LoadResString(260 + m_MenuLanguage)
    End If
 End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)

    If DstW = 0 Or DstH = 0 Then Exit Sub
    
    Dim B As Long, H As Long, F As Long, I As Long, newW As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    Dim hOldOb As Long
    Dim SrcDC As Long, tObj As Long, ttt As Long

    SrcDC = CreateCompatibleDC(hdc)

    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim hBrush As Long
        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(TransColor) 'MaskColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
        DeleteObject hBrush
    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If

    If Not useMask Then TransColor = -1

    newW = DstW - 1

    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            I = F + B
            If GetNearestColor(hdc, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
                With Data1(I)
                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                        Else
                            Data1(I) = BrushRGB
                        End If
                    Else
                        If isGreyscale Then
                            gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                        Else
                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                            Else
                                Data1(I) = Data2(I)
                            End If
                        End If
                    End If
                End With
            End If
        Next B
    Next H

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC: DeleteDC Sr2DC
    DeleteObject tObj: DeleteDC SrcDC
End Sub

Public Sub InitCmnDlg(myhwnd As Long)
    'Required to use custom colors
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim I As Integer
    For I = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(I) = 0
    Next I
    'need a window handle to run the functions
    mHwnd = myhwnd
End Sub

Public Function ShowColor(hWndParent As Long, DefColor As Long) As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    
    cc.rgbResult = DefColor
    cc.lStructSize = Len(cc)
    cc.hwndOwner = hWndParent
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function

Public Function ShowOpen(Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String) As String
    If mInitDir = "" Then mInitDir = App.Path & "\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr(0) + "*.*" + Chr(0)
    If mTitle = "" Then mTitle = App.Title
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = mHwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = mFilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = mInitDir
    OFName.lpstrTitle = mTitle
    OFName.flags = mflags
    If GetOpenFileName(OFName) Then
        ShowOpen = StripTerminator(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function

'Public Function ShowFont(fntName As String, fntBold As Boolean, fntItalic As Boolean, fntUnderline As Boolean, fntSize As Integer, fntStrikethru As Boolean) As Boolean
Public Function ShowFont(MyFont As StdFont) As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim RetVal As Long
    mFontName = ""
    lfont.lfHeight = 0
    lfont.lfWidth = 0
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfWeight = FW_NORMAL
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = MyFont.Name & vbNullChar
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = mHwnd
    cf.hdc = Printer.hdc
    cf.lpLogFont = pMem
    cf.iPointSize = MyFont.Size * 10
    cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)
    cf.nFontType = REGULAR_FONTTYPE
    cf.nSizeMin = 8
    cf.nSizeMax = 72
    RetVal = CHOOSEFONT(cf)
    If RetVal <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        MyFont.Name = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        MyFont.Bold = False
        MyFont.Italic = False
        MyFont.Underline = False
        MyFont.Strikethrough = False
        MyFont.Size = cf.iPointSize / 10
        If lfont.lfItalic = 255 Then MyFont.Italic = True
        If lfont.lfUnderline = 255 Then MyFont.Underline = True
        If lfont.lfWeight = 700 Then MyFont.Bold = True
        If lfont.lfStrikeOut = 255 Then MyFont.Strikethrough = True
        'mFontColor = cf.rgbColors
    Else
        ShowFont = False
    End If
    RetVal = GlobalUnlock(hMem)
    RetVal = GlobalFree(hMem)
End Function

Private Function StripTerminator(ByVal strString As String) As String
'gets rid of anything not required returned by API calls
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'Determines If The Current Window is Themed
Private Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'Returns The current Windows Theme Name
Private Sub GetThemeName(lngHwnd As Long)
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long
    Dim lPtrColorName As Long

    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(lngHwnd, StrPtr("ExplorerBar"))
    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        GetCurrentThemeName lPtrThemeFile, 260, lPtrColorName, 260, 0, 0
        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If iPos > 1 Then
            sThemeFile = Left$(sThemeFile, iPos - 1)
        End If
        m_sCurrentSystemThemename = bColorName
        iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
        If iPos > 1 Then
            m_sCurrentSystemThemename = Left$(m_sCurrentSystemThemename, iPos - 1)
        End If
        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid$(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left$(sThemeFile, iPos)
                Exit For
            End If
        Next iPos
        sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If
    On Error GoTo 0
End Sub

Public Sub InitialGradToolbar()
    Dim r As RECT
    PicTB.Cls
    Set PicTB.Picture = Nothing
    DrawLeft
    SetRect r, 0, 0, PicTB.ScaleWidth, PicTB.ScaleHeight
    DrawVGradientEx PicTB.hdc, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom
    APILineEx PicTB.hdc, r.Left, r.Bottom - 1, r.Right, r.Bottom - 1, ColorToolbar
    PicTB.Refresh
    DrawRight ColorFromRight, ColorToRight
    PicTB.Picture = PicTB.Image
End Sub

Private Function MinimalHeight() As Long
    Dim j As Long
    MinimalHeight = 29
    For j = 1 To m_ButtonCount
        If ToolbarItem(j).Type = Button Then
            If ToolbarItem(j).Height + 5 > MinimalHeight Then MinimalHeight = ToolbarItem(j).Height + 5
        End If
    Next j
    MinimalHeight = MinimalHeight * 15
End Function

Private Sub DrawCheckMark(Pic As PictureBox, m_IconSizeXP As Integer, X As Integer, Y As Integer, blnover As Boolean)
    Dim I As Integer, j As Integer, Z As Integer, m_clrCheckFore As OLE_COLOR, r As RECT
    
    SetRect r, X + 3, Y + 1, m_MnuItemHeight - 3, m_MnuItemHeight - 3
    If blnover Then
        DrawGradientInRectangle picChevron.hdc, ColorChevronPress, ColorChevronPress, r, VerticalGradient, True, ColorBorderPic
    Else
        DrawGradientInRectangle picChevron.hdc, ColorChevronSel, ColorChevronSel, r, VerticalGradient, True, ColorBorderPic
    End If
    m_clrCheckFore = vbBlack
    I = X + m_IconSizeXP \ 2 - 2
    For j = 0 To 2
        APILineEx Pic.hdc, I + j, Y + m_IconSizeXP \ 2 - 3 + j + 1, I + j, Y + m_IconSizeXP \ 2 + j, m_clrCheckFore
    Next j
    I = X + m_IconSizeXP \ 2 - 0
    For j = 1 To 4
        APILineEx Pic.hdc, I + j, Y + m_IconSizeXP \ 2 - j, I + j, Y + m_IconSizeXP \ 2 - j + 2, m_clrCheckFore
    Next j
End Sub

Private Sub DrawMenu(MnuState As jcBtnState, I As Integer)
    Dim r As RECT
    Const arrowWidth = 7
    Const arrowHeight = 4
    
    Dim X As New StdFont
    X.Size = 8
    X.Name = "MS Sans Serif"
    
    Set picChevron.Font = X
    
    If I = 0 Then
        SetRect r, 1, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight + 1, picChevron.ScaleWidth - 3, m_MnuItemHeight - 3
    Else
        SetRect r, 2, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight, picChevron.ScaleWidth - 5, m_MnuItemHeight - 1
    End If
    Select Case MnuState
        Case STA_NORMAL
            DrawGradientInRectangle picChevron.hdc, vbWhite, vbWhite, r, VerticalGradient, True, vbWhite
        Case STA_OVER
            If I = 0 Then
                If picChevron.Height > (ChevronHeight) * 15 Then
                    DrawGradientInRectangle picChevron.hdc, ColorToDown, ColorFromDown, r, VerticalGradient, True, ColorFromDown
                    'DrawGradientInRectangle picChevron.hdc, ColorChevronPress, ColorChevronPress, r, VerticalGradient, True, ColorFromDown
                Else
                    DrawGradientInRectangle picChevron.hdc, ColorFromOver, ColorToOver, r, VerticalGradient, True, ColorFromOver
                    'DrawGradientInRectangle picChevron.hdc, ColorChevronOver, ColorChevronOver, r, VerticalGradient, True, ColorFromOver
                End If
            Else
                DrawGradientInRectangle picChevron.hdc, ColorFromOver, ColorToOver, r, VerticalGradient, True, ColorBorderPic
                'DrawGradientInRectangle picChevron.hdc, ColorChevronOver, ColorChevronOver, r, VerticalGradient, True, ColorBorderPic
            End If
        Case STA_PRESSED
            If I = 0 Then
                DrawGradientInRectangle picChevron.hdc, ColorFromDown, ColorToDown, r, VerticalGradient, True, ColorFromDown
            End If
    End Select
    
    If I = 0 Then
        SetRect r, m_OffSet, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight, picChevron.ScaleWidth - 2 * m_OffSet - 1, (ChevronHeight - m_MnuItemHeight) + (I + 1) * m_MnuItemHeight - 2
        DrawCaption picChevron, m_MenuBtn(0), vbBlack, r, picChevron.Font, True
        DrawArrow picChevron, picChevron.ScaleWidth - arrowWidth - 2 * m_OffSet, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight + (m_MnuItemHeight - arrowHeight * 0.9) / 2
    Else
        SetRect r, m_MnuItemHeight + 2 + 8, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight, picChevron.ScaleWidth - 13, (ChevronHeight - m_MnuItemHeight) + (I + 1) * m_MnuItemHeight
        If m_ThemeColor = I - 1 Then
            If MnuState = STA_OVER Then
                DrawCheckMark picChevron, m_MnuItemHeight + 1, 0, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight, True
            Else
                DrawCheckMark picChevron, m_MnuItemHeight + 1, 0, (ChevronHeight - m_MnuItemHeight) + (I) * m_MnuItemHeight, False
            End If
        End If
        
        DrawCaption picChevron, m_MenuBtn(I), vbBlack, r, picChevron.Font
    End If
End Sub

'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== SubClassing
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub ISuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'hold place
End Sub
Private Sub ISuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    Select Case uMsg
        Case WM_SIZE, WM_MOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_NCACTIVATE, WM_CANCELMODE, WM_MOUSEACTIVATE
            picChevron.Visible = False
            m_TmpState = RightBtn_NORMAL
            DrawRight ColorFromRight, ColorToRight
            PicRight.Refresh
        Case WM_MOUSEMOVE
            If m_BtnIndex > -1 Then
            picChevron.Visible = False
            m_TmpState = RightBtn_NORMAL
            DrawRight ColorFromRight, ColorToRight
            PicRight.Refresh
            End If
    End Select
End Sub

Private Sub pvSubClass()
    With m_SubClassA
        Call .AddMsg(WM_SIZE, True)
        Call .AddMsg(WM_MOVE, True)
        Call .AddMsg(WM_LBUTTONDOWN, True)
        Call .AddMsg(WM_RBUTTONDOWN, False)
        Call .AddMsg(WM_MOUSEACTIVATE, True)
        Call .AddMsg(WM_CANCELMODE, True)
        Call .AddMsg(WM_NCACTIVATE, True)
        Call .SubClass(UserControl.Parent.hWnd, Me)
    End With
    With m_SubClassB
        Call .AddMsg(WM_MOUSEMOVE, True)
        Call .AddMsg(WM_MOUSEACTIVATE, False)
        Call .SubClass(PicTB.hWnd, Me)
    End With
'    With m_SubClassC
'        Call .AddMsg(WM_MOUSEACTIVATE, True)
'        Call .SubClass(PicRight.hWnd, Me)
'    End With
End Sub

Private Function fncMakeIcon(frmDC As Long, hBMP As Long, ByVal MaskClr As Long) As Long
    ' where frmDC   (in)  DC of the callng window
    '       hBMP    (in)  handle to a bitmap
    '       MaskClr (in)  if = -1 : pixel(0,0)
    ' Return value is a handle to the icon
    '       ipic    (out) icon picture
    
    Dim Bitmapdata As BITMAP  ' bitmap dimension
    Dim iWidth As Long
    Dim iHeight As Long
    Dim SrcDC As Long         ' copy of incoming bitmap
    Dim hSrc As Long
    Dim oldSrcObj As Long
    Dim MonoDC As Long        ' Mono mask (XOR)
    Dim MonoBmp As Long
    Dim oldMonoObj As Long
    Dim InvertDC As Long      ' Imverted mask (AND)
    Dim InvertBmp As Long
    Dim oldInvertObj As Long
    '
    Dim cBkColor As Long
    Dim icoinfo As ICONINFO

    ' validate input
    If hBMP = 0 Then Exit Function
    
    ' get size of bitmap
    If GetObject(hBMP, Len(Bitmapdata), Bitmapdata) = 0 Then Exit Function
    
    With Bitmapdata
        iWidth = .bmWidth
        iHeight = .bmHeight
    End With
   
    ' create copy of original, we will use it for both masks
    SrcDC = CreateCompatibleDC(0&)
    oldSrcObj = SelectObject(SrcDC, hBMP)
    
    ' get transparecy color
    If MaskClr = -1 Then
        MaskClr = GetPixel(SrcDC, 0, 0)
    End If
   
    ' mono mask (XOR) ............................................
    
    ' create mono DC/Bitmap for mask (XOR mask)
    MonoDC = CreateCompatibleDC(0&)
    MonoBmp = CreateCompatibleBitmap(MonoDC, iWidth, iHeight)
    oldMonoObj = SelectObject(MonoDC, MonoBmp)
    ' Set background of source to the mask color
    cBkColor = GetBkColor(SrcDC)   ' preserve original
    SetBkColor SrcDC, MaskClr
    ' copy bitmap and make monoDC mask in the process
    BitBlt MonoDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
    ' restore original backcolor
    SetBkColor SrcDC, cBkColor
    ' inverted mask (AND) .................................................

    ' create DC/bitmap for inverted image (AND mask)
    InvertDC = CreateCompatibleDC(frmDC)
    InvertBmp = CreateCompatibleBitmap(frmDC, iWidth, iHeight)
    oldInvertObj = SelectObject(InvertDC, InvertBmp)
    ' copy bitmap into it
    BitBlt InvertDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
    
    ' Invert background of image to create AND Mask
    SetBkColor InvertDC, vbBlack
    SetTextColor InvertDC, vbWhite
    BitBlt InvertDC, 0, 0, iWidth, iHeight, MonoDC, 0, 0, vbSrcAnd
    
    ' cleanup copy of original
    SelectObject SrcDC, oldSrcObj
    DeleteDC SrcDC
    
    ' Release MonoBmp And InvertBMP
    SelectObject MonoDC, oldMonoObj
    SelectObject InvertDC, oldInvertObj

    With icoinfo
        .fIcon = True
        .xHotspot = 16            ' Doesn't matter here
        .yHotspot = 16
        .hbmMask = MonoBmp
        .hbmColor = InvertBmp
    End With
      
    ' create 'output'
    fncMakeIcon = CreateIconIndirect(icoinfo)
    
CleanUp:
    ' Clean up
    DeleteObject icoinfo.hbmMask
    DeleteObject icoinfo.hbmColor
    DeleteDC MonoDC
    DeleteDC InvertDC
End Function

Private Function fncConvertIconToPic(hIcon As Long) As IPicture
    ' where hIcon   (in)  icon handle
    ' Return value is an interface managing a picture object and its properties
    '          (can be used to set a picture property)

    Dim iGuid As Guid
    Dim pDesc As pictDesc
     
    On Error Resume Next
     '--- check argument
    If hIcon = 0 Then Exit Function
    ' init GUID
    With iGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    ' fill picture description type
    With pDesc
       .cbSizeofStruct = Len(pDesc)
       .picType = vbPicTypeIcon
       .hImage = hIcon
        End With
    OleCreatePictureIndirect pDesc, iGuid, 1, fncConvertIconToPic
End Function

Private Sub ConvertToIcon(Index As Integer)
    With ToolbarItem(Index)
        If Not (.icon Is Nothing) Then
            If .icon.Type = vbPicTypeBitmap Then
                Set TmpTBarItem(Index).icon = fncConvertIconToPic(fncMakeIcon(UserControl.hdc, .icon.Handle, -1))
            Else
                Set TmpTBarItem(Index).icon = .icon
            End If
        End If
    End With
End Sub

