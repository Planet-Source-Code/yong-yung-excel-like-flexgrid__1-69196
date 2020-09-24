VERSION 5.00
Begin VB.UserControl ctxHookMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ctxHookMenu.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctxHookMenu.ctx":0034
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "ctxHookMenu.ctx":0346
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "ctxHookMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'==============================================================================
' ctxHookMenu.ctl
'
'   Subclassing Thunk (SuperClass V2) Project Samples
'   Copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'   Office XP menus control
'
' Modifications:
'
' 2002-10-28    WQW     Initial implementation
' 2002-11-10    WQW     Major refactoring for NT 4.0 compatibility
' 2003-10-29    GWN     Add Property - DrawStyle: Default Or Custom Style
'                       Added Function: SetCustomAttributes
'
'                       SetCustomAttributes: Allows The User To Customize
'                       The Menu.
'
' 2003-10-30    GWN     Sorted Scroll Effect on Window Animation
' 2003-11-05    NR      Added WM_MENUCHAR function for NT 4.0 OS
'                       Alt+shortcut was not working in NT 4.0
' 2003-11-05    GWN     Added MenuDraw Style : Allows The User To
'                       Select The Icon Draw Style -Dimmed Or UnDimmed
'
'                       Added Event 'Highlight' : Fires An Event Return The
'                       Selected Menu Caption
'
'                       Added Event CustomDrawItemFont
'                       Allows The User To Adjust The Font Properties Of
'                       Of The Menu Item.
'                       (Note: This Can Only Be Used When CustomDraw Is
'                       On And UseSystemFont Is False)
' 2003-11-07    NR      Added AutoColumn property.
' 2003-11-17    WQW     Fixed problem with WM_MENUCHAR.
' 2003-11-20    NR      Fixed problem with up/down arrow keys.
'==============================================================================
Option Explicit

'-- Added By Gary Noble
'-- Highlight Event - Fires When The User Selects
'-- A Menu Item
Public Event Highlight(strMenuCaption As String)
'-- CustomDraw Event - Fires When The User Selects
'-- A Menu Item - Only Works When useSystem Font Is False
'-- And Custom Draw Is Enabled
Public Event CustomDrawItemFont(Font As StdFont, Caption As String, ForeColour As OLE_COLOR)
Public Event CustomDrawItemHoverFont(SelectedFont As StdFont, Caption As String, SelectedForeColour As OLE_COLOR, SelectedBackColour As OLE_COLOR, SelectedBorderColour As OLE_COLOR)

Implements ISubclassingSink
Private Const MODULE_NAME As String = "ctxHookMenu"

#Const WEAK_REF_CURRENTMENU = 1
#Const WEAK_REF_ME = 0 '--- don't turn it on - its GPF-ing!!!

'==============================================================================
' API
'==============================================================================
Private lSpecialLineOffset As Long

'--- window messages
Private Const WM_DESTROY                As Long = &H2
Private Const WM_ERASEBKGND             As Long = &H14
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_SHOWWINDOW             As Long = &H18
Private Const WM_DRAWITEM               As Long = &H2B
Private Const WM_MEASUREITEM            As Long = &H2C
Private Const WM_WINDOWPOSCHANGING      As Long = &H46
Private Const WM_NCDESTROY              As Long = &H82
Private Const WM_NCCALCSIZE             As Long = &H83
Private Const WM_NCPAINT                As Long = &H85
Private Const WM_INITMENUPOPUP          As Long = &H117
Private Const WM_MENUSELECT             As Long = &H11F
Private Const WM_ENTERMENULOOP          As Long = &H211
Private Const WM_EXITMENULOOP           As Long = &H212
Private Const WM_MDISETMENU             As Long = &H230
Private Const WM_MDIGETACTIVE           As Long = &H229
Private Const WM_PRINT                  As Long = &H317
Private Const WM_PRINTCLIENT            As Long = &H318
Private Const WM_MENUCHAR               As Long = &H120

'--- menu flag
Private Const MF_GRAYED                 As Long = &H1
Private Const MF_DISABLED               As Long = &H2
Private Const MF_CHECKED                As Long = &H8
Private Const MF_POPUP                  As Long = &H10
Private Const MF_HILITE                 As Long = &H80&
Private Const MF_SYSMENU                As Long = &H2000
Private Const MF_BYCOMMAND              As Long = &H0&
Private Const MF_BYPOSITION             As Long = &H400&
'--- menu item info mask
Private Const MIIM_ID                   As Long = &H2
Private Const MIIM_SUBMENU              As Long = &H4
Private Const MIIM_TYPE                 As Long = &H10
Private Const MIIM_DATA                 As Long = &H20
'#if(WINVER >= 0x0500)
Private Const MIIM_STRING           As Long = &H40
Private Const MIIM_BITMAP           As Long = &H80
Private Const MIIM_FTYPE            As Long = &H100
'#endif /* WINVER >= 0x0500 */
'--- menu item info type
Private Const MFT_STRING                As Long = 0
Private Const MFT_OWNERDRAW             As Long = &H100
Private Const MFT_SEPARATOR             As Long = &H800
Private Const MFT_RIGHTJUSTIFY          As Long = &H4000
'--- for ownerdrawn items
Private Const ODT_MENU                  As Long = 1
Private Const ODS_SELECTED              As Long = &H1
Private Const ODS_HOTLIGHT              As Long = &H40
Private Const ODA_SELECT                As Long = &H2
'--- for GetSystemMetrics
Private Const SM_CYSCREEN               As Long = 1
Private Const SM_CXDLGFRAME             As Long = 7
Private Const SM_CXFRAME                As Long = 32
Private Const SM_CXEDGE                 As Long = 45
'--- for SetWindowLong (window styles)
Private Const WS_BORDER                 As Long = &H800000
Private Const WS_VISIBLE                As Long = &H10000000
Private Const WS_EX_DLGMODALFRAME       As Long = &H1&
Private Const WS_EX_WINDOWEDGE          As Long = &H100
'--- for GetClassLong
Private Const GCL_STYLE                 As Long = (-26)
Private Const CS_DROPSHADOW             As Long = &H20000
'--- for SetWindowPos
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOZORDER              As Long = &H4
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_DRAWFRAME             As Long = &H20
Private Const SWP_FLAGS                 As Long = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_DRAWFRAME
'--- for SystemParametersInfo
Private Const SPI_GETHIGHCONTRAST       As Long = 66
Private Const SPI_GETFLATMENU           As Long = &H1022
'--- for HIGHCONTRAST struct
Private Const HCF_HIGHCONTRASTON        As Long = &H1
Private Const HCF_AVAILABLE             As Long = &H2
'--- for GetDeviceCaps
Private Const BITSPIXEL                 As Long = 12
Private Const PLANES                    As Long = 14
'--- for registry
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const KEY_QUERY_VALUE           As Long = &H1
'--- for GetSysColor
Private Const COLOR_MENUBAR             As Long = 30
'--- for GetVersionEx
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
Private Const VER_PLATFORM_WIN2000      As Long = 1280
Private Const VER_PLATFORM_WINXP        As Long = 1281
'--- for menu columns
Private Const MF_MENUBREAK              As Long = &H40
Private Const MNC_EXECUTE               As Long = 2

'-- for Custom Colours
Private Const m_def_UserSideBarColour           As Long = 0
Private Const m_def_UserTopMenuBackColour       As Long = 0
Private Const m_def_UserTopMenuSelectedColour   As Long = 0
Private Const m_def_UserTopMenuHotColour        As Long = 0
Private Const m_def_UserTopMenuHotBorderColour  As Long = 0
Private Const m_def_UserMenuBorderColour        As Long = 0
Private Const m_def_UserCheckBackColour         As Long = 0
Private Const m_def_UserCheckBorderColour       As Long = 0
Private Const m_def_UserGradientOne             As Long = 0
Private Const m_def_UserGradientTwo             As Long = 0
Private Const m_def_UserSelectedMenuBackColour  As Long = 0
Private Const m_def_UserSelectedMenuBorderColour As Long = 0
Private Const m_def_UserSelectedItemForeColour  As Long = 0
Private Const m_def_UserUseGradient             As Boolean = False
Private Const m_def_UserUseTopMenuGradient      As Boolean = False
Private m_blnAutoColumnTop                      As Boolean
Private m_blnPopupLeftMost                      As Boolean
Private m_blnPopupAbove                         As Boolean

Private Declare Function InvalidateRectAsAny Lib "user32" Alias _
                          "InvalidateRect" (ByVal hwnd As Long, lpRect As Any, _
                          ByVal bErase As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal lpPoint As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByVal lpPoint As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hrgn As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Enum UcsDrawStyle
    DS_NORMAL = 1               ' Draws Normal Icon
    DS_XP = 2                   ' Draws Faded Icon
End Enum

Public Enum UcsMenuStyle
    MS_Default = 1
    MS_Custom = 2
End Enum

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Private Type MENUITEMINFO
    cbSize              As Long
    fMask               As Long
    fType               As Long
    fState              As Long
    wID                 As Long
    hSubMenu            As Long
    hbmpChecked         As Long
    hbmpUnchecked       As Long
    dwItemData          As Long
    dwTypeData          As Long
    cch                 As Long
    hbmpItem            As Long
End Type

Private Type MEASUREITEMSTRUCT
    CtlType             As Long
    CtlID               As Long
    itemID              As Long
    itemWidth           As Long
    ItemHeight          As Long
    ItemData            As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType             As Long
    CtlID               As Long
    itemID              As Long
    itemAction          As Long
    itemState           As Long
    hwndItem            As Long
    hDC                 As Long
    rcItem              As RECT
    ItemData            As Long
End Type

Private Type WINDOWPOS
    hwnd                As Long
    hWndInsertAfter     As Long
    x                   As Long
    y                   As Long
    cx                  As Long
    cy                  As Long
    Flags               As Long
End Type

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128      '  Maintenance string for PSS usage
End Type

Private Type HIGHCONTRAST
    cbSize              As Long
    dwFlags             As Long
    lpszDefaultScheme   As Long
End Type

'==============================================================================
' Constants and member vars
'==============================================================================

Private Const MASK_COLOR            As Long = &HFF00FF
Private Const DEF_SELECTDISABLED    As Boolean = True
Private Const DEF_BITMAPSIZE        As Long = 16
Private Const DEF_USESYSTEMFONT     As Boolean = True
Private Const STR_CLIENT_CLASS      As String = "MDIClient"
Private Const SEPARATOR_HEIGHT      As Long = 2

'-- Added By Gary Noble
Private Const m_def_DisplayShadow As Boolean = True
Private Const m_def_DrawStyle     As Integer = UcsMenuStyle.MS_Default
Private Const m_def_MenuDrawStyle As Integer = UcsDrawStyle.DS_NORMAL

'-- Added by NR
Private Const m_def_AutoColumn  As Integer = 0

Private m_oSubclass             As cSubclassingThunk
Private m_oClientSubclass       As cSubclassingThunk
Private m_cMenuSubclass         As Collection
Private m_cMemDC                As Collection
Private m_cBmps                 As Collection
Private m_ptLast                As POINTAPI
Private m_hFormMenu             As Long
Private m_hFormHwnd             As Long
Private m_hParentHwnd           As Long
Private m_lEdgeWidth            As Long '--- usually 2 px
Private m_lFrameWidth           As Long '--- usually 3 px
Private m_clrSelMenuBorder      As OLE_COLOR
Private m_clrSelMenuBack        As OLE_COLOR
Private m_clrSelMenuFore        As OLE_COLOR
Private m_clrSelCheckBack       As OLE_COLOR
Private m_clrMenuBorder         As OLE_COLOR
Private m_clrMenuBack           As OLE_COLOR
Private m_clrMenuFore           As OLE_COLOR
Private m_clrCheckBack          As OLE_COLOR
Private m_clrCheckFore          As OLE_COLOR
Private m_clrDisabledMenuBorder As OLE_COLOR
Private m_clrDisabledMenuBack   As OLE_COLOR
Private m_clrDisabledMenuFore   As OLE_COLOR
Private m_clrMenuBarBack        As OLE_COLOR
Private m_clrMenuPopupBack      As OLE_COLOR

'-- Added By Gary Noble (Phantom Man - PSC)
'-- Custom Draw Menu Attributes

Private m_MenuDrawStyle         As UcsDrawStyle

'-- Flag To Take Care Of The Menu Scroll Effect
Private m_blnFirstMenuInitialize As Boolean

Private m_lMenuHeight           As Long
Private m_lTextHeight           As Long
Private m_hLastMenu             As Long
Private m_bSelectDisabled       As Boolean
Private m_lBitmapSize           As Long
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_bUseSystemFont        As Boolean
Private m_cMenuInfo             As Collection
Private m_lInitMenuMsg          As Long
Private m_bExpectingPopup       As Boolean
Private m_bConstrainedColors    As Boolean
Private m_hLastSelMenu          As Long
Private m_rcLastSelMenu         As RECT
Private m_bLastSelMenuRightAlign As Boolean

'-- Added By Gary Noble (Phantom Man - PSC)
Private m_DrawStyle             As UcsMenuStyle
Private m_UserSideBarColour     As OLE_COLOR
Private m_UserTopMenuBackColour As OLE_COLOR
Private m_UserTopMenuSelectedColour As OLE_COLOR
Private m_UserTopMenuHotColour  As OLE_COLOR
Private m_UserTopMenuHotBorderColour As OLE_COLOR
Private m_UserMenuBorderColour  As OLE_COLOR
Private m_UserCheckBackColour   As OLE_COLOR
Private m_UserCheckBorderColour As OLE_COLOR
Private m_UserGradientOne       As OLE_COLOR
Private m_UserGradientTwo       As OLE_COLOR
Private m_UserUseGradient       As Boolean
Private m_UserUseTopMenuGradient As Boolean
Private m_UserSelectedItemForeColour As OLE_COLOR
Private m_UserSelectedMenuBackColour As OLE_COLOR
Private m_UserSelectedMenuBorderColour As OLE_COLOR

Private m_DisplayShadow             As Boolean

'-- Added by NR
Private m_AutoColumn            As Integer

#If DebugMode Then
Private m_sDebugID          As String
#End If

Private Enum UcsInitMeniType
    ucsIniMenu = 0
    ucsIniMainMenu
    ucsIniExitMenuLoop
    ucsIniEnterMenuLoop
    ucsIniParentForm
    ucsIniMenuChar
End Enum
'Default Property Values:
Const m_def_RightToLeft = False
'Property Variables:
Dim m_RightToLeft As Boolean
''
Public Property Let AutoColumn(ByVal iAutoColumn As Integer)

    m_AutoColumn = iAutoColumn
    PropertyChanged

End Property

'==============================================================================
' Properties
'==============================================================================
Public Property Get AutoColumn() As Integer

    AutoColumn = m_AutoColumn

End Property '

Property Let BitmapSize(ByVal lValue As Long)

    m_lBitmapSize = lValue
    pvGetMeasures
    PropertyChanged

End Property

Property Get BitmapSize() As Long

    BitmapSize = m_lBitmapSize

End Property

Private Property Get DEF_FONT() As StdFont

    Set DEF_FONT = New StdFont
    DEF_FONT.Name = "Tahoma"
    DEF_FONT.Size = 8

End Property

Public Property Let DisplayShadow(ByVal New_DisplayShadow As Boolean)

    m_DisplayShadow = New_DisplayShadow
    PropertyChanged "DisplayShadow"

End Property

Public Property Get DisplayShadow() As Boolean

    DisplayShadow = m_DisplayShadow

End Property

'-- Added By Gary Noble (Phantom Man - PSC)- Custom Attributes
Public Property Get DrawStyle() As UcsMenuStyle
Attribute DrawStyle.VB_Description = "Sets The Draw Style Of The Menu"

    DrawStyle = m_DrawStyle

End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As UcsMenuStyle)

    m_DrawStyle = New_DrawStyle
    PropertyChanged "DrawStyle"
    DrawMenuBar IIf(m_hParentHwnd <> 0, m_hParentHwnd, m_hFormHwnd)

End Property

Property Get Font() As StdFont

    Set Font = m_oFont

End Property

Property Set Font(ByVal oSrc As StdFont)

    With m_oFont
        .Bold = oSrc.Bold
        .Charset = oSrc.Charset
        .Italic = oSrc.Italic
        .Name = oSrc.Name
        .Size = oSrc.Size
        .Strikethrough = oSrc.Strikethrough
        .Underline = oSrc.Underline
        .Weight = oSrc.Weight
    End With
    pvGetMeasures
    PropertyChanged

End Property

Friend Property Set frBmps(ByVal oValue As Collection)

    Set m_cBmps = oValue
    PropertyChanged

End Property

Friend Property Get frBmps() As Collection

  Dim vElem As Variant

    Set frBmps = New Collection
    For Each vElem In m_cBmps
        frBmps.Add vElem, vElem(2)
    Next vElem

End Property

Friend Property Get frContainerMenus() As Collection

  Dim oCtl            As Object

    On Error Resume Next
        Set frContainerMenus = New Collection
        For Each oCtl In ParentControls
            If TypeOf oCtl Is Menu Then
                frContainerMenus.Add oCtl
            End If
        Next oCtl
    On Error GoTo 0

End Property

Friend Sub frSubclassPopup(ByVal hwnd As Long)

  Dim oSubclass       As cSubclassingThunk
  Dim lStyle          As Long
  Dim lExStyle        As Long

    On Error Resume Next
        '--- check if this is a popup menu from main menubar
        If Not m_bExpectingPopup Then
            Exit Sub
        End If
        Set oSubclass = m_cMenuSubclass("#" & hwnd)
        If oSubclass Is Nothing Then
            Set oSubclass = New cSubclassingThunk
            With oSubclass
#If WEAK_REF_ME Then
                .Subclass hwnd, Me, True, True
#Else
                .Subclass hwnd, Me, False, True
#End If
                .AddBeforeMsgs WM_ERASEBKGND, WM_NCCALCSIZE, WM_NCPAINT, _
                               WM_WINDOWPOSCHANGING, WM_PRINT, WM_SHOWWINDOW, WM_DESTROY
            End With
            m_cMenuSubclass.Add oSubclass, "#" & hwnd
        End If
        '--- fix styles
        lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        lStyle = GetWindowLong(hwnd, GWL_STYLE)
        oSubclass.Tag = Array(lStyle, lExStyle)
        SetWindowLong hwnd, GWL_EXSTYLE, lExStyle And (Not WS_EX_DLGMODALFRAME) And (Not WS_EX_WINDOWEDGE)
        SetWindowLong hwnd, GWL_STYLE, lStyle And (Not WS_BORDER)
        lStyle = GetClassLong(hwnd, GCL_STYLE)
        '--- win98: check if anything to modify
        If (lStyle And CS_DROPSHADOW) <> 0 Then
            SetClassLong hwnd, GCL_STYLE, lStyle And (Not CS_DROPSHADOW)
        End If
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_FLAGS
    On Error GoTo 0

End Sub

'==============================================================================
' Methods
'==============================================================================

Public Sub Init(hwnd As Long)

  Dim hClient         As Long

    '--- member vars
    m_hFormHwnd = hwnd
    m_hFormMenu = GetMenu(m_hFormHwnd)
    Set m_oSubclass = New cSubclassingThunk
    Set m_oClientSubclass = New cSubclassingThunk
    '--- get appearance info and init menu
    pvGetMeasures
    '--- subclass form
    With m_oSubclass
#If WEAK_REF_ME Then
        .Subclass m_hFormHwnd, Me, True, True
#Else
        .Subclass m_hFormHwnd, Me, False, True
#End If
        .AddBeforeMsgs WM_INITMENUPOPUP, WM_MEASUREITEM, WM_DRAWITEM, _
                       WM_MDISETMENU, WM_NCCALCSIZE, WM_MENUSELECT, _
                       WM_SYSCOLORCHANGE, WM_NCDESTROY, pvInitMenuMsg, _
                       WM_ENTERMENULOOP, WM_EXITMENULOOP, WM_MENUCHAR

    End With
    '--- special case: subclass MDI client
    hClient = FindWindowEx(hwnd, 0, STR_CLIENT_CLASS, vbNullString)
    If hClient <> 0 Then
        With m_oClientSubclass
#If WEAK_REF_ME Then
            .Subclass hClient, Me, True, True
#Else
            .Subclass hClient, Me, False, True
#End If
            .AddBeforeMsgs WM_MDISETMENU
        End With
    End If

End Sub

Private Property Get IsNT() As Boolean

  Static lPlatform    As Long
  Dim uVer            As OSVERSIONINFO

    If lPlatform = 0 Then
        uVer.dwOSVersionInfoSize = Len(uVer)
        If GetVersionEx(uVer) Then
            lPlatform = uVer.dwPlatformId
        End If
    End If
    IsNT = (lPlatform = VER_PLATFORM_WIN32_NT)

End Property

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub

'==============================================================================
' ISubclassingSink interface
'==============================================================================

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)

  Static bDelayed  As Boolean
  Static lsHwnd       As Long
  Dim hDC             As Long
  Dim rc              As RECT
  Dim wp              As WINDOWPOS
  Dim pt              As POINTAPI
  Dim oSub            As cSubclassingThunk
  Dim hMdiChild       As Long
  Dim hPrevMenuWnd    As Long
  Dim mii             As MENUITEMINFO
  Dim sBuffer         As String
  Dim hCurMenu        As Long
  Dim lType           As Long
  Dim bMainMenu       As Boolean
  Dim bMenuFound      As Boolean
  Dim lId             As Long
  Dim lIdx            As Long
  Dim lPos            As Integer
  Dim Ret             As Long

    If m_oSubclass Is Nothing Or m_oClientSubclass Is Nothing Then
        Exit Sub
    End If
    If hwnd = m_hFormHwnd Or hwnd = m_oClientSubclass.hwnd Then
        Select Case uMsg
          Case WM_INITMENUPOPUP
            '--- first, give WindowList menu a chance to fill visible MDI children
            lReturn = m_oSubclass.CallOrigWndProc(uMsg, wParam, lParam)
            bHandled = True
            '--- then, change type to ownerdrawn
            Call pvInitMenu(wParam, False)
          Case WM_MEASUREITEM
            If wParam = 0 Then
                '--- first, forward to child MDI window
                hMdiChild = pvGetMdiChild
                If hMdiChild <> 0 Then
                    If SendMessage(hMdiChild, uMsg, wParam, lParam) <> 0 Then
                        bHandled = True
                        lReturn = 1
                        Exit Sub
                    End If
                End If
                '--- then, process locally
                If pvMeasureItem(lParam) Then
                    bHandled = True
                    lReturn = 1
                End If
            End If
          Case WM_DRAWITEM
            If wParam = 0 Then
                '--- first, forward to child MDI window
                hMdiChild = pvGetMdiChild
                If hMdiChild <> 0 Then
                    If SendMessage(hMdiChild, uMsg, wParam, lParam) <> 0 Then
                        bHandled = True
                        lReturn = 1
                        Exit Sub
                    End If
                End If
                '--- then, process locally
                If pvDrawItem(lParam) Then
                    bHandled = True
                    lReturn = 1
                End If
            End If
          Case WM_NCCALCSIZE
            If m_hFormMenu = 0 Then
                m_hFormMenu = GetMenu(m_hFormHwnd)
                If m_hFormMenu <> 0 Then
                    '--- set main menu ownerdrawn
                    Call pvInitMenu(m_hFormMenu, True)
                End If
            End If
          Case WM_MENUSELECT
            If m_cMenuSubclass.Count > 0 Then
                hPrevMenuWnd = m_cMenuSubclass(m_cMenuSubclass.Count).hwnd
                '--- win9x: if not positioned yet -> delay message
                If IsWindowVisible(hPrevMenuWnd) = 0 And Not bDelayed Then
                    bDelayed = True
                    PostMessage hwnd, uMsg, wParam, lParam
                    lReturn = 0
                    bHandled = True
                    Exit Sub
                End If
            End If
            bDelayed = False
            m_hLastMenu = GetSubMenu(lParam, wParam And &HFFFF&)
            If m_hLastMenu = 0 Then
                m_hLastMenu = lParam
            End If
            '--- if system menu -> dont position at all
            If (wParam And (MF_SYSMENU * &H10000)) <> 0 Then
                m_hLastSelMenu = 0
              Else
                GetMenuItemRect IIf(lParam = m_hFormMenu, _
                                IIf((wParam And &H2000000) <> 0, _
                                m_hParentHwnd, _
                                m_hFormHwnd), _
                                hPrevMenuWnd), lParam, wParam And &HFFFF&, m_rcLastSelMenu
                '--- get item info
                With mii
                    If OsVersion >= &H40A Then '--- &H40A = win98 and later
                        .cbSize = Len(mii)
                        .fMask = MIIM_FTYPE
                      Else
                        .cbSize = Len(mii) - 4
                        .fMask = MIIM_TYPE
                        sBuffer = String(1024, 0)
                        .dwTypeData = StrPtr(sBuffer)
                        .cch = Len(sBuffer)
                    End If
                End With
                Call GetMenuItemInfo(lParam, wParam And &HFFFF&, 1, mii)
                m_bLastSelMenuRightAlign = (mii.fType And MFT_RIGHTJUSTIFY) <> 0
                m_hLastSelMenu = lParam
            End If
            '--- if MDI child -> flag and forward
            hMdiChild = pvGetMdiChild
            If hMdiChild <> 0 Then
                SendMessage hMdiChild, WM_MENUSELECT, &H2000000 Or (wParam And &H2000FFFF), lParam
            End If

            '--- menuchar function added for NT operating systems
            '--- added by NR 05/11/2003
            '--- improved and fixed by Vlad 17/11/2003
          Case WM_MENUCHAR
            '--- first, forward to child MDI  window
            Debug.Print "wm_menuchar2"
            hMdiChild = pvGetMdiChild

            If hMdiChild <> 0 Then
                lReturn = SendMessage(hMdiChild, pvInitMenuMsg, ucsIniMenuChar + (wParam And &HFFFF&) * &H10000, lParam)
                If lReturn <> 0 Then
                    bHandled = True
                    Exit Sub
                End If
            End If

            '---  then process  locally
            bHandled = pvHandleMenuChar(lParam, wParam And &HFFFF&, lReturn)

          Case WM_SYSCOLORCHANGE
            pvGetMeasures
          Case WM_NCDESTROY
            Call pvRestoreMenus(0)
            Do While m_cMenuSubclass.Count > 0
                m_cMenuSubclass.Remove 1
            Loop

            Do While m_cBmps.Count > 0
                m_cBmps.Remove 1
            Loop
            Do While m_cMemDC.Count > 0
                m_cMemDC.Remove 1
            Loop
            '--- release circular references (ISubclassingSink interface of Me)
            m_oSubclass.Unsubclass
            m_oClientSubclass.Unsubclass
#If WEAK_REF_CURRENTMENU Then
            CopyMemory VarPtr(g_oCurrentMenu), VarPtr(0), 4
#Else
            Set g_oCurrentMenu = Nothing
#End If

          Case WM_MDISETMENU
            If m_hFormMenu <> wParam Then
                Call pvRestoreMenus(0)
                m_hFormMenu = wParam
                '--- set main menu ownerdrawn
                Call pvInitMenu(m_hFormMenu, True)
            End If
          Case pvInitMenuMsg
            bHandled = True
            lReturn = 1
            Select Case (wParam And &HFFFF&)
              Case ucsIniMenu
                Call pvInitMenu(lParam, False)
              Case ucsIniMainMenu
                Call pvInitMenu(lParam, True)
              Case ucsIniExitMenuLoop
                m_bExpectingPopup = False
              Case ucsIniEnterMenuLoop
                m_bExpectingPopup = True
                m_hFormMenu = lParam
              Case ucsIniParentForm
                m_hParentHwnd = lParam
              Case ucsIniMenuChar
                Debug.Print "ucsinimenuchar"
                pvHandleMenuChar lParam, wParam \ &H10000, lReturn
            End Select
          Case WM_ENTERMENULOOP
            '--- first, forward to child MDI window
            hMdiChild = pvGetMdiChild
            If hMdiChild <> 0 Then
                If SendMessage(hMdiChild, pvInitMenuMsg, ucsIniEnterMenuLoop, m_hFormMenu) = 1 Then
                    Call SendMessage(hMdiChild, pvInitMenuMsg, ucsIniParentForm, m_hFormHwnd)
                    bHandled = True
                    lReturn = 1
                    Exit Sub
                End If
            End If
            '--- then process locally
            m_bExpectingPopup = True
          Case WM_EXITMENULOOP
            '--- first, forward to child MDI window
            hMdiChild = pvGetMdiChild
            If hMdiChild <> 0 Then
                If SendMessage(hMdiChild, pvInitMenuMsg, ucsIniExitMenuLoop, 0) = 1 Then
                    bHandled = True
                    lReturn = 1
                    RaiseEvent Highlight("")
                    Exit Sub
                End If
            End If
            '--- then process locally
            m_bExpectingPopup = False
            '-- Raise Highlight Event To Display Nothing As The Menu is Being Destroyed
            RaiseEvent Highlight("")
        End Select
      Else '--- popup menus
        Select Case uMsg
          Case WM_ERASEBKGND
            lsHwnd = hwnd
            m_blnFirstMenuInitialize = False
            pvGetBackground(hwnd).BitBlt wParam, -1, -2
            bHandled = True
            lReturn = 1
          Case WM_NCCALCSIZE
            m_blnFirstMenuInitialize = False
            CopyMemory VarPtr(rc), lParam, Len(rc)
            With rc
                .Left = .Left + 1
                .Top = .Top + 2
                .Right = .Right - 4
                .Bottom = .Bottom - 2
            End With
            If wParam Then
                CopyMemory lParam, VarPtr(rc), Len(rc)
                CopyMemory lParam + 2 * Len(rc), VarPtr(rc), Len(rc)
            End If
            bHandled = True
            lReturn = 0
          Case WM_NCPAINT
            '-- Ammended To Take Care Of The Scroll Effect
            '-- When Menu Animations Are Active.
            '-- Gary Noble (Phantom Man - PSC)
            If Not m_blnFirstMenuInitialize Then
                GetWindowRect hwnd, rc
                hDC = GetWindowDC(hwnd)
                ExcludeClipRect hDC, 1, 2, rc.Right - rc.Left - 6, rc.Bottom - rc.Top - 4
                pvGetBackground(hwnd).BitBlt hDC
                Call ReleaseDC(hwnd, hDC)
                bHandled = True
                lReturn = 0
              Else
                bHandled = False
                lReturn = 0
            End If
          Case WM_WINDOWPOSCHANGING
            CopyMemory VarPtr(wp), lParam, Len(wp)
            m_blnAutoColumnTop = False
            m_blnPopupLeftMost = False
            m_blnPopupAbove = False

            If (wp.Flags And SWP_NOMOVE) = 0 Then
                If m_hLastSelMenu <> 0 Then
                    GetWindowRect hwnd, rc

                    '-- Added Gary Noble
                    '-- Moves The Menu To The Right
                    '-- To Take Care Of The Painting When A Popup Menu Is Unloaded
                    'If Me.DisplayShadow Then wp.X = wp.X + 4

                    If m_hLastSelMenu = m_hFormMenu Then

                        '-- Reposition The Menu
                        '-- Gary Noble 2003
                        '--corrected next line NR
                        '   If Me.DisplayShadow Then wp.X = wp.X - 4
                        If m_rcLastSelMenu.Left - 5 > wp.x Then
                            m_blnPopupLeftMost = True
                            'lSpecialLineOffset = rc.Left + m_rcLastSelMenu.Left - wp.X
                            '--corrected NR
                            lSpecialLineOffset = m_rcLastSelMenu.Left - wp.x
                          ElseIf m_rcLastSelMenu.Left < 0 Then
                            wp.x = m_rcLastSelMenu.Left
                            'lSpecialLineOffset = (m_rcLastSelMenu.Left + rc.Left + wp.X) - m_rcLastSelMenu.Left
                            lSpecialLineOffset = 0
                          ElseIf wp.x > rc.Left + m_rcLastSelMenu.Left - 1 Then
                            lSpecialLineOffset = m_rcLastSelMenu.Left + rc.Left - wp.x
                          Else
                            lSpecialLineOffset = 0
                        End If

                        If m_rcLastSelMenu.Right > (Screen.Width / Screen.TwipsPerPixelX) Then
                            '   wp.X = m_rcLastSelMenu.Left
                            lSpecialLineOffset = m_rcLastSelMenu.Left - wp.x
                        End If

                        If wp.y > m_rcLastSelMenu.Bottom - 1 Then
                            wp.y = m_rcLastSelMenu.Bottom - 1
                            m_blnPopupAbove = True
                          Else
                            m_blnPopupAbove = True
                            If AutoColumn > 0 Then
                                m_blnAutoColumnTop = True
                              Else
                                m_blnAutoColumnTop = False
                            End If

                            wp.y = m_rcLastSelMenu.Top - (rc.Bottom - rc.Top - 4)

                        End If

                        If m_bLastSelMenuRightAlign Then
                            wp.x = wp.x + 5
                        End If
                      Else
                        If (rc.Bottom - rc.Top) + m_rcLastSelMenu.Top < GetSystemMetrics(SM_CYSCREEN) Then
                            wp.y = m_rcLastSelMenu.Top
                          Else

                            wp.y = GetSystemMetrics(SM_CYSCREEN) - (rc.Bottom - rc.Top)
                        End If
                        If m_bLastSelMenuRightAlign Then
                            wp.x = wp.x + 3
                        End If
                    End If
                    If wp.y < 0 Then
                        wp.y = 0
                    End If
                    CopyMemory lParam, VarPtr(wp), Len(wp)
                End If

                m_ptLast.x = wp.x
                m_ptLast.y = wp.y
            End If
          Case WM_PRINT
            m_blnFirstMenuInitialize = True
            pvGetBackground(hwnd).BitBlt wParam
            GetViewportOrgEx wParam, VarPtr(pt)
            SetViewportOrgEx wParam, pt.x + 1, pt.y + 2, 0
            Set oSub = m_cMenuSubclass("#" & hwnd)
            lReturn = oSub.CallOrigWndProc(WM_PRINTCLIENT, wParam, lParam)
            SetViewportOrgEx wParam, pt.x, pt.y, 0
            '--- winxp: remove clipping because the dc in wParam will be reused
            '---   for WM_PRINT-ing all menus systemwide!
            SelectClipRgn wParam, 0
            bHandled = True
          Case WM_SHOWWINDOW, WM_DESTROY
            On Error Resume Next
                If (wParam And &HFFFF&) = 0 Or uMsg = WM_DESTROY Then
                    '--- call original
                    Set oSub = m_cMenuSubclass("#" & hwnd)
                    lReturn = oSub.CallOrigWndProc(uMsg, wParam, lParam)
                    bHandled = True
                    '--- win9x and NT only: restore window styles
                    If Not IsNT Or OsVersion <= &H400 Then
                        SetWindowLong hwnd, GWL_STYLE, oSub.Tag(0) Or WS_VISIBLE Or WS_BORDER
                        SetWindowLong hwnd, GWL_EXSTYLE, oSub.Tag(1) Or WS_EX_DLGMODALFRAME Or WS_EX_WINDOWEDGE
                    End If
                    '--- remove subclasser (effectively unsubclassing)
                    m_cMenuSubclass.Remove "#" & hwnd
                    '-- Raise Highlight Event To Display Nothing As The Menu is Being Destroyed
                    RaiseEvent Highlight("")

                    ' If OsVersion = VER_PLATFORM_WIN2000 Then SendMessage lsHwnd, WM_NCPAINT, wParam, lParam
                    '--- remove cache (free resources)
                    m_cMemDC.Remove "#" & hwnd
                End If
            End Select
        End If

    On Error GoTo 0

End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)

    pvGetMeasures

End Sub

Public Property Get MenuDrawStyle() As UcsDrawStyle

    MenuDrawStyle = m_MenuDrawStyle

End Property

Public Property Let MenuDrawStyle(ByVal New_MenuDrawStyle As UcsDrawStyle)

    m_MenuDrawStyle = New_MenuDrawStyle
    PropertyChanged "MenuDrawStyle"

End Property

Private Property Get OsVersion() As Long

  Static lVersion     As Long
  Dim uVer            As OSVERSIONINFO

    If lVersion = 0 Then
        uVer.dwOSVersionInfoSize = Len(uVer)
        If GetVersionEx(uVer) Then
            lVersion = uVer.dwMajorVersion * &H100 + uVer.dwMinorVersion
        End If
    End If
    OsVersion = lVersion

End Property

Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long

  Dim clrFore         As UcsRgbQuad
  Dim clrBack         As UcsRgbQuad

    OleTranslateColor clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor clrSecond, 0, VarPtr(clrBack)
    With clrFore
        .R = (.R * lAlpha + clrBack.R * (255 - lAlpha)) / 255
        .G = (.G * lAlpha + clrBack.G * (255 - lAlpha)) / 255
        .B = (.B * lAlpha + clrBack.B * (255 - lAlpha)) / 255
    End With
    CopyMemory VarPtr(pvAlphaBlend), VarPtr(clrFore), 4

End Function

'-- For Win2000 Only
'-- Drawing From The PVGetBackGround Seemed To Cause A Few Problems
'-- In Win2000 - This Is the Work Around At Last
'-- Gary Noble - 12-11-2003
Private Sub pvDoWin2000Borders(oMemDC As cMemDC, rc As RECT, rcItem As RECT, rcPopup As RECT, rcPopupBtm As RECT, m_bLastSelMenuRightAlign As Boolean, lWidth As Long, lHeight As Long, lHorShadowStart As Long, lHorShadowEnd As Long)

    With oMemDC
        If m_blnAutoColumnTop Then
            If AutoColumn = 0 Then
                .FillRect lSpecialLineOffset, lHeight - 1, lSpecialLineOffset - rcItem.Right - rcItem.Left - 1, lHeight, m_clrMenuBack
                lHorShadowStart = lSpecialLineOffset + rcItem.Right - rcItem.Left - 2
              Else
                .FillRect lSpecialLineOffset + 1, lHeight - 1, lSpecialLineOffset + rcItem.Right - rcItem.Left - 1, lHeight, m_clrMenuBack
                lHorShadowStart = rcItem.Right - rcItem.Left - 2
            End If
          Else
            If rcItem.Bottom + rcPopupBtm.Bottom - rcPopup.Top + 2 * m_lFrameWidth <= GetSystemMetrics(SM_CYSCREEN) Then
                If Not m_bLastSelMenuRightAlign Then
                    .FillRect lSpecialLineOffset + 1, 0, lSpecialLineOffset + IIf(AutoColumn > 0, 0, 0) + rcItem.Right - rcItem.Left - 1, 1, m_clrMenuBack
                  Else
                    .FillRect lWidth - (rcItem.Right - rcItem.Left - 1), 0, lSpecialLineOffset - 1 + lWidth - 1, 1, m_clrMenuBack
                End If
                lHorShadowStart = 0
              ElseIf rcPopupBtm.Bottom > rcItem.Top Then
                If Not m_bLastSelMenuRightAlign Then
                    .FillRect lSpecialLineOffset + 1, lHeight - 1, lSpecialLineOffset + rcItem.Right - rcItem.Left - 1, lHeight, m_clrMenuBack
                    lHorShadowStart = rcItem.Right - rcItem.Left - 2
                  Else
                    .FillRect lWidth - (rcItem.Right - rcItem.Left - 1), lHeight - 1, lWidth - 1, lHeight, m_clrMenuBack
                    lHorShadowEnd = lHorShadowEnd - (rcItem.Right - rcItem.Left + 3)
                End If
            End If
        End If
    End With

End Sub

Private Function pvDrawItem(ByVal lParam As Long) As Boolean

  Dim lI              As Long
  Dim lJ              As Long
  Dim lK              As Long
  Dim clrBack         As Long
  Dim clrBorder       As Long
  Dim dis             As DRAWITEMSTRUCT
  Dim hMenu           As Long
  Dim sText           As String
  Dim lType           As Long
  Dim bMainMenu       As Boolean
  Dim lId             As Long
  Dim rc              As RECT
  Dim lState          As Long
  Dim vPic            As Variant
  Dim oPicMemDC       As cMemDC
  Dim vSplit          As Variant
  Dim mii             As MENUITEMINFO
  Dim bCustom         As Boolean
  Dim bExclude        As Boolean
  Dim oldColor        As OLE_COLOR
  Dim oFntOrig        As StdFont
  Dim oFntNew         As StdFont
  Dim oldColorHoverBack As OLE_COLOR
  Dim oldColorHoverBorder As OLE_COLOR

    bCustom = IIf(Me.DrawStyle = MS_Custom, False, True)

    '--- dereference structure
    CopyMemory VarPtr(dis), lParam, Len(dis)
    If dis.CtlType = ODT_MENU Then
        '--- win95: int->long conversion troubles
        If Not IsNT Then
            dis.itemID = (dis.itemID And &HFFFF&)
        End If
        '--- get menu info
        Call pvGetMenuInfo(dis.ItemData, hMenu, sText, lType, bMainMenu, lId)
        '--- Fire Our Custom Item Draw Event

        '--- if not found -> bail out immediately
        If dis.itemID <> lId Or dis.hwndItem <> hMenu Then
            RaiseEvent Highlight("")
            Exit Function
        End If
        '--- get menu state
        lState = GetMenuState(hMenu, dis.itemID, MF_BYCOMMAND)
        With New cMemDC
            .Init hMemoryDC:=dis.hDC
            '--- setup memory (buffer) device-context
            .Init dis.rcItem.Right - dis.rcItem.Left + 3, dis.rcItem.Bottom - dis.rcItem.Top + 1, dis.hDC
            .LoadBlt dis.hDC, dis.rcItem.Left, dis.rcItem.Top
            SetViewportOrgEx .hDC, -dis.rcItem.Left, -dis.rcItem.Top, 0
            '--- init device-context settings (font)
            .BackStyle = BS_TRANSPARENT
            oldColor = m_clrMenuFore

            If Not UseSystemFont Then
                '--- merge fonts
  Dim oFnt As StdFont
                Set oFnt = .Font
                oFnt.Name = Font.Name
                oFnt.Size = Font.Size
                oFnt.Bold = oFnt.Bold Or Font.Bold
                oFnt.Italic = oFnt.Italic Or Font.Italic
                oFnt.Underline = oFnt.Underline Or Font.Underline
                oFnt.Strikethrough = oFnt.Strikethrough Or Font.Strikethrough

                '-- Added Gary Noble - 5-11-2003
                '-- Raise The CustomDrawItemFont Event
                vSplit = Split(sText, vbTab)
                If Me.DrawStyle = MS_Custom Then
                    If UBound(vSplit) >= 0 Then
                        '-- Raise Our Custom Draw Item Event
                        RaiseEvent CustomDrawItemFont(oFnt, CStr(Replace(IIf(Left$(vSplit(0), 1) = "&", Right(vSplit(0), Len(vSplit(0)) - 1), vSplit(0)), "&&", "&")), oldColor)
                    End If
                End If
                Set .Font = oFnt
            End If

            '--- check if drawing main menu item
            If bMainMenu Then
                '--- fill background
                .FillRect dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right + 3, dis.rcItem.Bottom, IIf(bCustom, m_clrMenuBarBack, m_UserTopMenuBackColour)
                If Not bCustom And m_UserUseTopMenuGradient Then .FillGradient dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right + 4, dis.rcItem.Bottom, m_UserGradientOne, m_UserGradientTwo, True
                If (lState And MF_GRAYED) = 0 Then
                    '.ForeColor = m_clrMenuFore
                    .ForeColor = oldColor
                    If (dis.itemState And ODS_SELECTED) <> 0 Then
                        .Rectangle dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom, IIf(bCustom, m_clrMenuBack, m_UserTopMenuSelectedColour), , IIf(bCustom, m_clrMenuBorder, m_UserMenuBorderColour)
                        If Not bCustom And m_UserUseTopMenuGradient Then .FillGradient dis.rcItem.Left + 1, dis.rcItem.Top + 1, dis.rcItem.Right - 1, dis.rcItem.Bottom - 1, m_UserGradientOne, m_UserGradientTwo, True
                        .ForeColor = IIf(bCustom, m_clrSelMenuFore, m_UserSelectedItemForeColour)
                      ElseIf (dis.itemState And ODS_HOTLIGHT) <> 0 Then
                        .Rectangle dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom, IIf(bCustom, m_clrSelMenuBack, m_UserTopMenuHotColour), , IIf(bCustom, m_clrSelMenuBorder, m_UserTopMenuHotBorderColour)
                        If Not bCustom And m_UserUseTopMenuGradient Then .FillGradient dis.rcItem.Left + 1, dis.rcItem.Top + 1, dis.rcItem.Right - 1, dis.rcItem.Bottom - 1, m_UserGradientTwo, m_UserGradientOne, True
                        .ForeColor = IIf(bCustom, m_clrSelMenuFore, m_UserSelectedItemForeColour)
                    End If
                  Else
                    .ForeColor = m_clrDisabledMenuFore
                End If
                '--- draw text
                .DrawText sText, dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
                '--- draw main menu shadow
                If Me.DisplayShadow Then
                    If (dis.itemState And ODS_SELECTED) <> 0 Then
                        If m_bConstrainedColors Then
                            .FillRect dis.rcItem.Right, dis.rcItem.Top + 3, dis.rcItem.Right + 2, dis.rcItem.Bottom, vbButtonShadow
                          Else
                            For lJ = 0 To 2
                                For lK = 3 To dis.rcItem.Bottom - dis.rcItem.Top - 1
                                    .SetPixel dis.rcItem.Right + lJ, dis.rcItem.Top + lK, pvAlphaBlend(vbBlack, .GetPixel(dis.rcItem.Right + lJ, dis.rcItem.Top + lK), &H40 - lJ * (&H40 / 3))
                                Next lK
                            Next lJ
                        End If
                    End If
                End If
              Else
                If (lType And MFT_SEPARATOR) = 0 Then
                    vSplit = Split(sText, vbTab)
                    If UBound(vSplit) >= 0 Then
                        '--- figure out current bitmap appearance
                        vPic = pvFindIconInfo(vSplit(0))
                        If IsArray(vPic) Then
                            If Not vPic(0) Is Nothing Then
                                If (lState And MF_DISABLED) <> 0 Then
                                    Set oPicMemDC = pvGetBitmapDisabled(vPic(0), vPic(1), pvGetLuminance(m_clrMenuBack))
                                  ElseIf (lState And MF_CHECKED) <> 0 Then
                                    Set oPicMemDC = pvGetBitmapNormal(vPic(0), vPic(1))
                                  ElseIf (dis.itemState And ODS_SELECTED) <> 0 Then
                                    Set oPicMemDC = pvGetBitmapRaised(vPic(0), vPic(1))
                                  Else
                                    Set oPicMemDC = pvGetBitmapDimmed(vPic(0), vPic(1))
                                End If
                            End If
                        End If
                    End If
                End If
                '--- fill background
                If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                    If Not bCustom And m_UserUseGradient Then
                        .FillGradient dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Left + m_lMenuHeight + 3, dis.rcItem.Bottom, m_UserGradientOne, m_UserGradientTwo
                      Else
                        .FillRect dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Left + m_lMenuHeight + 3, dis.rcItem.Bottom, IIf(bCustom, m_clrMenuBack, m_UserSideBarColour)
                    End If
                  Else
                    If Not bCustom And m_UserUseGradient Then
                        .FillGradient dis.rcItem.Right - m_lMenuHeight - 4, dis.rcItem.Top, dis.rcItem.Right - 1, dis.rcItem.Bottom, m_UserGradientOne, m_UserGradientTwo
                      Else
                        .FillRect dis.rcItem.Right - m_lMenuHeight - 4, dis.rcItem.Top, dis.rcItem.Right - 1, dis.rcItem.Bottom, IIf(bCustom, m_clrMenuBack, m_UserSideBarColour)
                    End If
                End If
                '--- get ID of last menu item in current (popup) menu
                mii.cbSize = Len(mii) - IIf(OsVersion >= &H40A, 0, 4)
                mii.fMask = MIIM_ID
                Call GetMenuItemInfo(hMenu, GetMenuItemCount(hMenu) - 1, 1, mii)

                '--- if current item is the last one in the popup menu -> paint a white line at the bottom
                If mii.wID = lId Then
                    .FillRect dis.rcItem.Left, dis.rcItem.Bottom - 1, dis.rcItem.Right - 1, dis.rcItem.Bottom, m_clrMenuPopupBack
                End If

                If OsVersion >= &H40A Then '--- &H40A = win98 and later
                    mii.cbSize = Len(mii)
                    mii.fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_DATA Or MIIM_STRING
                  Else
                    mii.cbSize = Len(mii) - 4
                    mii.fMask = MIIM_ID Or MIIM_TYPE Or MIIM_DATA
                End If

                Call GetMenuItemInfo(hMenu, dis.itemID + 1, 0, mii)

                If (mii.fType And MF_MENUBREAK) <> 0 Then
                    .FillRect dis.rcItem.Left, dis.rcItem.Bottom - 1, dis.rcItem.Right - 1, dis.rcItem.Bottom, m_clrMenuPopupBack
                End If

                If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                    .FillRect dis.rcItem.Left + m_lMenuHeight + 3, dis.rcItem.Top, dis.rcItem.Right - 1, dis.rcItem.Bottom, m_clrMenuPopupBack
                  Else
                    .FillRect dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right - m_lMenuHeight - 4, dis.rcItem.Bottom, m_clrMenuPopupBack
                End If

                '-- Ammended - Gary Noble
                If AutoColumn > 0 Then .FillRect dis.rcItem.Right - 1, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom, IIf(bCustom, m_clrMenuBorder, m_UserMenuBorderColour)

                If (lType And MFT_SEPARATOR) = MFT_SEPARATOR Then
                    '--- draw separator line
                    If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                        .FillRect dis.rcItem.Left + m_lMenuHeight + 10, (dis.rcItem.Top + dis.rcItem.Bottom) \ 2 - 1, dis.rcItem.Right, (dis.rcItem.Top + dis.rcItem.Bottom) \ 2, IIf(Not bCustom, m_UserMenuBorderColour, m_clrMenuBorder)
                      Else
                        .FillRect dis.rcItem.Left, (dis.rcItem.Top + dis.rcItem.Bottom) \ 2 - 1, dis.rcItem.Right - 1 - m_lMenuHeight - 10, (dis.rcItem.Top + dis.rcItem.Bottom) \ 2, IIf(Not bCustom, m_UserMenuBorderColour, m_clrMenuBorder)
                    End If
                  Else
                    '--- if selected -> item background rect (enabled or disabled)
                    If (lState And MF_GRAYED) = 0 Then
                        If (dis.itemState And ODS_SELECTED) <> 0 Then
                            If Me.DrawStyle = MS_Custom And Not Me.UseSystemFont Then
                                '-- Raise Our Custom Draw Item Event
                                oldColorHoverBack = IIf(bCustom, m_clrSelMenuBack, m_UserSelectedMenuBackColour)
                                oldColorHoverBorder = IIf(bCustom, m_clrSelMenuBorder, m_UserSelectedMenuBorderColour)
                                RaiseEvent CustomDrawItemHoverFont(oFnt, CStr(Replace(IIf(Left$(vSplit(0), 1) = "&", Right(vSplit(0), Len(vSplit(0)) - 1), vSplit(0)), "&&", "&")), oldColor, oldColorHoverBack, oldColorHoverBorder)
                                Set .Font = oFnt
                                .ForeColor = oldColor
                                .Rectangle dis.rcItem.Left + 1, dis.rcItem.Top, dis.rcItem.Right - 2, dis.rcItem.Bottom - 1, oldColorHoverBack, , oldColorHoverBorder
                              Else
                                .ForeColor = IIf(bCustom, m_clrSelMenuFore, m_UserSelectedItemForeColour)
                                .Rectangle dis.rcItem.Left + 1, dis.rcItem.Top, dis.rcItem.Right - 2, dis.rcItem.Bottom - 1, IIf(bCustom, m_clrSelMenuBack, m_UserSelectedMenuBackColour), , IIf(bCustom, m_clrSelMenuBorder, m_UserSelectedMenuBorderColour)
                            End If
                            '.FillGradient dis.rcItem.Left + 2, dis.rcItem.Top + 1, dis.rcItem.Right - 3, dis.rcItem.Bottom - 2, m_UserGradienttwo, m_UserSelectedMenuBackColour, True
                            If UBound(vSplit) >= 0 Then RaiseEvent Highlight(CStr(Replace(IIf(Left$(vSplit(0), 1) = "&", Right(vSplit(0), Len(vSplit(0)) - 1), vSplit(0)), "&&", "&")))
                          Else
                            .ForeColor = oldColor
                            '.ForeColor = m_clrMenuFore
                        End If
                      Else
                        .ForeColor = m_clrDisabledMenuFore
                        If (dis.itemState And ODS_SELECTED) <> 0 And SelectDisabled Then
                            .Rectangle dis.rcItem.Left + 1, dis.rcItem.Top, dis.rcItem.Right - 2, dis.rcItem.Bottom - 1, m_clrDisabledMenuBack, , m_clrDisabledMenuBorder
                        End If
                    End If
                    '--- draw check square background and border
                    If (lState And MF_CHECKED) <> 0 Then
                        lI = (m_lMenuHeight - BitmapSize - 1) \ 2
                        lJ = (m_lMenuHeight - BitmapSize - 5) \ 2
                        If (lState And MF_DISABLED) <> 0 Then
                            clrBack = m_clrDisabledMenuBack
                            clrBorder = m_clrDisabledMenuBorder
                          ElseIf (dis.itemState And ODS_SELECTED) <> 0 Then
                            clrBack = m_clrSelCheckBack
                            clrBorder = m_clrSelMenuBorder
                          Else
                            clrBack = m_clrCheckBack
                            clrBorder = m_clrSelMenuBorder
                        End If
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            .Rectangle dis.rcItem.Left + lI, dis.rcItem.Top + lJ, dis.rcItem.Left + lI + BitmapSize + 4, dis.rcItem.Top + lJ + BitmapSize + 4, IIf(bCustom, clrBack, m_UserCheckBackColour), , IIf(bCustom, clrBorder, m_UserCheckBorderColour)
                          Else
                            .Rectangle dis.rcItem.Right - 4 - m_lMenuHeight + lI, dis.rcItem.Top + lJ, dis.rcItem.Right - 4 - m_lMenuHeight + lI + BitmapSize + 4, dis.rcItem.Top + lJ + BitmapSize + 4, IIf(bCustom, clrBack, m_UserCheckBackColour), , IIf(bCustom, clrBorder, m_UserCheckBorderColour)
                        End If
                    End If

                    '--- draw bitmap or check
                    If Not oPicMemDC Is Nothing Then
                        lI = (m_lMenuHeight - BitmapSize - 1) \ 2 - 1
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            oPicMemDC.TransBlt .hDC, dis.rcItem.Left + lI + 2, dis.rcItem.Top + lI, BitmapSize + 2, BitmapSize + 2, clrMask:=MASK_COLOR
                          Else
                            oPicMemDC.TransBlt .hDC, dis.rcItem.Right - 4 - m_lMenuHeight + lI + 2, dis.rcItem.Top + lI, BitmapSize + 2, BitmapSize + 2, clrMask:=MASK_COLOR
                        End If
                      ElseIf (lState And MF_CHECKED) <> 0 Then
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            lI = dis.rcItem.Left + m_lMenuHeight \ 2 - 2
                          Else
                            lI = dis.rcItem.Right - 4 - m_lMenuHeight \ 2 - 2
                        End If
                        For lJ = 0 To 2
                            .DrawLine lI + lJ, dis.rcItem.Top + m_lMenuHeight \ 2 - 2 + lJ + 1, lI + lJ, dis.rcItem.Top + m_lMenuHeight \ 2 + lJ + 1, m_clrCheckFore
                        Next lJ
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            lI = dis.rcItem.Left + m_lMenuHeight \ 2
                          Else
                            lI = dis.rcItem.Right - 4 - m_lMenuHeight \ 2
                        End If
                        For lJ = 1 To 4
                            .DrawLine lI + lJ, dis.rcItem.Top + m_lMenuHeight \ 2 - lJ + 1, lI + lJ, dis.rcItem.Top + m_lMenuHeight \ 2 - lJ + 3, m_clrCheckFore
                        Next lJ
                    End If

                    If UBound(vSplit) >= 0 Then
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            .DrawText vSplit(0), dis.rcItem.Left + m_lMenuHeight + 10, dis.rcItem.Top, dis.rcItem.Right - m_lMenuHeight, dis.rcItem.Bottom - 1, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                          Else
                            .DrawText vSplit(0), dis.rcItem.Left + m_lMenuHeight, dis.rcItem.Top, dis.rcItem.Right - 1 - m_lMenuHeight - 10, dis.rcItem.Bottom - 1, DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
                        End If
                    End If

                    If UBound(vSplit) > 0 Then
                        '--- draw shortcut keys
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            .DrawText vSplit(1), dis.rcItem.Left + m_lMenuHeight + 10, dis.rcItem.Top, dis.rcItem.Right - m_lMenuHeight, dis.rcItem.Bottom - 1, DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
                          Else
                            .DrawText vSplit(1), dis.rcItem.Left + m_lMenuHeight, dis.rcItem.Top, dis.rcItem.Right - 1 - m_lMenuHeight - 10, dis.rcItem.Bottom - 1, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                        End If

                    End If

                    '--- draw submenu arrow (if necessary)
                    bExclude = False
                    If (lState And MF_POPUP) <> 0 Then
                        bExclude = True
                        lI = (dis.rcItem.Top + dis.rcItem.Bottom - 1) \ 2
                        If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                            lJ = dis.rcItem.Right - m_lMenuHeight \ 3 - 4
                          Else
                            lJ = dis.rcItem.Left + m_lMenuHeight * 2 \ 3 - 4
                        End If
                        For lK = (m_lTextHeight + 3) \ 8 To 0 Step -1
                            If Not RightToLeft And (lType And MFT_RIGHTJUSTIFY) = 0 Then
                                .DrawLine lJ - lK, lI - lK, lJ - lK, lI + lK + 1, .ForeColor
                              Else
                                .DrawLine lJ + lK, lI - lK, lJ + lK, lI + lK + 1, .ForeColor
                            End If
                        Next lK
                    End If
                End If
            End If
            If .IsMemoryDC Then
                SetViewportOrgEx .hDC, 0, 0, 0
                .BitBlt dis.hDC, dis.rcItem.Left, dis.rcItem.Top
            End If
        End With
        '--- prevent further drawing (esp sub-menu arrow) and reduce flicker
        If bExclude Then
            ExcludeClipRect dis.hDC, dis.rcItem.Left, dis.rcItem.Top, dis.rcItem.Right, dis.rcItem.Bottom
        End If

        m_clrMenuFore = oldColor
        '--- handled
        pvDrawItem = True
    End If

End Function

Private Function pvFindIconInfo(ByVal sCaption As String) As Variant

  '--- ToDo: better handling
  
  Dim oCtl            As Object

    On Error Resume Next
        For Each oCtl In ParentControls
            If TypeOf oCtl Is Menu Then
                If Split(oCtl.Caption, vbTab)(0) <> sCaption Then
                  Else
                    pvFindIconInfo = m_cBmps("#" & pvGetCtlName(oCtl))
                    Exit Function
                End If
            End If
        Next oCtl
    On Error GoTo 0

End Function

Private Function pvGetBackground(ByVal hwnd As Long) As cMemDC

  Dim oMemDC          As cMemDC
  Dim rc              As RECT
  Dim rcItem          As RECT
  Dim rcPopup         As RECT
  Dim rcPopupBtm      As RECT
  Dim lI              As Long
  Dim lJ              As Long
  Dim lWidth          As Long
  Dim lHeight         As Long
  Dim v               As Variant
  Dim hWndFrm         As Long
  Dim lHorShadowStart As Long
  Dim lHorShadowEnd   As Long

    On Error Resume Next
        Set oMemDC = m_cMemDC("#" & hwnd)
        If oMemDC Is Nothing Then
            GetWindowRect hwnd, rc
            Set oMemDC = New cMemDC
            oMemDC.Init rc.Right - rc.Left, rc.Bottom - rc.Top
            With New cMemDC
                .Init , , , GetWindowDC(GetDesktopWindow())
                .BitBlt oMemDC.hDC, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, m_ptLast.x, m_ptLast.y
                Call ReleaseDC(GetDesktopWindow(), .hDC)
            End With
            lWidth = rc.Right - rc.Left - 2 * m_lFrameWidth + 1
            lHeight = rc.Bottom - rc.Top - 2 * m_lFrameWidth + 3
            lHorShadowEnd = lWidth - 1 + 3
            With oMemDC

                .Rectangle 0, 0, lWidth, lHeight, vbWindowBackground, , IIf(DrawStyle = MS_Default, m_clrMenuBorder, m_UserMenuBorderColour)
                '--- visually improves performance to clear the left band here
                .FillRect 1, 2, m_lMenuHeight + 4, lHeight - 2, m_clrMenuBack
                Call GetMenuItemRect(m_hFormHwnd, m_hLastMenu, 0, rcItem)
                '--- fix the line right below the main menu
                If lSpecialLineOffset < 2 Then lSpecialLineOffset = 0

                For lI = 0 To GetMenuItemCount(m_hFormMenu) - 1
                    '--- find opened main menu item
                    If (GetMenuState(m_hFormMenu, lI, MF_BYPOSITION) And MF_HILITE) <> 0 Then
                        If m_hLastMenu = GetSubMenu(m_hFormMenu, lI) Then
                            '--- get its popup menu dimensions
                            hWndFrm = IIf(m_hParentHwnd <> 0, m_hParentHwnd, m_hFormHwnd)
                            '--- win98: can't pass NULL for hwnd (so use hWndFrm)
                            Call GetMenuItemRect(hWndFrm, m_hLastMenu, 0, rcPopup)
                            Call GetMenuItemRect(hWndFrm, m_hLastMenu, GetMenuItemCount(m_hLastMenu) - 1, rcPopupBtm)
                            '--- get main menu item dimensions
                            Call GetMenuItemRect(hWndFrm, m_hFormMenu, lI, rcItem)
                            '--- if popup below main menu fix border

                            '-- Win2000 Caused Problems When Drawing The White
                            '-- Line From The Menu And Setting The Shadow Params
                            '-- To Keep Things Tidy Call The pvDoWin2000Borders
                            '-- Sub to Calculate The White Line Offset And Shadow Params
                            '-- Gary Noble - 12-11-2003
                            If OsVersion = VER_PLATFORM_WIN2000 Or OsVersion = VER_PLATFORM_WINXP Then
                                pvDoWin2000Borders oMemDC, rc, rcItem, rcPopup, rcPopupBtm, m_bLastSelMenuRightAlign, _
                                                   lWidth, lHeight, lHorShadowStart, lHorShadowEnd
                              Else
                                '-- amended NR 10/11/2003
                                If rc.Bottom > rcPopup.Bottom Then
                                    If Not m_bLastSelMenuRightAlign Then
                                        ' -- amended NR
                                        .FillRect lSpecialLineOffset + 1, 0, IIf(rcItem.Right + 5 > rc.Right, lSpecialLineOffset + (rcItem.Right - rcItem.Left + (rc.Right - rcItem.Right) - 6), rcItem.Right - rcItem.Left - 1 + lSpecialLineOffset), 1, IIf(m_UserUseGradient And DrawStyle = MS_Custom, m_clrMenuPopupBack, m_clrMenuBack)
                                      Else
                                        .FillRect lWidth - (rcItem.Right - rcItem.Left - 1), 0, lWidth - 1, 1, m_clrMenuBack
                                    End If
                                    lSpecialLineOffset = 0
                                    ' -- amended NR 10/11/2003
                                  Else 'If rcPopupBtm.Bottom > rcItem.Top Then

                                    If Not m_bLastSelMenuRightAlign Then
                                        '-- amended NR
                                        .FillRect lSpecialLineOffset + 1, lHeight - 1, IIf(rcItem.Right + 5 > rc.Right, lSpecialLineOffset + (rcItem.Right - rcItem.Left + (rc.Right - rcItem.Right) - 6), rcItem.Right - rcItem.Left - 1 + lSpecialLineOffset), lHeight, IIf(m_UserUseGradient And DrawStyle = MS_Custom, m_clrMenuPopupBack, m_clrMenuBack)
                                        lSpecialLineOffset = lSpecialLineOffset + IIf(rcItem.Left < rcPopupBtm.Left, 1, 0)
                                        lHorShadowStart = rcItem.Right - rcItem.Left - 2
                                      Else
                                        .FillRect lWidth - (rcItem.Right - rcItem.Left - 1), lHeight - 1, lWidth - 1, lHeight, m_clrMenuBack
                                        lHorShadowEnd = lHorShadowEnd - (rcItem.Right - rcItem.Left + 3)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next lI

                lHorShadowStart = lHorShadowStart - 2
                '--- shadow
                If Me.DisplayShadow Then
                    If m_bConstrainedColors Then
                        .FillRect lHorShadowStart + lSpecialLineOffset + 3, lHeight, lHorShadowEnd, lHeight + 2, vbButtonShadow
                        .FillRect lWidth, 3, lWidth + 2, lHeight, vbButtonShadow
                        '-- Draws The Shadow From The Left Of The Menu
                        '-- To The Start Of The lSpecialLineOffset
                        '-- Added By Gary Noble
                        If lSpecialLineOffset > 2 Then
                            .FillRect 3, lHeight, lSpecialLineOffset, lHeight + 2, vbButtonShadow
                        End If

                      Else
                        For lJ = 0 To 2
                            For lI = lHorShadowStart + lSpecialLineOffset + 3 To lHorShadowEnd
                                .SetPixel lI, lHeight + lJ, pvAlphaBlend(vbBlack, .GetPixel(lI, lHeight + lJ), (&H40 - lJ * (&H40 / 3)) * (IIf(lI <= 6, lI - 2, 4) / 4) * (IIf(lI >= lWidth, lWidth + 3 - lI, 4) / 4))
                            Next lI
                            For lI = 3 To lHeight - 1
                                .SetPixel lWidth + lJ, lI, pvAlphaBlend(vbBlack, .GetPixel(lWidth + lJ, lI), (&H40 - lJ * (&H40 / 3)) * (IIf(lI <= 6, lI - 2, 4) / 4))
                            Next lI
                        Next lJ
                        '-- Draws The Shadow From The Left Of The Menu
                        '-- To The Start Of The lSpecialLineOffset
                        '-- Added By Gary Noble
                        If lSpecialLineOffset > 2 Then
                            For lJ = 0 To 2
                                For lI = 3 To lSpecialLineOffset
                                    .SetPixel lI, lHeight + lJ, pvAlphaBlend(vbBlack, .GetPixel(lI, lHeight + lJ), (&H40 - lJ * (&H40 / 3)) * (IIf(lI <= 6, lI - 2, 4) / 4) * (IIf(lI >= lWidth, lWidth + 3 - lI, 4) / 4))
                                Next lI
                                For lI = 3 To lHeight - 1
                                    .SetPixel lWidth + lJ, lI, pvAlphaBlend(vbBlack, .GetPixel(lWidth + lJ, lI), (&H40 - lJ * (&H40 / 3)) * (IIf(lI <= 6, lI - 2, 4) / 4))
                                Next lI
                            Next lJ
                        End If

                    End If
                End If
            End With
            m_cMemDC.Add oMemDC, "#" & hwnd
        End If
QH:
        Set pvGetBackground = oMemDC
    On Error GoTo 0

End Function

Private Function pvGetBitmapDimmed(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR) As cMemDC

  Dim lI              As Long
  Dim lJ              As Long

    Set pvGetBitmapDimmed = New cMemDC
    With pvGetBitmapDimmed
        .Init BitmapSize + 2, BitmapSize + 2
        .Cls MASK_COLOR
        .PaintPicture oPic, (.Width - pvHM2Pix(oPic.Width)) \ 2, (.Height - pvHM2Pix(oPic.Height)) \ 2, clrMask:=MaskColor

        If MenuDrawStyle = DS_XP Then
            For lJ = 0 To BitmapSize
                For lI = 0 To BitmapSize
                    If .GetPixel(lI, lJ) <> MASK_COLOR Then
                        .SetPixel lI, lJ, pvAlphaBlend(m_clrMenuBack, .GetPixel(lI, lJ), 70)
                    End If
                Next lI
            Next lJ
        End If

    End With

End Function

Private Function pvGetBitmapDisabled(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR, ByVal TresholdLuminance As Long) As cMemDC

  Dim lI              As Long
  Dim lJ              As Long
  Dim lK              As Long

    Set pvGetBitmapDisabled = New cMemDC
    With pvGetBitmapDisabled
        .Init BitmapSize + 2, BitmapSize + 2
        .Cls MASK_COLOR
        .PaintPicture oPic, (.Width - pvHM2Pix(oPic.Width)) \ 2 - 1, (.Height - pvHM2Pix(oPic.Height)) \ 2 - 1, clrMask:=MaskColor

        If MenuDrawStyle = DS_XP Then
            For lJ = 0 To .Width
                For lI = 0 To .Height
                    lK = .GetPixel(lI, lJ)
                    If lK <> MASK_COLOR Then
                        If pvGetLuminance(lK) < TresholdLuminance Then
                            .SetPixel lI, lJ, vbButtonShadow
                          Else
                            .SetPixel lI, lJ, MASK_COLOR
                        End If
                    End If
                Next lI
            Next lJ
        End If
    End With

End Function

Private Function pvGetBitmapNormal(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR) As cMemDC

    Set pvGetBitmapNormal = New cMemDC
    With pvGetBitmapNormal
        .Init BitmapSize + 2, BitmapSize + 2
        .Cls MASK_COLOR
        .PaintPicture oPic, (.Width - pvHM2Pix(oPic.Width)) \ 2, (.Height - pvHM2Pix(oPic.Height)) \ 2, clrMask:=MaskColor
    End With

End Function

Private Function pvGetBitmapRaised(ByVal oPic As StdPicture, ByVal MaskColor As OLE_COLOR) As cMemDC

  Dim lI              As Long
  Dim lJ              As Long

    Set pvGetBitmapRaised = New cMemDC

    With pvGetBitmapRaised

        '-- Added By Gary Noble 5-11-2003
        If MenuDrawStyle = DS_XP Then
            .Init BitmapSize + 2, BitmapSize + 2
          Else
            .Init BitmapSize + 4, BitmapSize + 4
        End If

        .Cls MASK_COLOR
        .PaintPicture oPic, (.Width - pvHM2Pix(oPic.Width)) \ 2 - 1, (.Height - pvHM2Pix(oPic.Height)) \ 2 - 1, clrMask:=MaskColor
        If MenuDrawStyle = DS_XP Then
            For lJ = .Width To 2 Step -1
                For lI = .Width To 2 Step -1
                    If .GetPixel(lI - 2, lJ - 2) <> MASK_COLOR And .GetPixel(lI, lJ) = MASK_COLOR Then
                        .SetPixel lI, lJ, pvAlphaBlend(m_clrSelMenuBack, vbButtonShadow, 70)
                    End If
                Next lI
            Next lJ
        End If
    End With

End Function

Private Function pvGetColorBits() As Long

  Dim hDC             As Long

    hDC = GetWindowDC(GetDesktopWindow())
    pvGetColorBits = GetDeviceCaps(hDC, BITSPIXEL) * GetDeviceCaps(hDC, PLANES)
    Call ReleaseDC(GetDesktopWindow(), hDC)

End Function

Private Function pvGetCtlName(ByVal oCtl As Control) As String

    On Error Resume Next
        If oCtl.Index < 0 Then
            pvGetCtlName = oCtl.Name
          Else
            pvGetCtlName = oCtl.Name & ":" & oCtl.Index
        End If
    On Error GoTo 0

End Function

Private Function pvGetLuminance(ByVal clrColor As Long) As Long

  Dim rgbColor        As UcsRgbQuad

    OleTranslateColor clrColor, 0, VarPtr(rgbColor)
    pvGetLuminance = (rgbColor.R * 76& + rgbColor.G * 150& + rgbColor.B * 29&) \ 255&

End Function

Private Function pvGetMdiChild() As Long

    If Not m_oClientSubclass Is Nothing Then
        If m_oClientSubclass.hwnd <> 0 Then
            pvGetMdiChild = m_oClientSubclass.CallOrigWndProc(WM_MDIGETACTIVE, 0, 0)
        End If
    End If

End Function

Private Sub pvGetMeasures()

  '--- get border widths

    m_lEdgeWidth = GetSystemMetrics(SM_CXEDGE)
    m_lFrameWidth = GetSystemMetrics(SM_CXDLGFRAME)
    m_bConstrainedColors = pvIsHightContrast() Or pvGetColorBits() <= 8
    If m_bConstrainedColors Then
        '--- calc high-contrast colors
        m_clrSelMenuBorder = vbButtonText
        If pvIsHightContrast() Then
            m_clrSelMenuBack = vbHighlight
            m_clrSelMenuFore = vbHighlightText
          Else
            m_clrSelMenuFore = vbWindowText
            m_clrSelMenuBack = vbWindowBackground
        End If
        m_clrCheckBack = m_clrSelMenuBack
        m_clrCheckFore = m_clrSelMenuFore
        m_clrSelCheckBack = m_clrSelMenuBack
        m_clrMenuBorder = vbButtonText
        m_clrMenuBack = vbButtonFace
        m_clrMenuFore = vbMenuText
        m_clrDisabledMenuBorder = vbButtonText
        m_clrDisabledMenuBack = m_clrSelMenuBack ' pvAlphaBlend(m_clrMenuBack, vbWindowBackground, 128)
        m_clrDisabledMenuFore = vbGrayText
        If pvIsAppearanceXpStyle() Then
            m_clrMenuBarBack = GetSysColor(COLOR_MENUBAR)
          Else
            m_clrMenuBarBack = vbMenuBar
        End If
        m_clrMenuPopupBack = vbWindowBackground
      Else
        '--- calc normal colors
        m_clrSelMenuBorder = vbHighlight
        m_clrSelMenuBack = pvAlphaBlend(vbHighlight, vbWindowBackground, 70)
        m_clrSelMenuFore = vbMenuText
        m_clrCheckBack = pvAlphaBlend(vbWindowBackground, m_clrSelMenuBack, 128)
        m_clrCheckFore = m_clrSelMenuFore
        m_clrSelCheckBack = pvAlphaBlend(pvAlphaBlend(vbHighlight, m_clrSelMenuBack, 128), m_clrSelMenuBack, 128)
        m_clrMenuBorder = vbButtonShadow
        m_clrMenuBack = pvAlphaBlend(vbButtonFace, vbWindowBackground, 214)
        m_clrMenuFore = vbWindowText
        m_clrDisabledMenuBorder = vbButtonShadow
        m_clrDisabledMenuBack = pvAlphaBlend(m_clrMenuBack, vbWindowBackground, 128)
        m_clrDisabledMenuFore = vbGrayText
        If pvIsAppearanceXpStyle Then
            m_clrMenuBarBack = GetSysColor(COLOR_MENUBAR)
          Else
            m_clrMenuBarBack = vbMenuBar
        End If
        m_clrMenuPopupBack = vbWindowBackground
    End If
    '--- calc menu item height
    With New cMemDC
        .Init
        If UseSystemFont Then
            Set .Font = .SystemMenuFont
          Else
            Set .Font = Font
        End If
        m_lTextHeight = .TextHeight("ABCH") + 7
        m_lMenuHeight = m_lTextHeight
        '--- min space for icons
        If m_lMenuHeight < BitmapSize + 7 Then
            m_lMenuHeight = BitmapSize + 7
        End If
    End With
    '--- (re)init menu
    Call pvInitMenu(m_hFormMenu, True)

End Sub

Private Sub pvGetMenuInfo(ByVal lIdx As Long, hMenu As Long, sText As String, lType As Long, bMainMenu As Boolean, lId As Long)

  Dim vSplit          As Variant

    On Error Resume Next
        vSplit = Split(m_cMenuInfo(lIdx), Chr(1))
        hMenu = vSplit(0)
        sText = vSplit(1)
        lType = vSplit(2)
        bMainMenu = vSplit(3) <> 0
        lId = vSplit(4)
    On Error GoTo 0

End Sub

Private Function pvHandleMenuChar(ByVal hMenu As Long, ByVal lChar As Long, lReturn As Long) As Boolean

  '--- function provided by Vlad 17/11/2003 to cater for NT and below Accelerator keys

  Dim lIdx    As Long
  Dim mii As MENUITEMINFO

  Dim sText As String

  Dim lType As Long
  Dim lId As Long
  Dim vSplit As Variant

    On Error GoTo EH
    Debug.Print "pvHandleMenuChar "; Hex(hMenu), Hex(lChar); Timer
    For lIdx = 0 To GetMenuItemCount(hMenu) - 1
        With mii
            '---  don't need extended  info
            .cbSize = Len(mii) - 4
            .fMask = MIIM_DATA
            Call GetMenuItemInfo(hMenu, lIdx, 1, mii)
            If .dwItemData > 0 Then
                pvGetMenuInfo .dwItemData, hMenu, sText, lType, False, lId
                If (lType And MFT_SEPARATOR) = 0 Then
                    '--- replace menu char (&) with chr$(1) and take care of escaped menu char  ("&&" ->  "&")
                    sText = Replace(Replace(Replace(sText, "&&", Chr$(2)), "&", Chr$(1)), Chr$(2), "&")
                    vSplit = Split(sText, Chr$(1))
                    '--- if accel key  specified
                    If UBound(vSplit) > 0 Then
                        '--- check for menu char w/o accel  key
                        If Len(vSplit(UBound(vSplit))) > 0 Then
                            '--- check char after & (accelerator  key)
                            If lChar = Asc(LCase(Left(vSplit(UBound(vSplit)), 1))) Then
                                lReturn = MNC_EXECUTE * &H10000 + lIdx
                                pvHandleMenuChar = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next lIdx

    Exit Function

EH:
    Debug.Print Err.Description

End Function

Private Function pvHM2Pix(ByVal Value As Double) As Double

    pvHM2Pix = Value * 1440 / 2540 / Screen.TwipsPerPixelX

End Function

Private Sub pvInitMenu(ByVal hMenu As Long, ByVal bMainMenu As Boolean)

  Dim mii             As MENUITEMINFO
  Dim lIdx            As Long
  Dim hMdiChild       As Long
  Dim sBuffer         As String
  Dim lColumnCount    As Long
  Dim bSeparator      As Boolean

  Dim tmp_hMenu       As Long
  Dim tmp_sText       As String
  Dim tmp_fType       As Long
  Dim tmp_bMainMenu   As Boolean
  Dim tmp_lId         As Long

    lColumnCount = 0

    On Error GoTo EH
    If hMenu <> 0 Then
        '--- first, forward to child MDI window
        hMdiChild = pvGetMdiChild
        If hMdiChild <> 0 Then
            If SendMessage(hMdiChild, pvInitMenuMsg, IIf(hMenu = m_hFormMenu, ucsIniMainMenu, ucsIniMenu), hMenu) = 1 Then
                Exit Sub
            End If
        End If
        '--- then process locally
        sBuffer = String(1024, 0)
        For lIdx = 0 To GetMenuItemCount(hMenu) - 1
            With mii
                '--- get item info
                If OsVersion >= &H40A Then '--- &H40A = win98 and later
                    .cbSize = Len(mii)
                    .fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_DATA Or MIIM_STRING
                  Else
                    .cbSize = Len(mii) - 4
                    .fMask = MIIM_ID Or MIIM_TYPE Or MIIM_DATA
                End If
                .dwTypeData = StrPtr(sBuffer)
                .cch = Len(sBuffer)
                Call GetMenuItemInfo(hMenu, lIdx, 1, mii)
                '--- store info (if not stored already)
                If (.fType And MFT_OWNERDRAW) = 0 Then
                    .dwItemData = pvSetMenuInfo(hMenu, Left(StrConv(sBuffer, vbUnicode), .cch), .fType, bMainMenu, .wID) '--- save hMenu
                    tmp_fType = .fType
                  Else
                    pvGetMenuInfo .dwItemData, tmp_hMenu, tmp_sText, tmp_fType, tmp_bMainMenu, tmp_lId
                End If
                '--- set ownerdrawn & itemdata, clear bitmap
                If OsVersion >= &H40A Then
                    .cbSize = Len(mii)
                    .fMask = MIIM_FTYPE Or MIIM_DATA Or MIIM_BITMAP
                    .hbmpItem = 0
                  Else
                    .cbSize = Len(mii) - 4
                    .fMask = MIIM_TYPE Or MIIM_DATA
                End If

                If (tmp_fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
                    bSeparator = True
                  Else
                    bSeparator = False
                End If

                'amended NR Fixed Menu Up/Down key problem 20/11/2003
                
                If Not bSeparator Then
                    .fType = .fType And MF_MENUBREAK
                End If
                
                .fType = (.fType And (MFT_SEPARATOR Or MFT_RIGHTJUSTIFY)) Or MFT_OWNERDRAW

                If AutoColumn > 0 And Not bMainMenu Then
                    If Not bSeparator Then
                        If lColumnCount = AutoColumn Then
                            .fType = .fType Or MF_MENUBREAK
                            lColumnCount = 0
                        End If
                        lColumnCount = lColumnCount + 1
                    End If
                End If

                Call SetMenuItemInfo(hMenu, lIdx, 1, mii)
            End With
        Next lIdx
    End If
#If WEAK_REF_CURRENTMENU Then
    CopyMemory VarPtr(g_oCurrentMenu), VarPtr(Me), 4
#Else
    Set g_oCurrentMenu = Me
#End If

    Exit Sub

EH:
    Debug.Print "Error in pvInitMenu: "; Error
    Resume Next

End Sub

Private Property Get pvInitMenuMsg() As Long

    If m_lInitMenuMsg = 0 Then
        m_lInitMenuMsg = RegisterWindowMessage("InitMenuMsg")
    End If
    pvInitMenuMsg = m_lInitMenuMsg

End Property

Private Function pvIsAppearanceXpStyle() As Boolean

  Dim lValue          As Long

    SystemParametersInfo SPI_GETFLATMENU, 0, lValue, 0
    pvIsAppearanceXpStyle = (lValue <> 0)

End Function

Private Function pvIsHightContrast() As Boolean

  Const STR_HIGH_CONTRAST As String = "High Contrast"
  Dim hc              As HIGHCONTRAST

    hc.cbSize = Len(hc)
    Call SystemParametersInfo(SPI_GETHIGHCONTRAST, Len(hc), hc, 0)
    pvIsHightContrast = (hc.dwFlags And HCF_HIGHCONTRASTON) <> 0
    If Not pvIsHightContrast Then
        '--- hack: not working for language other than english!!!
        pvIsHightContrast = (Left(pvRegGetKeyValue(HKEY_CURRENT_USER, "Control Panel\Appearance", "Current"), Len(STR_HIGH_CONTRAST)) = STR_HIGH_CONTRAST)
    End If

End Function

Private Function pvMeasureItem(ByVal lParam As Long) As Boolean

  Dim mis             As MEASUREITEMSTRUCT
  Dim mii             As MENUITEMINFO
  Dim vSplit          As Variant
  Dim hMenu           As Long
  Dim sText           As String
  Dim lType           As Long
  Dim bMainMenu       As Boolean
  Dim lId             As Long
  Dim lRight          As Long

    '--- dereference structure
    CopyMemory VarPtr(mis), lParam, Len(mis)
    If mis.CtlType = ODT_MENU Then
        '--- get menu info
        Call pvGetMenuInfo(mis.ItemData, hMenu, sText, lType, bMainMenu, lId)
        '--- if not found -> degrade to 50px width
        If mis.itemID <> lId Or hMenu = 0 Then
            mis.itemWidth = 50
            CopyMemory lParam, VarPtr(mis), Len(mis)
            Exit Function
        End If
        With New cMemDC
            .Init
            If UseSystemFont Then
                Set .Font = .SystemMenuFont
              Else
                Set .Font = Font
            End If
            If (lType And MFT_SEPARATOR) = MFT_SEPARATOR Then
                mis.ItemHeight = SEPARATOR_HEIGHT
              Else
                mis.ItemHeight = m_lMenuHeight
            End If
            '--- calc text width minus underlines
            .DrawText sText, 0, 0, lRight, 0, DT_CALCRECT Or DT_SINGLELINE
            mis.itemWidth = lRight + IIf(bMainMenu, 2, m_lMenuHeight + 8 + m_lMenuHeight)
        End With
        CopyMemory lParam, VarPtr(mis), Len(mis)
        '--- handled
        pvMeasureItem = True
    End If

End Function

Private Function pvReadPictureProperty( _
                                       oBag As PropertyBag, _
                                       sPropName As String, _
                                       Optional DefaultValue As Variant) As StdPicture

  Dim ii              As ICONINFO
  Dim hr              As Long
  Dim imgColor        As StdPicture
  Dim imgMask         As StdPicture

    If IsArray(oBag.ReadProperty(sPropName, DefaultValue)) Then
        With New PropertyBag
            .Contents = oBag.ReadProperty(sPropName, DefaultValue)
            Set imgColor = .ReadProperty("c")
            Set imgMask = .ReadProperty("m")
        End With
        ii.fIcon = 1
        ii.hbmColor = imgColor.handle
        ii.hbmMask = imgMask.handle
        With New cMemDC
            Set pvReadPictureProperty = .IconToPicture(CreateIconIndirect(ii))
        End With
      Else
        Set pvReadPictureProperty = oBag.ReadProperty(sPropName, DefaultValue)
    End If

End Function

Private Function pvRegGetKeyValue( _
                                  lKeyRoot As Long, _
                                  sKeyName As String, _
                                  sValueName As String) As String

  Dim hr              As Long
  Dim hKey            As Long
  Dim sValue          As String
  Dim lValType        As Long
  Dim lValSize        As Long

    '--- open key
    hr = RegOpenKeyEx(lKeyRoot, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    If hr <> 0 Then
        Exit Function
    End If
    '--- query value size
    lValSize = 0
    hr = RegQueryValueEx(hKey, sValueName, 0, lValType, vbNullString, lValSize)
    If hr <> 0 Then
        Call RegCloseKey(hKey)
        Exit Function
    End If
    '--- get value
    sValue = String(lValSize + 1, 0)
    lValSize = Len(sValue)
    hr = RegQueryValueEx(hKey, sValueName, 0, lValType, sValue, lValSize)
    If hr <> 0 Then
        Call RegCloseKey(hKey)
        Exit Function
    End If
    '--- close key
    Call RegCloseKey(hKey)
    '--- ret value and trim
    If lValSize > 0 Then
        If (Asc(Mid(sValue, lValSize, 1)) = 0) Then
            pvRegGetKeyValue = Left(sValue, lValSize - 1)
          Else
            pvRegGetKeyValue = Left(sValue, lValSize)
        End If
    End If

End Function

Private Sub pvRestoreMenus(ByVal hMenu As Long)

  Dim hCurMenu        As Long
  Dim sText           As String
  Dim lType           As Long
  Dim bMainMenu       As Boolean
  Dim lId             As Long
  Dim mii             As MENUITEMINFO
  Dim lIdx            As Long

    lIdx = 1
    Do While m_cMenuInfo.Count >= lIdx
        pvGetMenuInfo lIdx, hCurMenu, sText, lType, bMainMenu, lId
        If hCurMenu <> hMenu And hMenu <> 0 Then
            lIdx = lIdx + 1
          Else
            With mii
                If OsVersion >= &H40A Then
                    .cbSize = Len(mii)
                    .fMask = MIIM_STRING Or MIIM_FTYPE Or MIIM_DATA
                    .hbmpItem = 0
                  Else
                    .cbSize = Len(mii) - 4
                    .fMask = MIIM_TYPE Or MIIM_DATA
                End If
                sText = StrConv(sText, vbFromUnicode)
                .dwTypeData = StrPtr(sText)
                .cch = Len(sText)
                .fType = lType
                Call SetMenuItemInfo(hCurMenu, lId, 0, mii)
            End With
            Call m_cMenuInfo.Remove(lIdx)
        End If
    Loop

End Sub

Private Function pvSetMenuInfo(ByVal hMenu As Long, sText As String, ByVal lType As Long, ByVal bMainMenu As Boolean, ByVal lId As Long) As Long

    m_cMenuInfo.Add hMenu & Chr(1) & sText & Chr(1) & lType & Chr(1) & Abs(bMainMenu) & Chr(1) & lId
    pvSetMenuInfo = m_cMenuInfo.Count

End Function

Private Sub pvWritePictureProperty( _
                                   oBag As PropertyBag, _
                                   sPropName As String, _
                                   ByVal oPic As StdPicture, _
                                   Optional DefaultValue As Variant)

  Dim ii              As ICONINFO
  Dim hr              As Long
  Dim oMemDC          As cMemDC

    If Not oPic Is Nothing Then
        If oPic.Type = vbPicTypeIcon Then
            Set oMemDC = New cMemDC
            hr = GetIconInfo(oPic.handle, ii)
            With New PropertyBag
                Call .WriteProperty("c", oMemDC.BitmapToPicture(ii.hbmColor))
                Call .WriteProperty("m", oMemDC.BitmapToPicture(ii.hbmMask))
                Call oBag.WriteProperty(sPropName, .Contents, DefaultValue)
            End With
            Exit Sub
        End If
    End If
    '--- else
    Call oBag.WriteProperty(sPropName, oPic, DefaultValue)

End Sub

'-- Added By Gary Noble To Test Left To Right Support
Public Property Get RightToLeft() As Boolean

    RightToLeft = m_RightToLeft

End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)

    m_RightToLeft = New_RightToLeft
    PropertyChanged "RightToLeft"

End Property

Property Get SelectDisabled() As Boolean

    SelectDisabled = m_bSelectDisabled

End Property

Property Let SelectDisabled(ByVal bValue As Boolean)

    m_bSelectDisabled = bValue
    PropertyChanged

End Property

Public Sub SetBitmap(Menu As Object, Pic As StdPicture, Optional MaskColor As OLE_COLOR = MASK_COLOR)

    SetBitmapByCaption pvGetCtlName(Menu), Pic, MaskColor

End Sub

Public Sub SetBitmapByCaption(ByVal sCtlName As String, Pic As StdPicture, Optional MaskColor As OLE_COLOR = MASK_COLOR)

  Dim sKey            As String

    On Error Resume Next
        sKey = "#" & sCtlName
        m_cBmps.Remove sKey
        If Not Pic Is Nothing Then
            m_cBmps.Add Array(Pic, MaskColor, sKey), sKey
        End If
        PropertyChanged
    On Error GoTo 0

End Sub

'-- Sets The custom Properties For DrawStyle Custom
'-- Created By: Gary Noble (Phantom Man - PSC)
Public Function SetCustomAttributes(TopMenuBackColour As OLE_COLOR, TopMenuSelectedColour As OLE_COLOR, _
                                    ExpandedMenuSelectedBackColour As OLE_COLOR, TopMenuHotColour As OLE_COLOR, SelectedItemBorderColour As OLE_COLOR, _
                                    MenuBorderColour As OLE_COLOR, SideBarColour As OLE_COLOR, CheckBoxColour As OLE_COLOR, _
                                    CheckBorderColour As OLE_COLOR, GradientColourOne As OLE_COLOR, GradientColourTwo As OLE_COLOR, _
                                    TopMenuHotBorder As OLE_COLOR, SelectedForeColour As OLE_COLOR, blnUseGradient As Boolean, blnUseHeaderGradient As Boolean)

  '-- Set Custom Attributes

    m_UserSelectedMenuBorderColour = SelectedItemBorderColour
    m_UserTopMenuBackColour = TopMenuBackColour
    m_UserTopMenuSelectedColour = TopMenuSelectedColour
    m_UserSelectedMenuBackColour = ExpandedMenuSelectedBackColour
    m_UserTopMenuHotColour = TopMenuHotColour
    m_UserTopMenuHotBorderColour = TopMenuHotBorder
    m_UserMenuBorderColour = MenuBorderColour
    m_UserSideBarColour = SideBarColour
    m_UserCheckBackColour = CheckBoxColour
    m_UserCheckBorderColour = CheckBorderColour
    m_UserGradientOne = GradientColourOne
    m_UserGradientTwo = GradientColourTwo
    m_UserUseGradient = blnUseGradient
    m_UserUseTopMenuGradient = blnUseHeaderGradient
    m_UserSelectedItemForeColour = SelectedForeColour
    DrawMenuBar IIf(m_hParentHwnd <> 0, m_hParentHwnd, m_hFormHwnd)

End Function

Public Property Get UserCheckBackColour() As Variant

    UserCheckBackColour = m_UserCheckBackColour

End Property

Public Property Let UserCheckBackColour(ByVal New_UserCheckBackColour As Variant)

    m_UserCheckBackColour = New_UserCheckBackColour
    PropertyChanged "UserCheckBackColour"

End Property

Public Property Get UserCheckBorderColour() As Variant

    UserCheckBorderColour = m_UserCheckBorderColour

End Property

Public Property Let UserCheckBorderColour(ByVal New_UserCheckBorderColour As Variant)

    m_UserCheckBorderColour = New_UserCheckBorderColour
    PropertyChanged "UserCheckBorderColour"

End Property

'==============================================================================
' Base class events
'==============================================================================

Private Sub UserControl_Initialize()

    Set m_cMenuSubclass = New Collection
    Set m_cBmps = New Collection
    Set m_cMemDC = New Collection
    Set m_cMenuInfo = New Collection
    Set m_oFont = New StdFont
    If g_oMenuHookImpl Is Nothing Then
        Set g_oMenuHookImpl = New cMenuHook
    End If
#If DebugMode Then
    DebugInit m_sDebugID, MODULE_NAME
#End If

End Sub

Private Sub UserControl_InitProperties()

    SelectDisabled = DEF_SELECTDISABLED
    BitmapSize = DEF_BITMAPSIZE
    UseSystemFont = DEF_USESYSTEMFONT
    Set Font = DEF_FONT
    Init UserControl.ContainerHwnd
    m_DrawStyle = m_def_DrawStyle
    m_DisplayShadow = m_def_DisplayShadow
    m_MenuDrawStyle = m_def_MenuDrawStyle
    m_UserSelectedMenuBackColour = m_def_UserSelectedMenuBackColour
    m_UserSelectedMenuBorderColour = m_def_UserSelectedMenuBorderColour
    m_UserTopMenuBackColour = m_def_UserTopMenuBackColour
    m_UserTopMenuSelectedColour = m_def_UserTopMenuSelectedColour
    m_UserTopMenuHotColour = m_def_UserTopMenuHotColour
    m_UserTopMenuHotBorderColour = m_def_UserTopMenuHotBorderColour
    m_UserMenuBorderColour = m_def_UserMenuBorderColour
    m_UserCheckBackColour = m_def_UserCheckBackColour
    m_UserCheckBorderColour = m_def_UserCheckBorderColour
    m_UserGradientOne = m_def_UserGradientOne
    m_UserGradientTwo = m_def_UserGradientTwo
    m_UserUseGradient = m_def_UserUseGradient
    m_UserUseTopMenuGradient = m_def_UserUseTopMenuGradient
    m_UserSelectedItemForeColour = m_def_UserSelectedItemForeColour
    m_UserSideBarColour = m_def_UserSideBarColour
    m_RightToLeft = m_def_RightToLeft

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Dim lIdx            As Long
  Dim vElem           As Variant

    On Error Resume Next
        ReDim vElem(0 To 2)
        With PropBag
            SelectDisabled = .ReadProperty("SelectDisabled", DEF_SELECTDISABLED)
            BitmapSize = .ReadProperty("BitmapSize", DEF_BITMAPSIZE)
            For lIdx = 1 To .ReadProperty("BmpCount", 0)
                Set vElem(0) = pvReadPictureProperty(PropBag, "Bmp:" & lIdx, Nothing)
                vElem(1) = .ReadProperty("Mask:" & lIdx, 0)
                vElem(2) = .ReadProperty("Key:" & lIdx, "#" & lIdx)
                m_cBmps.Add vElem, vElem(2)
            Next lIdx
            UseSystemFont = .ReadProperty("UseSystemFont", DEF_USESYSTEMFONT)
            Set Font = .ReadProperty("Font", DEF_FONT)
        End With
        m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)

        m_UserTopMenuBackColour = 16761765
        m_UserTopMenuSelectedColour = 16769990
        m_UserSelectedMenuBackColour = 13040639
        m_UserTopMenuHotColour = &H80FF&
        m_UserMenuBorderColour = 5910544
        m_UserSideBarColour = vbBlue
        m_UserCheckBackColour = 8108783
        m_UserCheckBorderColour = m_UserMenuBorderColour
        m_UserGradientOne = vbWhite
        m_UserGradientTwo = 16761765
        m_UserUseGradient = True
        m_UserUseTopMenuGradient = True

        m_UserTopMenuHotBorderColour = vbBlue
        m_MenuDrawStyle = PropBag.ReadProperty("MenuDrawStyle", m_def_MenuDrawStyle)
        m_DisplayShadow = PropBag.ReadProperty("DisplayShadow", m_def_DisplayShadow)
        m_AutoColumn = PropBag.ReadProperty("AutoColumn", m_def_AutoColumn)
        Init UserControl.ContainerHwnd

    On Error GoTo 0
    m_UserSelectedMenuBackColour = PropBag.ReadProperty("UserSelectedMenuBackColour", m_def_UserSelectedMenuBackColour)
    m_UserSelectedMenuBorderColour = PropBag.ReadProperty("UserSelectedMenuBorderColour", m_def_UserSelectedMenuBorderColour)
    m_UserTopMenuBackColour = PropBag.ReadProperty("UserTopMenuBackColour", m_def_UserTopMenuBackColour)
    m_UserTopMenuSelectedColour = PropBag.ReadProperty("UserTopMenuSelectedColour", m_def_UserTopMenuSelectedColour)
    m_UserTopMenuHotColour = PropBag.ReadProperty("UserTopMenuHotColour", m_def_UserTopMenuHotColour)
    m_UserTopMenuHotBorderColour = PropBag.ReadProperty("UserTopMenuHotBorderColour", m_def_UserTopMenuHotBorderColour)
    m_UserMenuBorderColour = PropBag.ReadProperty("UserMenuBorderColour", m_def_UserMenuBorderColour)
    m_UserCheckBackColour = PropBag.ReadProperty("UserCheckBackColour", m_def_UserCheckBackColour)
    m_UserCheckBorderColour = PropBag.ReadProperty("UserCheckBorderColour", m_def_UserCheckBorderColour)
    m_UserGradientOne = PropBag.ReadProperty("UserGradientOne", m_def_UserGradientOne)
    m_UserGradientTwo = PropBag.ReadProperty("UserGradientTwo", m_def_UserGradientTwo)
    m_UserUseGradient = PropBag.ReadProperty("UserUseGradient", m_def_UserUseGradient)
    m_UserUseTopMenuGradient = PropBag.ReadProperty("UserUseTopMenuGradient", m_def_UserUseTopMenuGradient)
    m_UserSelectedItemForeColour = PropBag.ReadProperty("UserSelectedItemForeColour", m_def_UserSelectedItemForeColour)
    m_UserSideBarColour = PropBag.ReadProperty("UserSideBarColour", m_def_UserSideBarColour)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", m_def_RightToLeft)

End Sub

Private Sub UserControl_Resize()

    Width = ScaleX(32 + m_lEdgeWidth, vbPixels)
    Height = ScaleY(32 + m_lEdgeWidth, vbPixels)

End Sub

Private Sub UserControl_Terminate()

    If g_oCurrentMenu Is Me Then
#If WEAK_REF_CURRENTMENU Then
        CopyMemory VarPtr(g_oCurrentMenu), VarPtr(0), 4
#Else
        Set g_oCurrentMenu = Nothing
#End If
    End If
#If DebugMode Then
    DebugTerm m_sDebugID
#End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Dim lIdx            As Long

    On Error Resume Next
        With PropBag
            Call .WriteProperty("SelectDisabled", SelectDisabled, DEF_SELECTDISABLED)
            Call .WriteProperty("BitmapSize", BitmapSize, DEF_BITMAPSIZE)
            Call .WriteProperty("BmpCount", m_cBmps.Count)
            For lIdx = 1 To m_cBmps.Count
                Call pvWritePictureProperty(PropBag, "Bmp:" & lIdx, m_cBmps(lIdx)(0), Nothing)
                Call .WriteProperty("Mask:" & lIdx, m_cBmps(lIdx)(1), 0)
                Call .WriteProperty("Key:" & lIdx, m_cBmps(lIdx)(2), "#" & lIdx)
            Next lIdx
            Call .WriteProperty("UseSystemFont", UseSystemFont, DEF_USESYSTEMFONT)
            Call .WriteProperty("Font", Font, DEF_FONT)
        End With
        Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)
        Call PropBag.WriteProperty("DisplayShadow", m_DisplayShadow, m_def_DisplayShadow)
        Call PropBag.WriteProperty("MenuDrawStyle", m_MenuDrawStyle, m_def_MenuDrawStyle)
        Call PropBag.WriteProperty("AutoColumn", m_AutoColumn, m_def_AutoColumn)

    On Error GoTo 0
    Call PropBag.WriteProperty("UserSelectedMenuBackColour", m_UserSelectedMenuBackColour, m_def_UserSelectedMenuBackColour)
    Call PropBag.WriteProperty("UserSelectedMenuBorderColour", m_UserSelectedMenuBorderColour, m_def_UserSelectedMenuBorderColour)
    Call PropBag.WriteProperty("UserTopMenuBackColour", m_UserTopMenuBackColour, m_def_UserTopMenuBackColour)
    Call PropBag.WriteProperty("UserTopMenuSelectedColour", m_UserTopMenuSelectedColour, m_def_UserTopMenuSelectedColour)
    Call PropBag.WriteProperty("UserTopMenuHotColour", m_UserTopMenuHotColour, m_def_UserTopMenuHotColour)
    Call PropBag.WriteProperty("UserTopMenuHotBorderColour", m_UserTopMenuHotBorderColour, m_def_UserTopMenuHotBorderColour)
    Call PropBag.WriteProperty("UserMenuBorderColour", m_UserMenuBorderColour, m_def_UserMenuBorderColour)
    Call PropBag.WriteProperty("UserCheckBackColour", m_UserCheckBackColour, m_def_UserCheckBackColour)
    Call PropBag.WriteProperty("UserCheckBorderColour", m_UserCheckBorderColour, m_def_UserCheckBorderColour)
    Call PropBag.WriteProperty("UserGradientOne", m_UserGradientOne, m_def_UserGradientOne)
    Call PropBag.WriteProperty("UserGradientTwo", m_UserGradientTwo, m_def_UserGradientTwo)
    Call PropBag.WriteProperty("UserUseGradient", m_UserUseGradient, m_def_UserUseGradient)
    Call PropBag.WriteProperty("UserUseTopMenuGradient", m_UserUseTopMenuGradient, m_def_UserUseTopMenuGradient)
    Call PropBag.WriteProperty("UserSelectedItemForeColour", m_UserSelectedItemForeColour, m_def_UserSelectedItemForeColour)
    Call PropBag.WriteProperty("UserSideBarColour", m_UserSideBarColour, m_def_UserSideBarColour)
    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)

End Sub

Public Property Let UserGradientOne(ByVal New_UserGradientOne As OLE_COLOR)

    m_UserGradientOne = New_UserGradientOne
    PropertyChanged "UserGradientOne"

End Property

Public Property Get UserGradientOne() As OLE_COLOR

    UserGradientOne = m_UserGradientOne

End Property

Public Property Get UserGradientTwo() As OLE_COLOR

    UserGradientTwo = m_UserGradientTwo

End Property

Public Property Let UserGradientTwo(ByVal New_UserGradientTwo As OLE_COLOR)

    m_UserGradientTwo = New_UserGradientTwo
    PropertyChanged "UserGradientTwo"

End Property

Public Property Let UserMenuBorderColour(ByVal New_UserMenuBorderColour As OLE_COLOR)

    m_UserMenuBorderColour = New_UserMenuBorderColour
    PropertyChanged "UserMenuBorderColour"

End Property

Public Property Get UserMenuBorderColour() As OLE_COLOR

    UserMenuBorderColour = m_UserMenuBorderColour

End Property

Public Property Get UserSelectedItemForeColour() As OLE_COLOR

    UserSelectedItemForeColour = m_UserSelectedItemForeColour

End Property

Public Property Let UserSelectedItemForeColour(ByVal New_UserSelectedItemForeColour As OLE_COLOR)

    m_UserSelectedItemForeColour = New_UserSelectedItemForeColour
    PropertyChanged "UserSelectedItemForeColour"

End Property

Public Property Get UserSelectedMenuBackColour() As OLE_COLOR

    UserSelectedMenuBackColour = m_UserSelectedMenuBackColour

End Property

Public Property Let UserSelectedMenuBackColour(ByVal New_UserSelectedMenuBackColour As OLE_COLOR)

    m_UserSelectedMenuBackColour = New_UserSelectedMenuBackColour
    PropertyChanged "UserSelectedMenuBackColour"

End Property

Public Property Get UserSelectedMenuBorderColour() As OLE_COLOR

    UserSelectedMenuBorderColour = m_UserSelectedMenuBorderColour

End Property

Public Property Let UserSelectedMenuBorderColour(ByVal New_UserSelectedMenuBorderColour As OLE_COLOR)

    m_UserSelectedMenuBorderColour = New_UserSelectedMenuBorderColour
    PropertyChanged "UserSelectedMenuBorderColour"

End Property

Public Property Get UserSideBarColour() As OLE_COLOR

    UserSideBarColour = m_UserSideBarColour

End Property

Public Property Let UserSideBarColour(ByVal New_UserSideBarColour As OLE_COLOR)

    m_UserSideBarColour = New_UserSideBarColour
    PropertyChanged "UserSideBarColour"

End Property

Public Property Let UserTopMenuBackColour(ByVal New_UserTopMenuBackColour As OLE_COLOR)

    m_UserTopMenuBackColour = New_UserTopMenuBackColour
    PropertyChanged "UserTopMenuBackColour"

End Property

Public Property Get UserTopMenuBackColour() As OLE_COLOR

    UserTopMenuBackColour = m_UserTopMenuBackColour

End Property

Public Property Let UserTopMenuHotBorderColour(ByVal New_UserTopMenuHotBorderColour As OLE_COLOR)

    m_UserTopMenuHotBorderColour = New_UserTopMenuHotBorderColour
    PropertyChanged "UserTopMenuHotBorderColour"

End Property

Public Property Get UserTopMenuHotBorderColour() As OLE_COLOR

    UserTopMenuHotBorderColour = m_UserTopMenuHotBorderColour

End Property

Public Property Let UserTopMenuHotColour(ByVal New_UserTopMenuHotColour As OLE_COLOR)

    m_UserTopMenuHotColour = New_UserTopMenuHotColour
    PropertyChanged "UserTopMenuHotColour"

End Property

Public Property Get UserTopMenuHotColour() As OLE_COLOR

    UserTopMenuHotColour = m_UserTopMenuHotColour

End Property

Public Property Let UserTopMenuSelectedColour(ByVal New_UserTopMenuSelectedColour As OLE_COLOR)

    m_UserTopMenuSelectedColour = New_UserTopMenuSelectedColour
    PropertyChanged "UserTopMenuSelectedColour"

End Property

Public Property Get UserTopMenuSelectedColour() As OLE_COLOR

    UserTopMenuSelectedColour = m_UserTopMenuSelectedColour

End Property

Public Property Let UserUseGradient(ByVal New_UserUseGradient As Boolean)

    m_UserUseGradient = New_UserUseGradient
    PropertyChanged "UserUseGradient"

End Property

Public Property Get UserUseGradient() As Boolean

    UserUseGradient = m_UserUseGradient

End Property

Public Property Let UserUseTopMenuGradient(ByVal New_UserUseTopMenuGradient As Boolean)

    m_UserUseTopMenuGradient = New_UserUseTopMenuGradient
    PropertyChanged "UserUseTopMenuGradient"

End Property

Public Property Get UserUseTopMenuGradient() As Boolean

    UserUseTopMenuGradient = m_UserUseTopMenuGradient

End Property

Property Get UseSystemFont() As Boolean

    UseSystemFont = m_bUseSystemFont

End Property

Property Let UseSystemFont(ByVal bValue As Boolean)

    m_bUseSystemFont = bValue
    pvGetMeasures
    PropertyChanged

End Property

