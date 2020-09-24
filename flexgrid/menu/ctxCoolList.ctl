VERSION 5.00
Begin VB.UserControl ctxCoolList 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   KeyPreview      =   -1  'True
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   102
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   1620
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.VScrollBar sbScroll 
      Height          =   1125
      Left            =   1215
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picDC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   2
      Top             =   0
      Width           =   1170
   End
End
Attribute VB_Name = "ctxCoolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
' UserControl:   ctxCoolList.ctl v1.3 (Long supp.)
' Author:        Carles P.V.
' Dependencies:  cLongScroll
' Last revision: 07.04.2003
'================================================================
'-- Ammended Original ucCoollist 5-11-2003 Gary Noble
'-- 5-11-2003 : Added ItemData And Pic To tItem Array
'--             Added SetTab Parameter - To Incorporate A Tabbed
'--                   Style List Box
'--             Added Draw Style: IsHookMenu Style Or Default
'================================================================
Option Explicit

Implements ISubclassingSink
Dim m_oSubclass As cSubclassingThunk
Private Const WM_MOUSEWHEEL As Integer = &H20A
'-- API:

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal dx As Long, ByVal dy As Long) As Long

Private aTabs() As Long

Public Enum EnumStyle
    DefaultStyle = 1
    IsHookMenuStyle = 2
End Enum

Private Type TRIVERTEX
    x     As Long
    y     As Long
    R     As Integer
    G     As Integer
    B     As Integer
    Alpha As Integer
End Type

Private Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const SM_CXVSCROLL         As Long = &H2
Private Const PS_SOLID             As Long = &H0
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H1
Private Const DT_LEFT              As Long = &H0
Private Const DT_CENTER            As Long = &H1
Private Const DT_RIGHT             As Long = &H2
Private Const DT_VCENTER           As Long = &H4
Private Const DT_WORDBREAK         As Long = &H10
Private Const DT_SINGLELINE        As Long = &H20

'-- Public enums.:

Public Enum AlignmentCts
    [AlignLeft]
    [AlignCenter]
    [AlignRight]
End Enum

Public Enum AppearanceCts
    [Flat]
    [3D]
End Enum

Public Enum BorderStyleCts
    [None]
    [Fixed Single]
End Enum

Public Enum OrderTypeCts
    [Ascendent]
    [Descendent]
End Enum

Public Enum SelectModeCts
    [Single]
    [Multiple]
End Enum

Public Enum SelectModeStyleCts
    [Standard]
    [Dither]
    [Gradient_V]
    [Gradient_H]
    [Box]
    [Underline]
    [byPicture]
End Enum

'-- Private Constants:

Private Const m_def_Appearance = 1
Private Const m_def_Alignment = DT_LEFT
Private Const m_def_BackNormal = vbWindowBackground
Private Const m_def_BackSelected = vbHighlight
Private Const m_def_BackSelectedG1 = vbHighlight
Private Const m_def_BackSelectedG2 = vbWindowBackground
Private Const m_def_BorderStyle = 1
Private Const m_def_BoxBorder = vbHighlightText
Private Const m_def_BoxOffset = 1
Private Const m_def_BoxRadius = 0
Private Const m_def_Focus = -1
Private Const m_def_FontNormal = vbWindowText
Private Const m_def_FontSelected = vbHighlightText
Private Const m_def_HoverSelection = 0
Private Const m_def_ItemHeightAuto = -1
Private Const m_def_ItemOffset = 0
Private Const m_def_ItemTextLeft = 2
Private Const m_def_OrderType = 0
Private Const m_def_SelectMode = 0
Private Const m_def_SelectModeStyle = 0
Private Const m_def_WordWrap = -1
Private Const m_def_DrawStyle = EnumStyle.DefaultStyle
'-- Property Variables:

Private m_Alignment        As AlignmentCts
Private m_Apeareance       As AppearanceCts
Private m_BackNormal       As OLE_COLOR
Private m_BackSelected     As OLE_COLOR
Private m_BackSelectedG1   As OLE_COLOR
Private m_BackSelectedG2   As OLE_COLOR
Private m_BoxBorder        As OLE_COLOR
Private m_BoxOffset        As Long
Private m_BoxRadius        As Long
Private m_Focus            As Boolean
Private m_FontNormal       As OLE_COLOR
Private m_FontSelected     As OLE_COLOR
Private m_HoverSelection   As Boolean
Private m_ItemHeight       As Long
Private m_ItemHeightAuto   As Boolean
Private m_ItemOffset       As Long
Private m_ItemTextLeft     As Long
Private m_ListIndex        As Long
Private m_OrderType        As OrderTypeCts
Private m_SelectionPicture As Picture
Private m_SelectMode       As SelectModeCts
Private m_SelectModeStyle  As SelectModeStyleCts
Private m_TopIndex         As Long
Private m_WordWrap         As Boolean

'-- Private Types:

Private Type tItem
    Text         As String
    Icon         As Integer
    IconSelected As Integer
    ItemData     As Integer
    Pic          As StdPicture
    Seperator    As Boolean
End Type

'-- Private Variables:

Private m_List()          As tItem    ' List array of items (Text, icons)
Private m_Selected()      As Boolean  ' List array of items (Selected/Unselected)
Private m_nItems          As Long     ' Number of Items

Private m_LastBar         As Long     ' Last scroll bar value
Private m_LastItem        As Long     ' Last Selected item
Private m_LastY           As Single   ' Last Y value [pixels] (prevents item repaint)
Private m_AnchorItemState As Boolean  ' Anchor item value (multiple selection)
Private m_EnsureVisible   As Boolean  ' Ensure visible last m_Selected item (ListIndex)

Private m_ItemRct()       As RECT2    ' Item rectangles
Private m_TextRct()       As RECT2    ' Item text rectangles
Private m_IconPt()        As POINTAPI ' Item icon positions

Private m_tmpItemHeight   As Long     ' Item height [pixels]
Private m_VisibleRows     As Long     ' Visible rows in control area
Private m_Scrolling       As Boolean  ' Scrolling by mouse
Private m_ScrollingY      As Long     ' Y Scrolling coordinate flag (scroll speed = f(Y))
Private m_HasFocus        As Boolean  ' Control has focus
Private m_Resizing        As Boolean  ' Prevent repaints when Resizing

Private m_lpImgList       As Object   ' Will point to ImageList control
Private m_ILScale         As Long     ' ImageList parent scale mode
Private m_ColorBack       As Long     ' Back color [Normal]
Private m_ColorBackSel    As Long     ' Back color [Selected]
Private m_ColorFont       As Long     ' Font color [Normal]
Private m_ColorFontSel    As Long     ' Font color [Selected]
Private m_ColorGradient1  As RGB      ' Gradient color from [Selected]
Private m_ColorGradient2  As RGB      ' Gradient color  to  [Selected]
Private m_ColorBox        As Long     ' Box border color

'-- Private Objects:

Private WithEvents m_Font As StdFont  ' Font object and Long scroll class wrapper
Attribute m_Font.VB_VarHelpID = -1
Private WithEvents m_olSB As CLongScroll
Attribute m_olSB.VB_VarHelpID = -1

'-- Event declarations:

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ListIndexChange()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Scroll()
Public Event TopIndexChange()
Private m_DrawStyle As EnumStyle

'-- AddItem
'-- 0 , ... , n-1 [n = ListCount]

Public Sub AddItem(ByVal Text As Variant, _
                   Optional ByVal Icon As Long, _
                   Optional ByVal IconSelected As Long, _
                   Optional lItemData As Long, _
                   Optional Pic As StdPicture)

  Dim vt As Variant

    vt = Split(Text, vbTab)

    '-- Add ('modify') last array item
    With m_List(m_nItems)
        .Text = CStr(Text)
        .Icon = Icon
        .IconSelected = IconSelected
        Set .Pic = Pic
        .ItemData = lItemData

        If UBound(vt) >= 1 Then
            If vt(1) = "-" Then
                .Seperator = True
              Else
                .Seperator = False
            End If
        End If
    End With
    '-- Increase items count
    m_nItems = m_nItems + 1

    '-- Redim. arrays
    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    '-- Readjust scroll bar and redraw item [?]
    pvReadjustScrollBar
    If (m_nItems < m_VisibleRows + 1) Then
        pvDrawItem (m_nItems - 1)
    End If

End Sub

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)

    m_Alignment = New_Alignment
    picDC_Paint

End Property

'========================================================================================
' Properties
'========================================================================================

'-- Alignment
Public Property Get Alignment() As AlignmentCts

    Alignment = m_Alignment

End Property

'-- Appearance
Public Property Get Appearance() As AppearanceCts

    Appearance = UserControl.Appearance

End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)

    UserControl.Appearance() = New_Appearance

End Property

'-- BackNormal
Public Property Get BackNormal() As OLE_COLOR

    BackNormal = m_BackNormal

End Property

Public Property Let BackNormal(ByVal New_BackNormal As OLE_COLOR)

    m_BackNormal = New_BackNormal
    m_ColorBack = pvGetLongColor(m_BackNormal)
    picDC.BackColor = m_ColorBack
    picDC_Paint

End Property

'-- BackSelected
Public Property Get BackSelected() As OLE_COLOR

    BackSelected = m_BackSelected

End Property

Public Property Let BackSelected(ByVal New_BackSelected As OLE_COLOR)

    m_BackSelected = New_BackSelected
    m_ColorBackSel = pvGetLongColor(m_BackSelected)
    picDC_Paint

End Property

'-- BackSelectedG1
Public Property Get BackSelectedG1() As OLE_COLOR

    BackSelectedG1 = m_BackSelectedG1

End Property

Public Property Let BackSelectedG1(ByVal New_BackSelectedG1 As OLE_COLOR)

    m_BackSelectedG1 = New_BackSelectedG1
    m_ColorGradient1 = pvGetRGBColors(pvGetLongColor(m_BackSelectedG1))
    picDC_Paint

End Property

'-- BackSelectedG2
Public Property Get BackSelectedG2() As OLE_COLOR

    BackSelectedG2 = m_BackSelectedG2

End Property

Public Property Let BackSelectedG2(ByVal New_BackSelectedG2 As OLE_COLOR)

    m_BackSelectedG2 = New_BackSelectedG2
    m_ColorGradient2 = pvGetRGBColors(pvGetLongColor(m_BackSelectedG2))
    picDC_Paint

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)

    UserControl.BorderStyle() = New_BorderStyle

End Property

'-- BorderStyle
Public Property Get BorderStyle() As BorderStyleCts

    BorderStyle = UserControl.BorderStyle

End Property

'-- BoxBorder
Public Property Get BoxBorder() As OLE_COLOR

    BoxBorder = m_BoxBorder

End Property

Public Property Let BoxBorder(ByVal New_BoxBorder As OLE_COLOR)

    m_BoxBorder = New_BoxBorder
    m_ColorBox = pvGetLongColor(m_BoxBorder)
    picDC_Paint

End Property

'-- BoxOffset
Public Property Get BoxOffset() As Long

    BoxOffset = m_BoxOffset

End Property

Public Property Let BoxOffset(ByVal New_BoxOffset As Long)

    If (New_BoxOffset <= m_tmpItemHeight \ 2) Then
        m_BoxOffset = New_BoxOffset
    End If
    picDC_Paint

End Property

Public Property Let BoxRadius(ByVal New_BoxRadius As Long)

    m_BoxRadius = New_BoxRadius
    picDC_Paint

End Property

'-- BoxRadius
Public Property Get BoxRadius() As Long

    BoxRadius = m_BoxRadius

End Property

'-- Clear
Public Sub Clear()

  '-- Hide scroll bar

    sbScroll.Visible = 0
    m_olSB.Max = 0

    '-- Clear and resize DC area
    picDC.Cls
    picDC.Move 0, 0, ScaleWidth, ScaleHeight

    '-- Reset item arrays
    ReDim m_List(0)
    ReDim m_Selected(0)
    m_nItems = 0

    '-- Reset item indexes
    m_LastItem = -1
    m_ListIndex = -1
    m_TopIndex = -1

End Sub

'-- Added By Gary Noble 5-11-2003
Public Property Get DrawStyle() As EnumStyle

    DrawStyle = m_DrawStyle

End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As EnumStyle)

    m_DrawStyle = New_DrawStyle
    PropertyChanged "DrawStyle"

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    sbScroll.Enabled = New_Enabled

End Property

'-- Enabled
Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property

'-- EndEdit
Public Sub EndEdit(Optional ByVal Modify As Boolean = 0)

    If (Modify) Then txtEdit_KeyPress 13 Else txtEdit_LostFocus

End Sub

'-- FindFirst
Public Function FindFirst(ByVal FindString As String, _
                          Optional ByVal StartIndex As Long = 0, _
                          Optional ByVal StartWith As Boolean = 0) As Long

  Dim lIdx As Long

    '-- No items
    If (m_nItems = 0) Then Err.Raise 381

    '-- Find first match
    For lIdx = StartIndex To m_nItems
        If (StartWith) Then
            If (InStr(1, LCase(m_List(lIdx).Text), LCase(FindString)) = 1) Then FindFirst = lIdx: Exit Function
          Else
            If (InStr(1, LCase(m_List(lIdx).Text), LCase(FindString)) > 1) Then FindFirst = lIdx: Exit Function
        End If
    Next lIdx

    '-- FindString not found
    FindFirst = -1

End Function

'-- Focus
Public Property Get Focus() As Boolean

    Focus = m_Focus

End Property

Public Property Let Focus(ByVal New_Focus As Boolean)

    m_Focus = New_Focus
    If (New_Focus) Then
        pvDrawFocus m_ListIndex
      Else
        pvDrawItem m_ListIndex
    End If

End Property

'-- Font
Public Property Get Font() As Font

    Set Font = m_Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    picDC_Paint

End Property

Public Property Let FontNormal(ByVal New_FontNormal As OLE_COLOR)

    m_FontNormal = New_FontNormal
    m_ColorFont = pvGetLongColor(m_FontNormal)
    SetTextColor picDC.hDC, m_ColorFont
    picDC_Paint

End Property

'-- FontNormal
Public Property Get FontNormal() As OLE_COLOR

    FontNormal = m_FontNormal

End Property

Public Property Let FontSelected(ByVal New_FontSelected As OLE_COLOR)

    m_FontSelected = New_FontSelected
    m_ColorFontSel = pvGetLongColor(m_FontSelected)
    picDC_Paint

End Property

'-- FontSelected
Public Property Get FontSelected() As OLE_COLOR

    FontSelected = m_FontSelected

End Property

'-- HoverSelection
Public Property Get HoverSelection() As Boolean

    HoverSelection = m_HoverSelection

End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)

    m_HoverSelection = New_HoverSelection
    pvDrawItem m_ListIndex
    pvDrawFocus m_ListIndex

End Property

'-- InsertItem
Public Sub InsertItem(ByVal Index As Long, _
                      ByVal Text As Variant, _
                      Optional ByVal Icon As Long, _
                      Optional ByVal IconSelected As Long, _
                      Optional lItemData As Long, _
                      Optional Pic As StdPicture)

  Dim lIdx As Long

    '-- No items or out of bounds
    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    '-- Increase item count and redim. arrays
    m_nItems = m_nItems + 1
    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    '-- Move next items
    For lIdx = m_nItems - 1 To Index Step -1
        m_List(lIdx + 1) = m_List(lIdx)
        m_Selected(lIdx + 1) = m_Selected(lIdx)
    Next lIdx

    '-- 'Insert'
    m_List(Index).Text = CStr(Text)
    m_List(Index).Icon = Icon
    m_List(Index).IconSelected = IconSelected
    m_List(Index).ItemData = lItemData
    Set m_List(Index).Pic = Pic
    m_Selected(Index) = 0

    '-- Readjust scroll bar and refresh list
    pvReadjustScrollBar
    m_EnsureVisible = 0
    If (m_ListIndex > -1 And Index <= m_ListIndex) Then
        ListIndex = m_ListIndex + 1
    End If
    picDC_Paint

End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

  '


End Sub

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)

    On Error Resume Next

        '    If hwnd = hwnd Then
        If uMsg = WM_MOUSEWHEEL Then
            If Me.ListCount > 0 And Me.ListIndex <= m_nItems - 1 Then
                If wParam < 0 Then

                    If Me.ListIndex >= Me.ListCount - 1 Then
                        Me.ListIndex = Me.ListCount - 1
                      Else
                        If Me.ListIndex + 5 > Me.ListCount - 1 Then
                            Me.ListIndex = ListCount - 1
                          Else
                            Me.ListIndex = Me.ListIndex + 5
                        End If
                        pvDrawList
                        RaiseEvent Click
                    End If

                  Else
                    If Me.ListIndex <= 0 Then
                        Me.ListIndex = 0
                      Else
                        If Me.ListIndex - 5 < 0 Then
                            Me.ListIndex = 0
                          Else
                            Me.ListIndex = Me.ListIndex - 5
                        End If
                        pvDrawList
                        RaiseEvent Click
                    End If
                End If

            End If
            bHandled = False
            lReturn = 0
        End If
        '   End If

    On Error GoTo 0

End Sub

Public Property Let ItemData(ByVal Index As Long, ByVal Data As Integer)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).ItemData = Data
    pvDrawItem Index
    pvDrawFocus m_ListIndex

End Property

'-- ItemData
Public Property Get ItemData(ByVal Index As Long) As Integer

    On Error Resume Next

        If (m_nItems = 0 Or Index > m_nItems) Then Index = m_nItems
        ItemData = m_List(Index).ItemData

End Property

'-- ItemHeight
Public Property Get ItemHeight() As Long

    ItemHeight = m_ItemHeight

End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Long)

    m_ItemHeight = New_ItemHeight
    UserControl_Resize
    picDC_Paint

End Property

Public Property Let ItemHeightAuto(ByVal New_ItemHeightAuto As Boolean)

    m_ItemHeightAuto = New_ItemHeightAuto
    UserControl_Resize
    picDC_Paint

End Property

'-- ItemHeightAuto
Public Property Get ItemHeightAuto() As Boolean

    ItemHeightAuto = m_ItemHeightAuto

End Property

Public Property Let ItemIcon(ByVal Index As Long, ByVal Data As Long)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).Icon = Data
    pvDrawItem Index
    pvDrawFocus m_ListIndex

End Property

'-- ItemIcon
Public Property Get ItemIcon(ByVal Index As Long) As Long

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemIcon = m_List(Index).Icon

End Property

'-- ItemIconSelected
Public Property Get ItemIconSelected(ByVal Index As Long) As Long

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemIconSelected = m_List(Index).IconSelected

End Property

Public Property Let ItemIconSelected(ByVal Index As Long, ByVal Data As Long)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).IconSelected = Data
    pvDrawItem Index
    pvDrawFocus m_ListIndex

End Property

Public Property Let ItemOffset(ByVal New_ItemOffset As Long)

    If (New_ItemOffset <= m_tmpItemHeight) Then
        m_ItemOffset = New_ItemOffset
    End If
    pvCalculateRects
    If (sbScroll.Visible) Then pvReadjustRects sbScroll.Width
    picDC_Paint

End Property

'-- ItemOffset
Public Property Get ItemOffset() As Long

    ItemOffset = m_ItemOffset

End Property

'-- ItemSelected
Public Property Get ItemSelected(ByVal Index As Long) As Boolean

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemSelected = m_Selected(Index)

End Property

Public Property Let ItemSelected(ByVal Index As Long, ByVal Data As Boolean)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    Select Case Data
      Case -1
        If (m_SelectMode = [Single]) Then
            ListIndex = Index
          Else
            m_Selected(Index) = -1
            pvDrawItem Index
            If (Index = m_ListIndex) Then pvDrawFocus Index
        End If
      Case 0
        If (Not (m_SelectMode = [Single])) Then
            m_Selected(Index) = 0
            pvDrawItem Index
            If (Index = m_ListIndex) Then pvDrawFocus Index
        End If
    End Select

End Property

'-- Added on 07.02.2002:

Public Property Get ItemSeperator(ByVal Index As Long) As Boolean

    If (m_nItems = 0 Or Index > m_nItems) Then Index = m_nItems
    ItemSeperator = m_List(Index).Seperator

End Property

'-- ItemText
Public Property Get ItemText(ByVal Index As Long) As String

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemText = m_List(Index).Text

End Property

Public Property Let ItemText(ByVal Index As Long, ByVal Data As String)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).Text = CStr(Data)
    pvDrawItem Index
    pvDrawFocus m_ListIndex

End Property

'-- ItemTextLeft
Public Property Get ItemTextLeft() As Long

    ItemTextLeft = m_ItemTextLeft

End Property

Public Property Let ItemTextLeft(ByVal New_ItemTextLeft As Long)

    m_ItemTextLeft = New_ItemTextLeft
    pvCalculateRects
    If (sbScroll.Visible) Then pvReadjustRects sbScroll.Width
    picDC_Paint

End Property

'-- <ListCount>
Public Property Get ListCount() As Long

    ListCount = m_nItems

End Property

'-- ListIndex
Public Property Get ListIndex() As Long

    ListIndex = m_ListIndex

End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)

    On Error Resume Next

        If (New_ListIndex < -1 Or New_ListIndex > m_nItems - 1) Then Err.Raise 380

        If (txtEdit.Visible) Then txtEdit_LostFocus

        If (New_ListIndex < 0 Or m_nItems = 0) Then
            m_ListIndex = -1
            m_LastY = -1
          Else
            m_ListIndex = New_ListIndex
        End If

        '-- Unselect last / Select actual [Single selection mode]
        If (m_SelectMode = [Single]) Then
            If (m_LastItem > -1) Then m_Selected(m_LastItem) = 0
            If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = -1
        End If

        '-- Draw last (delete Focus) ...
        pvDrawItem m_LastItem
        m_LastItem = m_ListIndex
        '-- ... and draw actual (draw Focus)
        pvDrawItem m_ListIndex
        pvDrawFocus m_ListIndex

        '-- Ensure visible actual Selected item
        If (m_EnsureVisible) Then
            If (m_ListIndex < m_olSB And m_ListIndex > -1) Then
                m_olSB = m_ListIndex
              ElseIf (m_ListIndex > m_olSB + m_VisibleRows - 1) Then
                m_olSB = m_ListIndex - m_VisibleRows + 1
            End If
          Else
            m_EnsureVisible = -1
        End If

        RaiseEvent ListIndexChange

    On Error GoTo 0

End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)

    Set picDC.Font = m_Font
    UserControl_Resize

End Sub

'========================================================================================
' Scroll bar
'========================================================================================

Private Sub m_olSB_Change()

    If (m_LastBar <> m_olSB) Then

        '-- Store last value and reset last mouse y pos.
        m_LastBar = m_olSB
        m_LastY = -1
        '-- Force focus lost of edition text box
        If (txtEdit.Visible) Then
            txtEdit_LostFocus
        End If
        '-- Redraw list [?]
        If (m_ListIndex = m_LastItem) Then
            pvDrawList
        End If
        '-- Raise events
        RaiseEvent Scroll
        RaiseEvent TopIndexChange
    End If

End Sub

'-- ModifyItem
Public Sub ModifyItem(ByVal Index As Long, _
                      Optional ByVal Text As Variant = vbEmpty, _
                      Optional ByVal Icon As Long = -1, _
                      Optional ByVal IconSelected As Long = -1)

  '-- No items or out of bounds

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    '-- Modify
    If (Text <> vbEmpty) Then m_List(Index).Text = CStr(Text)
    If (Icon > -1) Then m_List(Index).Icon = Icon
    If (IconSelected > -1) Then m_List(Index).IconSelected = IconSelected

    '-- Redraw item
    pvDrawItem Index
    pvDrawFocus m_ListIndex

End Sub

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)

    Set picDC.MouseIcon = New_MouseIcon

End Property

'-- MouseIcon
Public Property Get MouseIcon() As Picture

    Set MouseIcon = picDC.MouseIcon

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    picDC.MousePointer() = New_MousePointer

End Property

'-- MousePointer
Public Property Get MousePointer() As MousePointerConstants

    MousePointer = picDC.MousePointer

End Property

'-- Order
Public Sub Order()

  Dim lIdx0 As Long
  Dim lIdx1 As Long
  Dim lIdx2 As Long
  Dim lDist As Long
  Dim xItem As tItem
  Dim bDesc As Boolean

    If (m_nItems > 1) Then

        lIdx0 = 0
        bDesc = (m_OrderType = [Descendent])

        If (m_SelectMode = [Single]) Then
            If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = 0
        End If

        Do
            lDist = lDist * 3 + 1
        Loop Until lDist > m_nItems

        Do
            lDist = lDist \ 3
            For lIdx1 = lDist + lIdx0 To m_nItems + lIdx0 - 1

                xItem = m_List(lIdx1)
                lIdx2 = lIdx1

                Do While (m_List(lIdx2 - lDist).Text > xItem.Text) Xor bDesc
                    m_List(lIdx2) = m_List(lIdx2 - lDist)
                    lIdx2 = lIdx2 - lDist
                    If (lIdx2 - lDist < lIdx0) Then Exit Do
                Loop
                m_List(lIdx2) = xItem
            Next lIdx1
        Loop Until lDist = 1

        ListIndex = -1
        m_olSB = 0

        '-- Unselect all and refresh
        ReDim m_Selected(0 To m_nItems)
        picDC_Paint
    End If

End Sub

Public Property Let OrderType(ByVal New_OrderType As OrderTypeCts)

    m_OrderType = New_OrderType

End Property

'-- OrderType
Public Property Get OrderType() As OrderTypeCts

    OrderType = m_OrderType

End Property

'========================================================================================
' Scrolling & Events
'========================================================================================

'-- Click()
Private Sub picDC_Click()

    If (m_ListIndex > -1) Then RaiseEvent Click

End Sub

'-- DblClick()
Private Sub picDC_DblClick()

    If (m_ListIndex > -1) Then RaiseEvent DblClick

End Sub

'-- MouseDown(Button, Shift, x, y)
Private Sub picDC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    picDC.AutoRedraw = False

  Dim lSelectedListIndex As Long

    If (Button = vbRightButton) Then
        '-- Right button
        RaiseEvent MouseDown(Button, Shift, x, y)

      Else
        '-- Get item index
        lSelectedListIndex = m_olSB + y \ m_tmpItemHeight
        '-- Update status (Selected/Unselected)
        If (lSelectedListIndex > -1 And lSelectedListIndex < m_nItems) Then
            Select Case m_SelectMode
              Case [Single]
                m_Selected(lSelectedListIndex) = -1
              Case [Multiple]
                m_Selected(lSelectedListIndex) = Not m_Selected(lSelectedListIndex)
                m_AnchorItemState = m_Selected(lSelectedListIndex)
            End Select
            '-- Store last y mouse pos. and force new ListIndex
            m_LastY = y
            ListIndex = lSelectedListIndex
        End If
        '-- Scrolling flag/Event
        m_Scrolling = -1
        RaiseEvent MouseDown(Button, Shift, x, y)
    End If

End Sub

'-- MouseMove(Button, Shift, x, y)
Private Sub picDC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim lSelectedListIndex As Long

    '-- Scrolling pos.
    m_ScrollingY = y

    '-- Up/Down [?]
    If (y < 0) Then
        pvScrollUp
        RaiseEvent MouseMove(Button, Shift, x, y)
        Exit Sub
    End If
    If (y > ScaleHeight) Then
        pvScrollDown
        RaiseEvent MouseMove(Button, Shift, x, y)
        Exit Sub
    End If

    '-- Hover selection [?]
    If (m_HoverSelection Or Button) And (y \ m_tmpItemHeight <> m_LastY \ m_tmpItemHeight) Then
        '-- Get item index
        lSelectedListIndex = m_olSB + (y \ m_tmpItemHeight)
        '-- Force new ListIndex
        If (lSelectedListIndex >= 0 And lSelectedListIndex < m_nItems) Then
            m_Selected(lSelectedListIndex) = m_AnchorItemState
            ListIndex = lSelectedListIndex
            m_LastY = y
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

'-- MouseUp(Button, Shift, x, y)
Private Sub picDC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picDC.AutoRedraw = True
    '-- Disable scrolling flag
    m_Scrolling = 0
    '-- Reset anchor state and raise event
    m_AnchorItemState = -1
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub picDC_Paint()

  Dim rFocusRect As RECT2

    '-- Design mode [?]
    If (Not Ambient.UserMode) Then

        '-- Clear
        picDC.Cls
        '-- Text alignment
        Select Case m_Alignment
          Case 0
            picDC.CurrentX = m_ItemTextLeft + m_ItemOffset
          Case 1
            picDC.CurrentX = (ScaleWidth - picDC.TextWidth(Ambient.DisplayName)) \ 2
          Case 2
            picDC.CurrentX = (ScaleWidth - picDC.TextWidth(Ambient.DisplayName)) - m_ItemOffset
        End Select
        picDC.CurrentY = m_ItemOffset
        '-- Show control name
        SetTextColor picDC.hDC, m_ColorFont
        picDC.Print Ambient.DisplayName
        '-- Item size
        SetRect rFocusRect, 0, 0, ScaleWidth, m_tmpItemHeight
        DrawFocusRect picDC.hDC, rFocusRect

      Else
        '-- Refresh list
        If (Not m_Resizing) Then pvDrawList
    End If

End Sub

Private Sub pvCalculateRects()

  Dim lIdx As Long

    '-- Calc. DC rectangles
    For lIdx = 0 To m_VisibleRows - 1
        SetRect m_ItemRct(lIdx), 0, lIdx * m_tmpItemHeight, ScaleWidth, lIdx * m_tmpItemHeight + m_tmpItemHeight
        SetRect m_TextRct(lIdx), m_ItemOffset + m_ItemTextLeft, lIdx * m_tmpItemHeight + m_ItemOffset, ScaleWidth - m_ItemOffset, lIdx * m_tmpItemHeight + m_tmpItemHeight - m_ItemOffset
        m_IconPt(lIdx).x = m_ItemOffset
        m_IconPt(lIdx).y = m_ItemOffset
    Next lIdx

End Sub

Private Sub pvDitherEffect(ByVal hDC As Long, lpRect As RECT2, ByVal lColor As Long)

  Dim hBrush As Long

    '-- Apply 'dither' effect
    hBrush = SelectObject(hDC, CreateSolidBrush(lColor))
    PatBlt hDC, lpRect.x1, lpRect.y1, lpRect.x2 - lpRect.x1, lpRect.y2 - lpRect.y1, &HA000C9
    DeleteObject SelectObject(hDC, hBrush)

End Sub

Private Sub pvDrawBox(ByVal hDC As Long, lpRect As RECT2, ByVal Offset As Long, ByVal Radius As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long)

  Dim hPen   As Long
  Dim hBrush As Long

    '-- Create Pen/Brush
    hPen = SelectObject(hDC, CreatePen(PS_SOLID, 1, ColorBorder))
    hBrush = SelectObject(hDC, CreateSolidBrush(ColorFill))

    '-- Paint box
    InflateRect lpRect, -Offset, -Offset
    RoundRect hDC, lpRect.x1, lpRect.y1, lpRect.x2, lpRect.y2, Radius, Radius
    InflateRect lpRect, Offset, Offset

    '-- Destroy Pen/Brush
    DeleteObject SelectObject(hDC, hPen)
    DeleteObject SelectObject(hDC, hBrush)

End Sub

'-- pvDrawFocus
Private Sub pvDrawFocus(lIndex As Long)

    If (Not m_Focus Or Not m_HasFocus) Then Exit Sub

    '-- Item out of area ?
    If (lIndex < m_olSB Or lIndex > m_olSB + m_VisibleRows - 1) Then Exit Sub

    '-- Draw it
    SetTextColor picDC.hDC, m_ColorFont
    DrawFocusRect picDC.hDC, m_ItemRct(lIndex - m_olSB)

End Sub

'-- pvDrawItem
Private Sub pvDrawItem(ByVal lIndex As Long)

  Dim lRctIdx As Long
  Dim vt As Variant
  Dim ii As Integer
  Dim rc As RECT2

    '-- Item out of area [?]
    If (lIndex < m_olSB Or lIndex > m_olSB + m_VisibleRows - 1) Then Exit Sub
    If (lIndex > UBound(m_List) - 1) Then Exit Sub

    '-- Reset underlined font style
    picDC.FontUnderline = 0

    '-- Rect. index
    lRctIdx = lIndex - m_olSB

    '-- Draw m_Selected Item
    If (m_Selected(lIndex)) Then

        '-- Draw back area
        Select Case m_SelectModeStyle

          Case [Standard]
            pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBackSel
            SetTextColor picDC.hDC, m_ColorFontSel

          Case [Dither] ' Effect will be applied after drawing icon
            pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBack
            SetTextColor picDC.hDC, m_ColorFontSel

          Case [Gradient_V]
            pvPaintGradient picDC.hDC, m_ItemRct(lRctIdx), m_ColorGradient1, m_ColorGradient2, GRADIENT_FILL_RECT_V
            SetTextColor picDC.hDC, m_ColorFontSel

          Case [Gradient_H]
            pvPaintGradient picDC.hDC, m_ItemRct(lRctIdx), m_ColorGradient1, m_ColorGradient2, GRADIENT_FILL_RECT_H
            SetTextColor picDC.hDC, m_ColorFontSel

          Case [Box]
            pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBack
            pvDrawBox picDC.hDC, m_ItemRct(lRctIdx), m_BoxOffset, m_BoxRadius, m_ColorBackSel, m_ColorBox
            SetTextColor picDC.hDC, m_ColorFontSel

          Case [Underline]
            pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBack
            SetTextColor picDC.hDC, m_ColorFontSel
            picDC.FontUnderline = -1

          Case [byPicture]
            If (Not SelectionPicture Is Nothing) Then
                picDC.PaintPicture SelectionPicture, 0, m_ItemRct(lRctIdx).y1, m_ItemRct(lRctIdx).x2, m_tmpItemHeight
              Else
                pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBackSel
            End If
            SetTextColor picDC.hDC, m_ColorFontSel
        End Select

        '-- Draw icon
        If (Not m_lpImgList Is Nothing) Then
            On Error Resume Next 'Image list icon # out of bounds
                If (m_WordWrap) Then
                    m_lpImgList.ListImages(m_List(lIndex).IconSelected).Draw picDC.hDC, ScaleX(m_ItemOffset, vbPixels, m_ILScale), ScaleY(m_ItemRct(lRctIdx).y1 + m_ItemOffset, vbPixels, m_ILScale), 1
                  Else
                    m_lpImgList.ListImages(m_List(lIndex).IconSelected).Draw picDC.hDC, ScaleX(m_ItemOffset, vbPixels, m_ILScale), ScaleY(m_ItemRct(lRctIdx).y1 + (m_tmpItemHeight - m_lpImgList.ImageHeight) \ 2, vbPixels, m_ILScale), 1
                End If
            On Error GoTo 0
        End If

        '-- Apply dither effect (*)
        If (m_SelectModeStyle = [Dither]) Then
            pvDitherEffect picDC.hDC, m_ItemRct(lRctIdx), m_ColorBackSel
        End If

      Else
        '-- Draw back area
        pvFillBackground picDC.hDC, m_ItemRct(lRctIdx), m_ColorBack
        SetTextColor picDC.hDC, m_ColorFont

        '-- Draw icon
        If (Not m_lpImgList Is Nothing) Then
            On Error Resume Next 'Image list icon # out of bounds
                If (m_WordWrap) Then
                    m_lpImgList.ListImages(m_List(lIndex).Icon).Draw picDC.hDC, ScaleX(m_ItemOffset, vbPixels, m_ILScale), ScaleY(m_ItemRct(lRctIdx).y1 + m_ItemOffset, vbPixels, m_ILScale), 1
                  Else
                    m_lpImgList.ListImages(m_List(lIndex).Icon).Draw picDC.hDC, ScaleX(m_ItemOffset, vbPixels, m_ILScale), ScaleY(m_ItemRct(lRctIdx).y1 + (m_tmpItemHeight - m_lpImgList.ImageHeight) \ 2, vbPixels, m_ILScale), 1
                End If
            On Error GoTo 0
        End If
    End If

    '-- Draw text...
    On Error Resume Next

        If Me.DrawStyle = IsHookMenuStyle Then
            vt = Split(m_List(lIndex).Text, vbTab)
            For ii = 1 To UBound(vt)
                LSet rc = m_TextRct(lRctIdx)
                rc.x1 = aTabs(ii - 1)
                If ii = 2 Then
                    picDC.ForeColor = &H404040   ' &HC0C0C0
                  Else
                    picDC.ForeColor = Me.FontNormal
                End If

                If lIndex = Me.ListIndex Then picDC.ForeColor = Me.FontSelected

                If vt(ii) = "-" Then
                    LSet rc = m_TextRct(lRctIdx)
                    rc.x1 = 5
                    '-- Seperator Line
                    picDC.Line (rc.x1, rc.y1 + 6)-((rc.x2 / 2.8), rc.y1 + 7), &H808080, BF
                    picDC.Line (rc.x1, rc.y1 + 6.5)-((rc.x2 / 2.8), rc.y1 + 7.5), vbButtonFace, BF
                    picDC.Line ((rc.x2 / 1.3) - (Len(vt(ii)) / 2) - 13, rc.y1 + 6)-(rc.x2 - 5, rc.y1 + 7), &H808080, BF
                    picDC.Line ((rc.x2 / 1.3) - (Len(vt(ii)) / 2) - 13, rc.y1 + 6.5)-(rc.x2 - 5, rc.y1 + 7.5), vbButtonFace, BF

                    LSet rc = m_TextRct(lRctIdx)
                    If lIndex <> Me.ListIndex Then picDC.ForeColor = vbApplicationWorkspace
                    DrawText picDC.hDC, "( Seperator )", Len("( Seperator )"), rc, DT_CENTER
                  Else
                    If ii = 2 And lIndex <> Me.ListIndex Then picDC.ForeColor = vbApplicationWorkspace
                    If vt(1) <> "-" Then DrawText picDC.hDC, vt(ii), Len(vt(ii)), rc, m_Alignment Or DT_WORDBREAK
                End If
            Next ii

            If Not m_List(lIndex).Pic Is Nothing Then
                picDC.PaintPicture m_List(lIndex).Pic, 5, rc.y1, 16, 16
            End If

          Else
            '-- Draw text...
            If (m_WordWrap) Then
                DrawText picDC.hDC, m_List(lIndex).Text, Len(m_List(lIndex).Text), m_TextRct(lRctIdx), m_Alignment Or DT_WORDBREAK
              Else
                DrawText picDC.hDC, m_List(lIndex).Text, Len(m_List(lIndex).Text), m_TextRct(lRctIdx), DT_SINGLELINE Or DT_VCENTER
            End If

        End If

End Sub

'========================================================================================
' Private
'========================================================================================

'-- pvDrawList
Private Sub pvDrawList()

  Dim lIdx As Long

    If (Extender.Visible And UBound(m_List)) Then
        '-- Draw visible rows
        For lIdx = m_olSB To m_olSB + m_VisibleRows - 1
            pvDrawItem lIdx
        Next lIdx
        '-- Draw focus
        pvDrawFocus m_ListIndex
    End If

End Sub

Private Sub pvFillBackground(ByVal hDC As Long, lpRect As RECT2, ByVal lColor As Long)

  Dim hBrush As Long

    '-- Fill background rect.
    hBrush = CreateSolidBrush(lColor)
    FillRect hDC, lpRect, hBrush
    DeleteObject hBrush

End Sub

Private Function pvGetLongColor(ByVal lColor As OLE_COLOR) As Long

  '-- Translate to long color

    If (lColor And &H80000000) Then
        pvGetLongColor = GetSysColor(lColor And &H7FFFFFFF)
      Else
        pvGetLongColor = lColor
    End If

End Function

Private Function pvGetRGBColors(ByVal lColor As Long) As RGB

  Dim sHexColor As String

    '-- Get RGB ints.
    sHexColor = String$(6 - Len(Hex$(lColor)), "0") & Hex$(lColor)
    pvGetRGBColors.R = "&H" & Mid$(sHexColor, 5, 2) & "00"
    pvGetRGBColors.G = "&H" & Mid$(sHexColor, 3, 2) & "00"
    pvGetRGBColors.B = "&H" & Mid$(sHexColor, 1, 2) & "00"

End Function

Private Sub pvPaintGradient(ByVal hDC As Long, lpRect As RECT2, lColor1 As RGB, lColor2 As RGB, ByVal Direction As Long)

  Dim tVertex(1) As TRIVERTEX
  Dim tGradient  As GRADIENT_RECT

    '-- From point/color
    With tVertex(0)
        .x = lpRect.x1
        .y = lpRect.y1
        .R = lColor1.R
        .G = lColor1.G
        .B = lColor1.B
        .Alpha = 0
    End With
    '-- To point/color
    With tVertex(1)
        .x = lpRect.x2
        .y = lpRect.y2
        .R = lColor2.R
        .G = lColor2.G
        .B = lColor2.B
        .Alpha = 0
    End With
    '-- Gradient flags
    With tGradient
        .UpperLeft = 0
        .LowerRight = 1
    End With

    '-- Paint gradient
    GradientFillRect hDC, tVertex(0), 2, tGradient, 1, Direction

End Sub

Private Sub pvReadjustRects(ByVal Offset As Long)

  Dim lIdx As Long

    '-- Adjust right side
    For lIdx = 0 To m_VisibleRows - 1
        m_ItemRct(lIdx).x2 = ScaleWidth - Offset
        m_TextRct(lIdx).x2 = ScaleWidth - m_ItemOffset - Offset
    Next lIdx

End Sub

'//

Private Sub pvReadjustScrollBar()

    If (m_nItems > m_VisibleRows) Then

        If (Not sbScroll.Visible) Then
            '-- Show scroll bar
            sbScroll.Visible = -1
            sbScroll.Refresh
            m_olSB.LargeChange = m_VisibleRows
            '-- Update item rects. right margin
            pvReadjustRects sbScroll.Width
            '-- Repaint control area
            picDC_Paint
        End If

      Else
        '-- Hide scroll bar
        sbScroll.Visible = 0
        '-- Update item rects. right margin
        pvReadjustRects 0
    End If

    '-- Update sbScroll max value and force update
    m_olSB.Max = m_nItems - m_VisibleRows
    m_olSB.Value = m_olSB.Value

End Sub

'-- pvScrollDown
Private Sub pvScrollDown()

  Dim lTimer As Long ' Timer counter
  Dim lDelay As Long ' Scrolling delay

    lDelay = 500 - 20 * (m_ScrollingY - ScaleHeight - 1)
    If (lDelay < 40) Then lDelay = 40

    '-- Scroll while MouseDown and mouse pos. > "Control bottom"
    Do While m_Scrolling And m_ScrollingY > ScaleHeight - 1
        If (GetTickCount - lTimer > lDelay) Then
            lTimer = GetTickCount
            If (m_ListIndex < m_nItems - 1) Then
                If (m_SelectMode = [Multiple]) Then
                    m_Selected(m_ListIndex + 1) = m_AnchorItemState
                End If
                ListIndex = m_ListIndex + 1
            End If
        End If
        DoEvents
    Loop

End Sub

'//

'-- pvScrollUp
Private Sub pvScrollUp()

  Dim lTimer As Long ' Timer counter
  Dim lDelay As Long ' Scrolling delay

    lDelay = 500 + 20 * m_ScrollingY
    If (lDelay < 40) Then lDelay = 40

    '-- Scroll while MouseDown and mouse pos. < "Control top"
    Do While m_Scrolling And m_ScrollingY < 0
        If (GetTickCount - lTimer > lDelay) Then
            lTimer = GetTickCount
            If (m_ListIndex > 0) Then
                If (m_SelectMode = [Multiple]) Then
                    m_Selected(m_ListIndex - 1) = m_AnchorItemState
                End If
                ListIndex = m_ListIndex - 1
            End If
        End If
        DoEvents
    Loop

End Sub

'//

Private Sub pvSetColors()

  '-- Translate colors

    m_ColorBack = pvGetLongColor(m_BackNormal)
    m_ColorBackSel = pvGetLongColor(m_BackSelected)
    m_ColorGradient1 = pvGetRGBColors(pvGetLongColor(m_BackSelectedG1))
    m_ColorGradient2 = pvGetRGBColors(pvGetLongColor(m_BackSelectedG2))
    m_ColorBox = pvGetLongColor(m_BoxBorder)
    m_ColorFont = pvGetLongColor(m_FontNormal)
    m_ColorFontSel = pvGetLongColor(m_FontSelected)

End Sub

'-- RemoveItem
Public Sub RemoveItem(ByVal Index As Long)

  Dim lIdx As Long

    '-- No items or out of bounds
    If (m_nItems = 0 Or Index > m_nItems - 1) Then Err.Raise 381

    '-- Move next items
    If (Index < m_nItems) Then
        For lIdx = Index To m_nItems - 1
            m_List(lIdx) = m_List(lIdx + 1)
            m_Selected(lIdx) = m_Selected(lIdx + 1)
        Next lIdx
    End If
    '-- Decrease items count
    m_nItems = m_nItems - 1

    '-- Redim. arrays
    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    '-- Readjust scroll bar and reset EnsureVisible flag
    pvReadjustScrollBar
    m_EnsureVisible = 0

    '-- Force new ListIndex [?]
    If (Index < m_ListIndex) Then
        If (m_ListIndex > -1) Then ListIndex = m_ListIndex - 1
      ElseIf (Index = m_ListIndex) Then
        ListIndex = -1
    End If

    '-- Refresh/Clean [?]
    If (m_nItems < m_VisibleRows) Then
        picDC.Cls
    End If
    picDC_Paint

End Sub

'-- <SelectedCount>
Public Property Get SelectedCount() As Long

  Dim lIdx As Long

    SelectedCount = 0
    For lIdx = 0 To m_nItems
        If (m_Selected(lIdx)) Then SelectedCount = SelectedCount + 1
    Next lIdx

End Property

Public Property Set SelectionPicture(ByVal New_SelectionPicture As Picture)

    Set m_SelectionPicture = New_SelectionPicture
    picDC_Paint

End Property

'-- SelectionPicture
Public Property Get SelectionPicture() As Picture

    Set SelectionPicture = m_SelectionPicture

End Property

'-- SelectMode
Public Property Get SelectMode() As SelectModeCts

    SelectMode = m_SelectMode

End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectModeCts)

  Dim lIdx As Long

    m_SelectMode = New_SelectMode

    If (Ambient.UserMode) Then
        If (New_SelectMode = [Single]) Then
            '-- Unselect all and select actual
            If (m_ListIndex > -1) Then
                For lIdx = LBound(m_List) To m_nItems
                    If (lIdx <> m_ListIndex) Then m_Selected(lIdx) = 0
                Next lIdx
                m_Selected(m_ListIndex) = -1
                pvDrawItem m_ListIndex
                pvDrawFocus m_ListIndex
            End If
        End If
    End If

    pvReadjustScrollBar
    picDC_Paint

End Property

'-- SelectModeStyle
Public Property Get SelectModeStyle() As SelectModeStyleCts

    SelectModeStyle = m_SelectModeStyle

End Property

Public Property Let SelectModeStyle(ByVal New_SelectModeStyle As SelectModeStyleCts)

    m_SelectModeStyle = New_SelectModeStyle
    picDC_Paint

End Property

'========================================================================================
' Methods
'========================================================================================

'-- SetImageList
Public Sub SetImageList(ImageListControl)

  '-- Point to ImageList control

    Set m_lpImgList = ImageListControl

    '-- Get parent scale mode
    On Error Resume Next
        m_ILScale = m_lpImgList.Parent.ScaleMode
    On Error GoTo 0

    '-- Refresh control
    picDC_Paint

End Sub

Public Function SetTabs(varTabs As Variant)

    On Error Resume Next

      Dim vSplit As Variant
      Dim i As Integer

        varTabs = varTabs & ","
        vSplit = Split(varTabs, ",")

        Erase aTabs()
        ReDim aTabs(0)
        aTabs(0) = 0

        ReDim aTabs(UBound(vSplit))

        If UBound(vSplit) > 1 Then
            For i = 1 To UBound(vSplit)
                aTabs(i - 1) = CLng(vSplit(i - 1))
            Next i
        End If

    On Error GoTo 0

End Function

'========================================================================================
' UserControl
'========================================================================================

'-- StartEdit
Public Sub StartEdit()

  '-- Item is selected...

    If (m_ListIndex > -1) Then

        '-- Let TextBox keyboard control
        KeyPreview = 0

        With txtEdit
            '-- Get TextBox item font properties
            Set .Font = m_Font
            If (m_Selected(m_ListIndex) And m_SelectModeStyle <> [Underline]) Then
                .BackColor = m_ColorBackSel
                .ForeColor = m_ColorFontSel
              Else
                .BackColor = m_ColorBack
                .ForeColor = m_ColorFont
            End If

            '-- Set alignment. Locate and resize TextBox
            If (m_WordWrap) Then
                .Alignment = Choose(m_Alignment + 1, 0, 2, 1)
                .Move m_ItemTextLeft + m_ItemOffset, (m_ListIndex - m_olSB) * m_tmpItemHeight + m_ItemOffset, m_ItemRct(m_ListIndex - m_olSB).x2 - m_ItemTextLeft - 2 * m_ItemOffset, m_tmpItemHeight - 2 * m_ItemOffset
              Else
                .Alignment = 0
                .Move m_ItemTextLeft + m_ItemOffset, (m_ListIndex - m_olSB) * m_tmpItemHeight + (m_tmpItemHeight - picDC.TextHeight("")) \ 2, m_ItemRct(m_ListIndex - m_olSB).x2 - m_ItemTextLeft - 2 * m_ItemOffset, 1
            End If

            '-- Get item text and turn TextBox to visible
            .Text = m_List(m_ListIndex).Text
            .SelStart = 0
            .SelLength = Len(txtEdit)
            .Visible = -1
            .SetFocus
        End With
    End If

End Sub

'-- TopIndex
Public Property Get TopIndex() As Long

    TopIndex = m_olSB

End Property

Public Property Let TopIndex(ByVal New_TopIndex As Long)

    On Error Resume Next

        If (New_TopIndex < 0 Or New_TopIndex > m_nItems - m_VisibleRows) Then Err.Raise 380

        m_TopIndex = New_TopIndex
        m_olSB = New_TopIndex

        RaiseEvent TopIndexChange

End Property

'========================================================================================
' Text Box (item edition)
'========================================================================================

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

  '-- Enabled new line in WordWrap mode

    If (m_WordWrap) Then
        If (KeyAscii = vbKeyReturn) Then
            m_List(m_ListIndex).Text = txtEdit
            txtEdit_LostFocus
        End If
        '-- Don't allow new line in disabled WordWrap mode
      Else
        If (KeyAscii = vbKeyReturn) Then
            m_List(m_ListIndex).Text = txtEdit
            txtEdit_LostFocus
        End If
    End If

    '-- Cancel edition
    If (KeyAscii = vbKeyEscape) Then
        txtEdit_LostFocus
    End If

End Sub

Private Sub txtEdit_LostFocus()

  '-- Hide edit TextBox and let ListBox keyboard control

    txtEdit.Visible = 0
    KeyPreview = -1

End Sub

Private Sub UserControl_EnterFocus()

  '-- Control has got focus

    m_HasFocus = -1
    pvDrawFocus m_ListIndex

End Sub

Private Sub UserControl_ExitFocus()

  '-- Control has lost focus

    m_HasFocus = 0
    pvDrawItem m_ListIndex

End Sub

Private Sub UserControl_Initialize()

  '-- Initialize arrays

    ReDim m_List(0)
    ReDim m_Selected(0)
    Set m_oSubclass = New cSubclassingThunk

    m_oSubclass.AddBeforeMsgs WM_MOUSEWHEEL
    m_oSubclass.Subclass hwnd, Me

    '-- Initialize position flags
    m_EnsureVisible = -1 ' Ensure visible last selected
    m_LastItem = -1      ' Last selected
    m_LastY = -1         ' Last Y coordinate

    '-- Initialize font object
    Set m_Font = New StdFont

    '-- Set system default scroll bar width
    sbScroll.Width = GetSystemMetrics(SM_CXVSCROLL)

    '-- Initialize 'Long' scroll bar wrapper class
    Set m_olSB = New CLongScroll
    Set m_olSB.Client = sbScroll

End Sub

'//

Private Sub UserControl_InitProperties()

    UserControl.Appearance = m_def_Appearance
    UserControl.BorderStyle = m_def_BorderStyle

    Set picDC.Font = Ambient.Font
    Set m_Font = Ambient.Font

    m_FontNormal = m_def_FontNormal
    m_FontSelected = m_def_FontSelected
    m_BackNormal = m_def_BackNormal
    m_BackSelected = m_def_BackSelected
    m_BackSelectedG1 = m_def_BackSelectedG1
    m_BackSelectedG2 = m_def_BackSelectedG2

    m_BoxBorder = m_def_BoxBorder
    m_BoxOffset = m_def_BoxOffset
    m_BoxRadius = m_def_BoxRadius

    m_Alignment = m_def_Alignment
    m_Focus = m_def_Focus
    m_HoverSelection = m_def_HoverSelection
    m_WordWrap = m_def_WordWrap

    m_ItemHeight = picDC.TextHeight("")
    m_ItemHeightAuto = m_def_ItemHeightAuto
    m_ItemOffset = m_def_ItemOffset
    m_ItemTextLeft = m_def_ItemTextLeft

    m_OrderType = m_def_OrderType
    Set m_SelectionPicture = Nothing
    m_SelectMode = m_def_SelectMode
    m_SelectModeStyle = m_def_SelectModeStyle

    m_ListIndex = -1
    m_TopIndex = -1
    pvSetColors
    m_DrawStyle = m_def_DrawStyle

End Sub

'-- KeyDown(KeyCode, Shift)
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    picDC.AutoRedraw = False

    If (m_nItems = 0 Or m_ListIndex = -1) Then
        '-- Empty list or no item selected...
        RaiseEvent KeyDown(KeyCode, Shift)
        Exit Sub
    End If

    Select Case KeyCode

      Case vbKeyUp       '{Up arrow}
        If (m_ListIndex > 0) Then
            ListIndex = m_ListIndex - 1
            RaiseEvent Click
        End If

      Case vbKeyDown     '{Down arrow}
        If (m_ListIndex < m_nItems - 1) Then
            ListIndex = m_ListIndex + 1
            RaiseEvent Click
        End If

      Case vbKeyPageDown '{PageDown}
        If (m_ListIndex < m_nItems - m_VisibleRows - 1) Then
            ListIndex = m_ListIndex + m_VisibleRows
          Else
            ListIndex = m_nItems - 1
        End If

      Case vbKeyPageUp   '{PageUp}
        If (m_ListIndex > m_VisibleRows) Then
            ListIndex = m_ListIndex - m_VisibleRows
          Else
            ListIndex = 0
        End If
        RaiseEvent Click

      Case vbKeyHome     '{Start}
        ListIndex = 0
        RaiseEvent Click

      Case vbKeyEnd      '{End}
        ListIndex = m_nItems - 1
        RaiseEvent Click

      Case vbKeySpace    '{Space} Select/Unselect
        If (m_SelectMode <> 0 And m_ListIndex > -1) Then
            m_Selected(m_ListIndex) = Not m_Selected(m_ListIndex)
            pvDrawItem m_ListIndex
            pvDrawFocus m_ListIndex
        End If
        RaiseEvent Click
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

'-- KeyPress(KeyAscii)
Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

'-- KeyPress(KeyCode, Shift)
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    picDC.AutoRedraw = True
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Dim sTmp As String

    With PropBag

        UserControl.Appearance = .ReadProperty("Appearance", m_def_Appearance)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        UserControl.Enabled = .ReadProperty("Enabled", -1)

        Set picDC.Font = .ReadProperty("Font", Ambient.Font)
        Set m_Font = .ReadProperty("Font", Ambient.Font)

        m_FontNormal = .ReadProperty("FontNormal", m_def_FontNormal)
        m_FontSelected = .ReadProperty("FontSelected", m_def_FontSelected)
        m_BackNormal = .ReadProperty("BackNormal", m_def_BackNormal)
        picDC.BackColor = .ReadProperty("BackNormal", m_def_BackNormal)
        m_BackSelected = .ReadProperty("BackSelected", m_def_BackSelected)
        m_BackSelectedG1 = .ReadProperty("BackSelectedG1", m_def_BackSelectedG1)
        m_BackSelectedG2 = .ReadProperty("BackSelectedG2", m_def_BackSelectedG2)

        m_BoxBorder = .ReadProperty("BoxBorder", m_def_BoxBorder)
        m_BoxOffset = .ReadProperty("BoxOffset", m_def_BoxOffset)
        m_BoxRadius = .ReadProperty("BoxRadius", m_def_BoxRadius)

        m_Alignment = .ReadProperty("Alignment", m_def_Alignment)
        m_Focus = .ReadProperty("Focus", m_def_Focus)
        m_HoverSelection = .ReadProperty("HoverSelection", m_def_HoverSelection)
        m_WordWrap = .ReadProperty("WordWrap", m_def_WordWrap)

        m_ItemOffset = .ReadProperty("ItemOffset", m_def_ItemOffset)
        m_ItemHeightAuto = .ReadProperty("ItemHeightAuto", m_def_ItemHeightAuto)
        m_ItemTextLeft = .ReadProperty("ItemTextLeft", m_def_ItemTextLeft)

        m_OrderType = .ReadProperty("OrderType", m_def_OrderType)
        Set m_SelectionPicture = .ReadProperty("SelectionPicture", Nothing)
        m_SelectMode = .ReadProperty("SelectMode", m_def_SelectMode)
        m_SelectModeStyle = .ReadProperty("SelectModeStyle", m_def_SelectModeStyle)

        picDC.MousePointer = .ReadProperty("MousePointer", 0)
        Set picDC.MouseIcon = .ReadProperty("MouseIcon", Nothing)
    End With

    '-- Item height initialization
    sTmp = PropBag.ReadProperty("ItemHeight", 0)
    If (sTmp < picDC.TextHeight("")) Then
        m_ItemHeight = picDC.TextHeight("")
      Else
        m_ItemHeight = sTmp
    End If

    '-- ListIndex/TopIndex initialization
    m_ListIndex = -1
    m_TopIndex = -1

    '-- Get colors
    pvSetColors
    m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)

End Sub

Private Sub UserControl_Resize()

  '-- Set item height

    If (m_ItemHeightAuto) Then
        m_tmpItemHeight = picDC.TextHeight("")
      Else
        If (m_ItemHeight < picDC.TextHeight("")) Then
            m_tmpItemHeight = picDC.TextHeight("")
          Else
            m_tmpItemHeight = m_ItemHeight
        End If
    End If

    '-- Get visible rows and readjust control height
    m_VisibleRows = ScaleHeight \ m_tmpItemHeight
    Height = (m_VisibleRows) * m_tmpItemHeight * Screen.TwipsPerPixelX + (Height - ScaleHeight * Screen.TwipsPerPixelY)

    '-- Locate and resize DC area, calc. rects and readjust scroll bar
    m_Resizing = -1
    '-- Resize controls
    picDC.Move 0, 0, ScaleWidth - IIf(sbScroll.Visible, sbScroll.Width, 0), ScaleHeight
    With sbScroll
        .Move ScaleWidth - .Width, 0, .Width, ScaleHeight
        .Visible = 0
    End With
    '-- Redim. arrays
    ReDim m_ItemRct(m_VisibleRows - 1)
    ReDim m_TextRct(m_VisibleRows - 1)
    ReDim m_IconPt(m_VisibleRows - 1)
    '-- Recalc. rects.
    pvCalculateRects
    '-- Readjust scroll bar
    pvReadjustScrollBar
    m_Resizing = 0

End Sub

Private Sub UserControl_Terminate()

    Set m_oSubclass = Nothing
    '-- Erase arrays
    Erase m_List
    Erase m_Selected
    '-- Destroy image list reference
    Set m_lpImgList = Nothing
    '-- Destroy <LongScroll> reference
    Set m_olSB = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag

        Call .WriteProperty("Appearance", UserControl.Appearance, 1)
        Call .WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
        Call .WriteProperty("Enabled", UserControl.Enabled, -1)

        Call .WriteProperty("Font", picDC.Font, Ambient.Font)
        Call .WriteProperty("FontNormal", m_FontNormal, m_def_FontNormal)
        Call .WriteProperty("FontSelected", m_FontSelected, m_def_FontSelected)
        Call .WriteProperty("BackNormal", m_BackNormal, m_def_BackNormal)
        Call .WriteProperty("BackSelected", m_BackSelected, m_def_BackSelected)
        Call .WriteProperty("BackSelectedG1", m_BackSelectedG1, m_def_BackSelectedG1)
        Call .WriteProperty("BackSelectedG2", m_BackSelectedG2, m_def_BackSelectedG2)

        Call .WriteProperty("BoxBorder", m_BoxBorder, m_def_BoxBorder)
        Call .WriteProperty("BoxOffset", m_BoxOffset, m_def_BoxOffset)
        Call .WriteProperty("BoxRadius", m_BoxRadius, m_def_BoxRadius)

        Call .WriteProperty("Alignment", m_Alignment, m_def_Alignment)
        Call .WriteProperty("Focus", m_Focus, m_def_Focus)
        Call .WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
        Call .WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)

        Call .WriteProperty("ItemHeight", m_ItemHeight, 0)
        Call .WriteProperty("ItemHeightAuto", m_ItemHeightAuto, m_def_ItemHeightAuto)
        Call .WriteProperty("ItemOffset", m_ItemOffset, m_def_ItemOffset)
        Call .WriteProperty("ItemTextLeft", m_ItemTextLeft, m_def_ItemTextLeft)

        Call .WriteProperty("OrderType", m_OrderType, m_def_OrderType)
        Call .WriteProperty("SelectionPicture", m_SelectionPicture, Nothing)
        Call .WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
        Call .WriteProperty("SelectModeStyle", m_SelectModeStyle, m_def_SelectModeStyle)

        Call .WriteProperty("MousePointer", picDC.MousePointer, 0)
        Call .WriteProperty("MouseIcon", picDC.MouseIcon, Nothing)
    End With
    Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)

End Sub

'-- WordWrap
Public Property Get WordWrap() As Boolean

    WordWrap = m_WordWrap

End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)

    m_WordWrap = New_WordWrap
    picDC_Paint

End Property


