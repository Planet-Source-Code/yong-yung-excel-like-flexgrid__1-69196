VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10110
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Project1.ctxHookMenu ctxHookMenu1 
      Left            =   1560
      Top             =   5280
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2415
      Left            =   360
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.TextBox txtEditor 
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexEditor 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   100
      Cols            =   100
      BackColorFixed  =   12648447
      GridColorFixed  =   8454143
      _NumberOfBands  =   1
      _Band(0).Cols   =   100
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnufreezeRW 
         Caption         =   "Freeze Row"
      End
      Begin VB.Menu mnufrezeeCOL 
         Caption         =   "Frezee Col"
      End
      Begin VB.Menu MnuFilldown 
         Caption         =   "Fill Down"
      End
      Begin VB.Menu mnuChart 
         Caption         =   "Chart"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColNumber As Integer
Dim RowNumber As Integer
Dim currcol As Integer
Dim currrow As Integer




Private Sub Form_Load()
    ' Setting The Height and Width of The TextBox to The Size of the Cells of the Flex Grid
    txtEditor.Height = MSFlexEditor.CellHeight
    txtEditor.Width = MSFlexEditor.CellWidth
     
    'Function for Setting the Left and Top of the TextBox to Tune of the Col and Row of the FlexGrid
    Set_Flex_TextBox_Pos
End Sub

Private Sub Form_Resize()
MSFlexEditor.Width = Form1.Width - 250

MSFlexEditor.Height = Form1.Height - 4500
End Sub

Private Sub mnuChart_Click()
Dim i As Integer
Dim j As Integer
Dim col As Integer
Dim row As Integer

MSChart1.TitleText = "Chart1"

With MSFlexEditor
'      If .Row <> .RowSel Or .Col <> .ColSel Then
         If .row < .RowSel Then
            iLowRow = .row
            iHiRow = .RowSel
         Else
            iLowRow = .RowSel
            iHiRow = .row
         End If
         If .col < .ColSel Then
            iLowCol = .col
            iHiCol = .ColSel
         Else
            iLowCol = .ColSel
            iHiCol = .col
         End If

row = 1

MSChart1.RowCount = (iHiRow - iLowRow) + 1
MSChart1.ColumnCount = (iHiCol - iLowCol) + 1


For i = iLowRow To iHiRow
      col = 1
    For j = iLowCol To iHiCol
          
        MSChart1.row = row
        MSChart1.Column = col
        MSChart1.Data = Val(.TextMatrix(i, j))
        MSChart1.RowLabel = "ROW" & i
        MSChart1.ColumnLabel = "Col" & j
        
        col = col + 1
    Next
    row = row + 1
Next
'MSChart1.Refresh
MSChart1.chartType = 3
MSChart1.Refresh
MSChart1.Visible = True




End With





End Sub

Private Sub MnuFilldown_Click()
test
End Sub

Private Sub mnufreezeRW_Click()
 
 MSFlexEditor.FixedRows = RowNumber + 1
 If currcol < RowNumber Then
    MSFlexEditor.row = RowNumber + 1
    
    If txtEditor.Visible Then
     Set_Flex_TextBox_Pos
    End If

Else
      MSFlexEditor.row = currrow
End If

End Sub

Private Sub mnufrezeeCOL_Click()
   MSFlexEditor.FixedCols = ColNumber + 1
  If currcol <= ColNumber Then
  
    MSFlexEditor.col = ColNumber + 1
        
    If txtEditor.Visible Then
        Set_Flex_TextBox_Pos
    End If
  Else
      MSFlexEditor.col = currcol
  End If

End Sub

Private Sub MSFlexEditor_Click()
    'Function for Setting the Left and Top of the TextBox to Tune of the Col and Row of the FlexGrid
    txtEditor.Visible = True
    Set_Flex_TextBox_Pos
    
   If txtEditor.Visible Then
    txtEditor.SetFocus
    End If
End Sub
Private Sub MSFlexEditor_KeyPress(KeyAscii As Integer)

' If KeyAscii Then
       txtEditor.Visible = True
       Set_Flex_TextBox_Pos
    
    txtEditor.SetFocus
 '   End If

End Sub

Private Sub MSFlexEditor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

 If Button = vbRightButton Then
  ColNumber = MSFlexEditor.MouseCol
                RowNumber = MSFlexEditor.MouseRow
 PopupMenu mnupopup
 
 End If

End Sub

Private Sub MSFlexEditor_Scroll()
txtEditor.Visible = False
End Sub

'Private Sub MSFlexEditor_EnterCell()
'Set_Flex_TextBox_Pos
    
 '   txtEditor.SetFocus
'End Sub

Private Sub txtEditor_Change()
    MSFlexEditor.Text = txtEditor.Text 'Setting the Text to the Active Cell
End Sub

Private Sub txtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    'Moving of the TextBox to the Tune of the Arrow Keys
  Dim fixedcol As Integer
  Dim fixedrow As Integer
  
  fixedcol = MSFlexEditor.FixedCols
  fixedrow = MSFlexEditor.FixedRows
  
  
    'left
    If KeyCode = 37 Then
        If Not MSFlexEditor.col <= fixedcol Then
            MSFlexEditor.col = MSFlexEditor.col - 1
            Set_Flex_TextBox_Pos
        End If
    End If
    'down
    
    If KeyCode = 40 Then
        If Not MSFlexEditor.row = MSFlexEditor.Rows - 1 Then
            MSFlexEditor.row = MSFlexEditor.row + 1
            Set_Flex_TextBox_Pos
        End If
    End If
    
    'up
    If KeyCode = 38 Then
        If Not MSFlexEditor.row <= fixedrow Then
            MSFlexEditor.row = MSFlexEditor.row - 1
            Set_Flex_TextBox_Pos
        End If
    End If
    'right
        If KeyCode = 39 Then
        If Not MSFlexEditor.col = MSFlexEditor.Cols - 1 Then
            MSFlexEditor.col = MSFlexEditor.col + 1
            Set_Flex_TextBox_Pos
        End If
    End If
  
  If txtEditor.Visible Then
    txtEditor.SetFocus
End If
End Sub
Private Sub Set_Flex_TextBox_Pos()
    'Setting Text Positions
    txtEditor.Left = MSFlexEditor.CellLeft + MSFlexEditor.Left - 10
    txtEditor.Top = MSFlexEditor.CellTop + MSFlexEditor.Top - 10
    txtEditor.Text = MSFlexEditor.Text
    currrow = MSFlexEditor.row
    currcol = MSFlexEditor.col
End Sub

Private Sub txtEditor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       txtEditor.Visible = True
     If Not MSFlexEditor.row = MSFlexEditor.Rows - 1 Then
            MSFlexEditor.row = MSFlexEditor.row + 1
            Set_Flex_TextBox_Pos
        End If
    End If
    txtEditor.SetFocus
End Sub


Private Sub test()

 With MSFlexEditor
'      If .Row <> .RowSel Or .Col <> .ColSel Then
         If .row < .RowSel Then
            iLowRow = .row
            iHiRow = .RowSel
         Else
            iLowRow = .RowSel
            iHiRow = .row
         End If
         If .col < .ColSel Then
            iLowCol = .col
            iHiCol = .ColSel
         Else
            iLowCol = .ColSel
            iHiCol = .col
         End If
           
           
     '      For k = iLowCol To iHiCol
    '     .TextMatrix(0, k) = .TextMatrix(0, iLowCol)
  
  '                  If Str = "" Then
   '                   Str = .TextMatrix(0, k)
    '                Else
     '                 Str = Str & vbTab & .TextMatrix(0, k)
      '              End If
           '   Next
   '          Str = Str & vbCrLf
               
               
               
               
         For i = iLowRow To iHiRow
            For j = iLowCol To iHiCol
                    
          .TextMatrix(i, j) = .TextMatrix(iLowRow, iLowCol)
                    
         '   If j = iHiCol Then
         '       Str = Str & .TextMatrix(i, j) & vbTab
          '  Else
           '   Str = Str & .TextMatrix(i, j) & vbTab
            
            
           ' End If
            
            
           Next
           
          'Str = Str & vbCrLf
         Next
         
'      End If
   End With
  ' Clipboard.Clear
  ' Clipboard.SetText Str
 '  MousePointer = mp_Default
 '  MsgBox "Copied to Clipbroad", vbInformation, "Copy"
   Exit Sub
End Sub

