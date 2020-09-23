VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAPIparser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Viewer Parser"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   5745
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select APV File"
      Filter          =   "APV Files (*.apv)|*.apv|All Files|*.*"
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   5790
   End
   Begin VB.PictureBox mockListBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   240
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   5
      Top             =   1545
      Width           =   5805
      Begin VB.VScrollBar lbScroll 
         Height          =   3165
         Left            =   5460
         Max             =   1000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   270
      Left            =   5745
      TabIndex        =   3
      Top             =   570
      Width           =   330
   End
   Begin VB.ComboBox cboSection 
      Height          =   315
      ItemData        =   "frmAPIparser.frx":0000
      Left            =   240
      List            =   "frmAPIparser.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1065
      Width           =   5760
   End
   Begin VB.TextBox Text2 
      Height          =   1170
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   5490
      Width           =   5880
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   5430
   End
   Begin VB.Label Label1 
      Caption         =   "Change apv file to Win16api, Win32api or WinCEapi.apv then hit the ! button"
      Height          =   270
      Left            =   210
      TabIndex        =   4
      Top             =   270
      Width           =   5580
   End
End
Attribute VB_Name = "frmAPIparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is just a semi-finished project. This form is an interface for the parser
' The parse is in the module & stands alone or can be converted to a class

' The Mock ListBox (picturebox) & associated code was to come up with a way
' to load the 52,000 constants without exceeding listbox limitations & taking
' a ton of time to parse & add items to the listbox.

' I have not finished this project & I don't expect to. The mock listbox isn't
' complete either as it should allow multiple selections

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT As Long = &H400
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

' properties of the mock listbox
Private mockTopIndex As Long        ' current index in relation to total items
Private mockListIndex As Long       ' current selected item if any
Private mockListCount As Long       ' number of items to fit in mock listbox
Private mockListItems() As String   ' array of those items
Private mockScrollRatio As Single   ' ratio of scrollbar max to total items
Private hBrushFill(0 To 1) As Long  ' 2 color brushes for background fill
Private bScrolling As Boolean       ' indicates if dragging scrollbar
Private LastScroll As Long          ' previous value of scrollbar


Private Sub cboSection_Click()
Call mockListBox_Resize
txtSearch = ""
Text2 = ""
End Sub

Private Sub Command1_Click()

dlgCommon.Flags = cdlOFNFileMustExist
On Error GoTo NoGo
dlgCommon.ShowOpen

On Error GoTo ErrorRoutine
DoEvents
APIfile = dlgCommon.FileName
Text1.Text = dlgCommon.FileName

If cboSection.ListIndex < 0 Then
    cboSection.ListIndex = 0
Else
    Call cboSection_Click
End If

Exit Sub

ErrorRoutine:
MsgBox Err.Description, vbExclamation + vbOKOnly, "Error: " & Err.Description
NoGo:
End Sub

Private Sub UpdateMockLB()

Dim X As Long, tRect As RECT
Dim f As Long

' get height of picbox font
DrawText mockListBox.hdc, "Wy", 2, tRect, DT_CALCRECT Or DT_SINGLELINE
' stretch rect horizontally
tRect.Right = mockListBox.ScaleWidth - lbScroll.Width
tRect.Left = 3
mockListBox.Cls

' fill listbox with items
For X = 0 To mockListCount - 1
    If X + mockTopIndex = mockListIndex And mockTopIndex > -1 Then
        ' listIndex item
        tRect.Left = 0
        FillRect mockListBox.hdc, tRect, hBrushFill(1)
        SetTextColor mockListBox.hdc, vbWhite
        tRect.Left = 3
    End If
    ' draw the text
    DrawText mockListBox.hdc, mockListItems(X), -1, tRect, DT_SINGLELINE Or DT_VCENTER
    ' move rect to next item
    OffsetRect tRect, 0, tRect.Bottom - tRect.Top
    
    If X + mockTopIndex = mockListIndex Then
        ' change colors back if needed
        SetTextColor mockListBox.hdc, vbBlack

'    Else
        'If X + mockTopIndex + 1 = APIsectionCount(cboSection.ListIndex) Then Exit For
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
APIfile = ""
If hBrushFill(0) Then
    DeleteObject hBrushFill(0)
    DeleteObject hBrushFill(1)
End If
End Sub

Private Sub lbScroll_Change()

If Len(lbScroll.Tag) Then Exit Sub

' when scrollbar adjusted, ensure listbox updated appropriately
If bScrolling Then
    mockTopIndex = lbScroll.Value * mockScrollRatio
Else
    mockTopIndex = mockTopIndex + (lbScroll.Value - LastScroll)
    LastScroll = lbScroll.Value
End If
' ensure when scrollbar should be at bottom it is, otherwise
If mockTopIndex >= APIsectionCount(cboSection.ListIndex) - mockListCount Then
    mockTopIndex = APIsectionCount(cboSection.ListIndex) - mockListCount
    lbScroll.Tag = "NoRecurse"
    lbScroll.Value = lbScroll.Max
    lbScroll.Tag = ""
End If
' get listing of the items to be displayed & repaint
mockListItems = GetApiListing(cboSection.ListIndex, mockTopIndex, mockTopIndex + mockListCount - 1)
UpdateMockLB
'mockListBox.Refresh

' when drag scrolling, need to know when stopped
' Without subclassing, this works well....
' Drag scrolling has captured mouseinput & GetCapture returns its hWnd
' but when done dragging, a Change event fires here & GetCapture returns zero
If GetCapture() = 0 Or bScrolling = False Then

    ' calculate the adjusted scrollbar position and/or adjusted manually tracked flags
    lbScroll.Tag = "NoRecurse"
    LastScroll = mockTopIndex / mockScrollRatio
    If LastScroll = 0 And mockTopIndex > 0 Then
        LastScroll = 1
    ElseIf LastScroll < 0 Then
        LastScroll = 0
    ElseIf lbScroll.Value = lbScroll.Max Or LastScroll >= lbScroll.Max Then
        If mockTopIndex + mockListCount < APIsectionCount(cboSection.ListIndex) Then
            LastScroll = lbScroll.Max - 5
        Else
            LastScroll = lbScroll.Max
        End If
    End If
    lbScroll.Value = LastScroll
    lbScroll.Tag = ""
    bScrolling = False
    
End If

End Sub

Private Sub lbScroll_Scroll()
bScrolling = True
Call lbScroll_Change
End Sub

Private Sub mockListBox_KeyDown(KeyCode As Integer, Shift As Integer)

' basically the keyboard navigation for the listbox

If mockTopIndex < 0 Then Exit Sub

Dim Index As Long
Select Case KeyCode
Case vbKeyHome
    mockListIndex = 0
    mockTopIndex = 0
    LastScroll = 0
    If lbScroll.Value = LastScroll Then
        UpdateMockLB
    Else
        lbScroll.Value = LastScroll
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case vbKeyEnd
    mockListIndex = APIsectionCount(cboSection.ListIndex) - 1
    LastScroll = lbScroll.Max
    If lbScroll.Value = LastScroll Then
        UpdateMockLB
    Else
        mockTopIndex = mockListIndex
        lbScroll.Value = LastScroll
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case vbKeyDown
    mockListIndex = mockListIndex + 1
    If mockListIndex >= APIsectionCount(cboSection.ListIndex) Then mockListIndex = APIsectionCount(cboSection.ListIndex) - 1
    If mockListIndex < mockTopIndex + mockListCount - 1 Then
        UpdateMockLB
    Else
        mockTopIndex = mockListIndex - mockListCount + 1
        Call lbScroll_Change
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case vbKeyUp
    mockListIndex = mockListIndex - 1
    If mockListIndex < 0 Then mockListIndex = 0
    If mockListIndex > mockTopIndex Then
        UpdateMockLB
    Else
        mockTopIndex = mockListIndex
        Call lbScroll_Change
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case vbKeyPageDown
    If mockListIndex < 0 Then
        mockListIndex = mockTopIndex + mockListCount - 1
        UpdateMockLB
    Else
        If mockListIndex < mockTopIndex + mockListCount - 1 Then
            mockListIndex = mockTopIndex + mockListCount - 1
            UpdateMockLB
        Else
            mockListIndex = mockListIndex + mockListCount
            If mockListIndex > APIsectionCount(cboSection.ListIndex) - 1 Then mockListIndex = APIsectionCount(cboSection.ListIndex) - 1
            mockTopIndex = mockListIndex - mockListCount + 1
            Call lbScroll_Change
        End If
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case vbKeyPageUp
    If mockListIndex < 0 Then
        mockListIndex = mockTopIndex
        UpdateMockLB
    Else
        If mockListIndex > mockTopIndex Then
            mockListIndex = mockTopIndex
            UpdateMockLB
        Else
            mockListIndex = mockTopIndex - mockListCount
            If mockListIndex < 0 Then mockListIndex = 0
            mockTopIndex = mockListIndex
            Call lbScroll_Change
        End If
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
Case Is > vbKeySpace
    Index = GetAlphaIndex(cboSection.ListIndex, KeyCode + 0)
    If Index > -1 Then
        mockTopIndex = Index
        mockListIndex = Index
        LastScroll = mockTopIndex / mockScrollRatio
        If LastScroll = lbScroll.Value Then
            Call lbScroll_Change
        Else
            lbScroll.Value = LastScroll
        End If
        Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
    End If
End Select
End Sub

Private Sub mockListBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' select the appropriate item when mouse clicks

If mockTopIndex < 0 Then Exit Sub

If Button = vbLeftButton Then
    Dim tRect As RECT
    ' get height of font
    DrawText mockListBox.hdc, "Wy", 2, tRect, DT_CALCRECT Or DT_SINGLELINE
    ' calculate which item was clicked & redraw it
    mockListIndex = Y \ tRect.Bottom
    OffsetRect tRect, 0, tRect.Bottom * mockListIndex
    mockListIndex = mockListIndex + mockTopIndex
    If mockListIndex > APIsectionCount(cboSection.ListIndex) - 1 Then mockListIndex = APIsectionCount(cboSection.ListIndex) - 1
    UpdateMockLB
    
    ' fill in the parsed item
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
    txtSearch.Tag = "NoRecurse"
    txtSearch = mockListItems(mockListIndex - mockTopIndex)
    txtSearch.Tag = ""
End If
End Sub

Private Sub mockListBox_Resize()

' not used as resize at this moment, but is used as a "Clear Listbox" routine

Dim tRect As RECT
' create pens if needed & reposition the scrollbar
If hBrushFill(0) = 0 Then
    hBrushFill(0) = CreateSolidBrush(vbWhite)
    hBrushFill(1) = CreateSolidBrush(vbBlue)
End If
lbScroll.Move mockListBox.ScaleWidth - lbScroll.Width, 0, lbScroll.Width, mockListBox.ScaleHeight

' calculate how many items will fit in the listbox
DrawText mockListBox.hdc, "Wy", 2, tRect, DT_CALCRECT Or DT_SINGLELINE
mockListCount = mockListBox.Height \ tRect.Bottom + 1
If tRect.Bottom * mockListCount > mockListBox.ScaleHeight Then mockListCount = mockListCount - 1

' set flags if we don't have any items for the selected API section
If APIsectionCount(cboSection.ListIndex) > 0 Then mockTopIndex = 0 Else mockTopIndex = -1

' now reset some values
mockListIndex = -1
lbScroll.Tag = "NoRecurse"
lbScroll.Value = 0
LastScroll = 0
lbScroll.LargeChange = mockListCount
lbScroll.Tag = ""
bScrolling = False

' determine size of scrollbar & size of listitem array
If APIsectionCount(cboSection.ListIndex) > 1000 Then
    mockScrollRatio = APIsectionCount(cboSection.ListIndex) / 1000
    lbScroll.Max = 1000
ElseIf APIsectionCount(cboSection.ListIndex) < mockListCount Then
    If mockTopIndex = 0 Then
        mockListCount = APIsectionCount(cboSection.ListIndex)
    Else
        mockListCount = 0
    End If
    lbScroll.Max = 0
    mockScrollRatio = 1
Else
    lbScroll.Max = APIsectionCount(cboSection.ListIndex) - 1
    mockScrollRatio = 1
End If

' fill the list item array
If mockListCount Then
    mockListItems = GetApiListing(cboSection.ListIndex, mockTopIndex, mockTopIndex + mockListCount - 1)
End If
UpdateMockLB
End Sub

Private Sub txtSearch_Change()

' use the text box to search for an item in the listbox

If mockTopIndex < 0 Then Exit Sub

If Len(txtSearch.Tag) Then Exit Sub
If Len(txtSearch) = 0 Then Exit Sub

Dim Index As Long, I As Integer, nextOffset As Long, nextIndex As Long

nextIndex = -1
For I = 1 To Len(txtSearch)
    nextIndex = SearchAPIFile(cboSection.ListIndex, Mid$(txtSearch.Text, 1, I), apBeginsWith, nextIndex, nextOffset)
    If nextIndex < 0 Then Exit For
    Index = nextIndex
Next

If Index > -1 Then
    ' update the listbox as appropriate
    mockListIndex = Index
    If Index < mockTopIndex Or _
        Index > mockTopIndex + mockListCount - 1 Then
        mockTopIndex = Index
        LastScroll = mockTopIndex / mockScrollRatio
        If LastScroll = lbScroll.Value Then
            Call lbScroll_Change
        Else
            lbScroll.Value = LastScroll
        End If
    Else
        UpdateMockLB
    End If
    Text2 = ParseDBsection(cboSection.ListIndex, mockListIndex, mockListItems(mockListIndex - mockTopIndex))
End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

' allow up/down keys in search box to be used in listbox too
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    Call mockListBox_KeyDown(KeyCode, Shift)
End If
End Sub
