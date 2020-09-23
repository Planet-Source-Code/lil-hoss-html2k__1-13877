VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl ctlHTMLColor 
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ScaleHeight     =   2970
   ScaleWidth      =   4365
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   1080
   End
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   60000
      TextRTF         =   $"ctlColor.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   60000
      TextRTF         =   $"ctlColor.ctx":00DF
   End
End
Attribute VB_Name = "ctlHTMLColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////
'//                This Code is Copyright VB Web 1999                   //
'//         You MAY NOT re-distribute this source code. Instead,        //
'//                     please provide a link to                        //
'//       http://www.vbweb.co.uk/dev/?devpad.htm                        //
'//                                                                     //
'//       If you would like to become a beta tester, please email       //
'//   devpadbeta@vbweb.f9.co.uk with your Name, VB Version and PC Spec  //
'//                                                                     //
'//           Please report any bugs to bugs@vbweb.f9.co.uk             //
'/////////////////////////////////////////////////////////////////////////

Option Explicit
'// Subclassing
Implements ISubclass
Private blnJustOutTag As Boolean
Private blnJustDeletingTag As Boolean
Private blnInQuote As Boolean
'Private m_emr As EMsgResponse
Private bSubclassing As Boolean
Private bDelLeftTag As Boolean
'// Win API Constants
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private sText As String
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)

Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const CFM_BACKCOLOR = &H4000000
Private Const EM_LINEFROMCHAR = &HC9

Private Const EM_EXSETSEL = (WM_USER + 55)

Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_PASTE = &H302
Private Const WM_UNDO = &H304

Private Const EM_SETCHARFORMAT = (WM_USER + 68)
Private Const EM_GETCHARFORMAT = (WM_USER + 58)
Private Const EM_POSFROMCHAR = &HD6&
Private Const SCF_SELECTION = &H1&
Private Const LF_FACESIZE = 32
Private Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type
Private Const EM_CANUNDO = &HC6
Private Const EM_CANPASTE = (WM_USER + 50)

'// Win API
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'// Private Variables
Private m_Colour_Comment As OLE_COLOR
Private m_Colour_Keyword As OLE_COLOR
Private m_Colour_Text As OLE_COLOR

Private m_lngSelPos As Long
Private m_lngSelLen As Long
Private m_blnBusy As Boolean
Public bBusy As Boolean
Private m_blnChanged As Boolean

Private m_lngNextLine As Long
Private m_lngLastLine As Long
Private m_lngLastLinePos As Long

Private nBufferLen As Long
Private i As Long
Private blnAll As Boolean
Private Status As Status

'// Default Property Values
Const m_def_Border = False
Const m_def_TextRTF = ""
Const m_def_SelRTF = ""
Const m_def_AutoVerbMenu = 0
Const m_def_BulletIndent = 0
Const m_def_Locked = 0
Const m_def_MultiLine = 0
Const m_def_OLEDragMode = 0
Const m_def_RightMargin = 60000
Const m_def_ScrollBars = 0
Const m_def_Text = ""
Const m_def_ToolTipText = ""
Const m_def_Colour_Comment = &H8000&
Const m_def_Colour_Keyword = &H7F0000
Const m_def_Colour_Text = &H0&
Const m_def_Indent = 4
Const m_def_AutoIndent = True
Const m_def_CancelColour = False
Const m_def_Saved = True
Const m_def_Modified = False
'// Property Variables
Dim m_Mode As ModeConstants
Dim m_CancelColour As Boolean
Dim m_Saved As Boolean
Dim m_Modified As Boolean
Dim m_Indent As Integer
Dim m_AutoIndent As Boolean
Dim m_FileName As String
Dim m_IgnoreVBHeader As Boolean
'Dim bInJavaComment As Boolean
'// Event Declarations
Event BeforeCut(Cancel As Boolean)
Event BeforeCopy(Cancel As Boolean)
Event BeforePaste(Cancel As Boolean)
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event SelChange()
Event HScroll()
Event VScroll()
Event ProgressStart()
Event ProgressChange(ByVal PercentComplete As Long)
Event ProgressComplete()
Event LoadFile(ByVal strFileName As String)
Private bMainLocked As Boolean
'////////////////////////////////////////////////////////////
Public Sub ColourText()
    m_blnBusy = True
    LockMain
    SaveCursorPos
    '// Return the text
    rtfMain.SelStart = 0
    rtfMain.SelLength = CharacterCount
    rtfMain.SelColor = m_Colour_Text
    RestoreCursorPos
    UnLockMain
    '// We are not busy
    m_blnBusy = False
    m_blnChanged = False
End Sub
'////////////////////////////////////////////////////////////
Public Sub AboutBox()
    '// Shows our about box
    MsgBox "Copyright 1999 VB Web - www.vbweb.f9.co.uk" & Chr(13) & "    Written By James Crowley of VB Web Internet Solutions"
End Sub
Public Sub ColourDocument(nType As ModeConstants)
'    If nType = vbwrtf Then Exit Sub
'    If nType = vbwText Then
'        m_blnBusy = True
'        LockMain
'        SetSelection 0, -1
'        rtfMain.Text = rtfMain.Text
'    End If
    
    m_blnBusy = True
    '// Get the text
    rtfTemp.Text = rtfMain.Text
    blnAll = True
    RunColorize True
    Mode = nType
    '// Highlight the text, except the last return
    rtfTemp.SelStart = 0
    rtfTemp.SelLength = TempCharacterCount
    LockMain
    SaveCursorPos
    '// Return the text
    rtfMain.TextRTF = rtfTemp.SelRTF
    RestoreCursorPos
    '// Get the new line pos
    m_lngLastLine = SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
    m_lngLastLinePos = rtfMain.SelStart + rtfMain.SelLength
    UnLockMain
    '// We are not busy
    m_blnBusy = False
    m_blnChanged = False
    rtfTemp.Text = ""
End Sub
'////////////////////////////////////////////////////////////
Public Sub Cut()
    SendMessage rtfMain.hWnd, WM_CUT, 0&, 0&
End Sub
'////////////////////////////////////////////////////////////
Public Sub Copy()
    SendMessage rtfMain.hWnd, WM_COPY, 0&, 0&
End Sub
'////////////////////////////////////////////////////////////
Public Sub Paste()
    '// Get the clipboard data
    
    If IsNotCode Then
        SendMessage rtfMain.hWnd, WM_PASTE, 0&, 0&
    Else
        If Clipboard.GetFormat(vbCFText) = False Then
            '// Not text, get out of here
            Exit Sub
        End If
        InsertCode Clipboard.GetText, True
    End If
End Sub
'////////////////////////////////////////////////////////////
Public Sub SelectAll()
    rtfMain.SelStart = 0
    rtfMain.SelLength = CharacterCount
End Sub
'////////////////////////////////////////////////////////////
Public Sub Append()
    Dim strClipboard As String
    '// Get the clipboard data
    If Clipboard.GetFormat(vbCFText) Then
        strClipboard = Clipboard.GetText(vbCFText)
        Clipboard.Clear
        Clipboard.SetText strClipboard & rtfMain.SelText, vbCFText
    ElseIf Clipboard.GetFormat(vbCFRTF) Then
        strClipboard = Clipboard.GetText(vbCFText)
        Clipboard.Clear
        Clipboard.SetText strClipboard & rtfMain.SelRTF, vbCFRTF
    Else
        ErrHandler , "Invalid Clipboard Format", "Append"
    End If
End Sub

'////////////////////////////////////////////////////////////
Public Sub InsertCode(strCode As String, bReplaceSelection As Boolean)
Dim lngTempPos As Long
    Modified = True
    m_blnChanged = True
'    If Mode = vbwText Then
'        If Not (bReplaceSelection) Then
'            rtfMain.SelLength = 0
'        End If
'        rtfMain.SelText = strCode
'        Exit Sub
'    End If
    m_blnBusy = True
    LockMain
    '// Get the text
'    If Right(strCode, 2) = vbCrLf Then
'        strCode = Left(strCode, Len(strCode) - 2)
'    End If
    rtfTemp.Text = strCode  '""
    'SendMessage rtfTemp.hwnd, WM_PASTE, 0&, 0&
    'rtfTemp.SelStart = 0
    'rtfTemp.SelLength = TempCharacterCount
    'rtfTemp.SelText = rtfTemp.SelText
   ' SendMessage rtfTemp.hwnd, WM_COPY, 0&, 0&
   ' SendMessage rtfTemp.hwnd, WM_PASTE, 0&, 0&
    '// colour the whole lot in the temp box
    blnAll = True
    RunColorize True
    '// Highlight the text, except the last return
    rtfTemp.SelStart = 0
    rtfTemp.SelLength = TempCharacterCount
    'If bReplaceSelection Then
    '    RestoreCursorPos
    'Else
    '    rtfMain.SelStart = m_lngSelPos
    'End If
    
    '// Return the text
    rtfMain.SelRTF = rtfTemp.SelRTF
    'lngTempPos = m_lngSelPos
    'SaveCursorPos
    '// then just colour the top line
    'rtfMain.SelStart = lngTempPos
    'ColourCurLine
    m_blnChanged = True
    '// We are not busy
    m_blnBusy = False
    '// Get the new line pos
    m_lngLastLine = SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
    m_lngLastLinePos = rtfMain.SelStart + rtfMain.SelLength
    '// Allow redrawing and refresh the contents
    UnLockMain
    rtfTemp.Text = ""
End Sub
'////////////////////////////////////////////////////////////
Public Sub Clear()
    rtfMain.Text = ""
    ColourText
End Sub
'////////////////////////////////////////////////////////////
Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
Attribute Find.VB_Description = "Searches the text in a RichTextBox control for a given string."
    If IsMissing(vEnd) Then vEnd = -1
    Find = rtfMain.Find(bstrString, vStart, vEnd, vOptions)
End Function
'////////////////////////////////////////////////////////////
Public Sub UpTo(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute UpTo.VB_Description = "Moves the insertion point up to, but not including, the first character that is a member of the specified character set in a RichTextBox control."
    rtfMain.UpTo bstrCharacterSet, vForward, vNegate
End Sub
'////////////////////////////////////////////////////////////
Public Sub Span(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute Span.VB_Description = "Selects text in a RichTextBox control based on a set of specified characters."
    rtfMain.Span bstrCharacterSet, vForward, vNegate
End Sub
'////////////////////////////////////////////////////////////
Public Sub SelPrint(ByVal lHDC As Long, Optional ByVal vStartDoc As Variant)
Attribute SelPrint.VB_Description = "Sends formatted text in a RichTextBox control to a device for printing."
    rtfMain.SelPrint lHDC, vStartDoc
End Sub
'////////////////////////////////////////////////////////////
Public Function SaveFile(strFileName As String, FileType As LoadSaveConstants, Optional bIgnoreSave As Boolean = False) As Boolean
    On Error GoTo DiskErr
    rtfMain.SaveFile strFileName, FileType
    If bIgnoreSave = False Then
        SaveFile = True
        Saved = True
        Modified = False
        m_FileName = strFileName
    End If
    Exit Function
DiskErr:
    Err.Raise Err, "SaveFile", Error
End Function
'////////////////////////////////////////////////////////////
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
    rtfMain.Refresh
End Sub
'////////////////////////////////////////////////////////////
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    rtfMain.OLEDrag
End Sub
'////////////////////////////////////////////////////////////
Public Function LoadFile(ByVal strFileName, ByVal FileType As LoadSaveConstants) As Boolean
    On Error GoTo DiskErr
    rtfMain.LoadFile strFileName, FileType
    If FileType = rtfRTF Then
        '// if we have just loaded an RTF file, we are not colouring!
        'Mode = vbwRTF
    End If
    '// if we are colouring, then colour!
    'If Mode = vbwVB Then ColourVB
    LoadFile = True
    Saved = True
    Modified = False
    m_FileName = strFileName
    RaiseEvent LoadFile(strFileName)
    Exit Function
DiskErr:
    ErrHandler Err, Error, "Doc.LoadFile"
End Function
Public Function LoadTemplate(strFileName As String)
    On Error Resume Next
    rtfMain.LoadFile strFileName, rtfText
End Function
'////////////////////////////////////////////////////////////
'// Properties
'////////////////////////////////////////////////////////////


Public Property Get CancelColour() As Boolean
    CancelColour = m_CancelColour
End Property
Public Property Let CancelColour(ByVal New_CancelColour As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    m_CancelColour = New_CancelColour
    PropertyChanged "CancelColour"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Saved() As Boolean
    Saved = m_Saved
End Property
Public Property Let Saved(ByVal New_Saved As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Saved = New_Saved
    PropertyChanged "Saved"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Modified() As Boolean
    Modified = m_Modified
End Property
Public Property Let Modified(ByVal New_Modified As Boolean)
    m_Modified = New_Modified
    PropertyChanged "Modified"
End Property
'////////////////////////////////////////////////////////////
Public Property Get CanPaste() As Boolean
   CanPaste = SendMessage(rtfMain.hWnd, EM_CANPASTE, 0&, 0&)
End Property
'////////////////////////////////////////////////////////////
Public Property Get CanCopy() As Boolean
   If (rtfMain.SelLength > 0) Then
      CanCopy = True
   End If
End Property
'////////////////////////////////////////////////////////////
Public Property Get AutoIndent() As Boolean
    AutoIndent = m_AutoIndent
End Property

Public Property Let AutoIndent(ByVal New_AutoIndent As Boolean)
    m_AutoIndent = New_AutoIndent
    PropertyChanged "AutoIndent"
End Property
''////////////////////////////////////////////////////////////
'Public Property Get AutoVerbMenu() As Boolean
'    AutoVerbMenu = rtfMain.AutoVerbMenu
'End Property
'Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
'    rtfMain.AutoVerbMenu = New_AutoVerbMenu
'    PropertyChanged "AutoVerbMenu"
'End Property
''////////////////////////////////////////////////////////////
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = rtfMain.BackColor
'End Property
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    rtfMain.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
''////////////////////////////////////////////////////////////
'Public Property Get BulletIndent() As Single
'    BulletIndent = rtfMain.BulletIndent
'End Property
'Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
'    rtfMain.BulletIndent = New_BulletIndent
'    PropertyChanged "BulletIndent"
'End Property
'////////////////////////////////////////////////////////////
Public Property Get Enabled() As Boolean
    Enabled = rtfMain.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    rtfMain.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
'////////////////////////////////////////////////////////////
Public Property Get FileName() As String
    FileName = m_FileName
End Property
Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property
'////////////////////////////////////////////////////////////
Public Property Get FileTitle() As String
    Dim strPath As String
    strPath = rtfMain.FileName
    FileTitle = Right$(strPath, Len(strPath) - InStrRev(strPath, "\"))
End Property
'////////////////////////////////////////////////////////////
Public Property Get Locked() As Boolean
    Locked = rtfMain.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    rtfMain.Locked = New_Locked
    PropertyChanged "Locked"
End Property
'////////////////////////////////////////////////////////////
Public Property Get MouseIcon() As Picture
    Set MouseIcon = rtfMain.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtfMain.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
''////////////////////////////////////////////////////////////
'Public Property Get MultiLine() As Boolean
'    MultiLine = rtfMain.MultiLine
'End Property
'Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'    rtfMain.MultiLine = New_MultiLine
'    PropertyChanged "MultiLine"
'End Property
''////////////////////////////////////////////////////////////
'Public Property Get OLEDragMode() As OLEDragConstants
'    OLEDragMode = rtfMain.OLEDragMode
'End Property
'Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
'    rtfMain.OLEDragMode = New_OLEDragMode
'    PropertyChanged "OLEDragMode"
'End Property
''////////////////////////////////////////////////////////////
'Public Property Get OLEDropMode() As Integer
'    OLEDropMode = rtfMain.OLEDropMode
'End Property
'Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
'    rtfMain.OLEDropMode() = New_OLEDropMode
'    PropertyChanged "OLEDropMode"
'End Property
''////////////////////////////////////////////////////////////
'Public Property Get RightMargin() As Single
'    RightMargin = rtfMain.RightMargin
'End Property
'Public Property Let RightMargin(ByVal New_RightMargin As Single)
'    rtfMain.RightMargin = New_RightMargin
'    PropertyChanged "RightMargin"
'End Property
'////////////////////////////////////////////////////////////
'Public Property Get ScrollBars() As ScrollBarsConstants
'    ScrollBars = rtfMain.ScrollBars
'End Property
'Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsConstants)
'    rtfMain.ScrollBars = New_ScrollBars
'    PropertyChanged "ScrollBars"
'End Property
'////////////////////////////////////////////////////////////
Public Property Get Text() As String
    Text = rtfMain.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    rtfMain.Text = New_Text
    PropertyChanged "Text"
End Property
''////////////////////////////////////////////////////////////
'Public Property Get ToolTipText() As String
'    ToolTipText = rtfMain.ToolTipText
'End Property
'Public Property Let ToolTipText(ByVal New_ToolTipText As String)
'    rtfMain.ToolTipText = New_ToolTipText
'    PropertyChanged "ToolTipText"
'End Property
''////////////////////////////////////////////////////////////
Public Property Get TabStop() As Boolean
    TabStop = rtfMain.TabStop
End Property
Public Property Let TabStop(ByVal New_Tab As Boolean)
    rtfMain.TabStop = New_Tab
    PropertyChanged "TabStop"
End Property
'////////////////////////////////////////////////////////////
Public Property Get IndentValue() As Integer
    IndentValue = m_Indent
End Property

Public Property Let IndentValue(ByVal New_Indent As Integer)
    m_Indent = New_Indent
    PropertyChanged "Indent"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Mode() As ModeConstants
    Mode = m_Mode
End Property
Private Function LoadKeywords(sFile As String) As String
On Error Resume Next
    Dim iFileNum As Integer
    iFileNum = FreeFile
    Open App.Path & "\keywords\" & sFile For Input As iFileNum
    LoadKeywords = Input(LOF(iFileNum), iFileNum)
    Close #iFileNum
End Function
Public Property Let Mode(ByVal New_Mode As ModeConstants)
    m_Mode = New_Mode
    Select Case m_Mode

    Case Else
        
    End Select
    PropertyChanged "Mode"
End Property
'////////////////////////////////////////////////////////////
Public Property Get SelLength() As Long
    SelLength = rtfMain.SelLength
End Property
Public Property Let SelLength(ByVal New_Len As Long)
    rtfMain.SelLength = New_Len
End Property
'////////////////////////////////////////////////////////////
Public Property Get SelStart() As Long
    SelStart = rtfMain.SelStart
End Property
Public Property Let SelStart(ByVal New_Start As Long)
    rtfMain.SelStart = New_Start
End Property
'////////////////////////////////////////////////////////////
Public Property Let SelText(ByVal New_Sel As String)
    If New_Sel = "" Then
        rtfMain.SelText = ""
    Else
        InsertCode New_Sel, True
    End If
End Property
Public Property Get SelText() As String
    SelText = rtfMain.SelText
End Property
'////////////////////////////////////////////////////////////
Public Property Get hWnd() As Long
    hWnd = rtfMain.hWnd
End Property
'////////////////////////////////////////////////////////////
Public Property Get LineCount() As Long
Attribute LineCount.VB_Description = "Returns the number of lines"
Attribute LineCount.VB_MemberFlags = "400"
    LineCount = SendMessage(rtfMain.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Property
Public Property Let LineCount(ByVal New_LineCount As Long)
    Err.Raise 382
End Property
'////////////////////////////////////////////////////////////
Public Property Get GetFirstLineVisible() As Long
Attribute GetFirstLineVisible.VB_Description = "Returns the first line visible on the control"
Attribute GetFirstLineVisible.VB_MemberFlags = "400"
    GetFirstLineVisible = SendMessage(rtfMain.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Property
Public Property Let GetFirstLineVisible(ByVal New_GetFirstLineVisible As Long)
    Err.Raise 382
End Property
'////////////////////////////////////////////////////////////
Public Property Get CurrentColumn() As Long
Attribute CurrentColumn.VB_Description = "Returns the current column"
Attribute CurrentColumn.VB_MemberFlags = "400"
    Dim lngCurLine As Long
    '// Current Line
    lngCurLine = 1 + rtfMain.GetLineFromChar(rtfMain.SelStart)
    '// Column
    CurrentColumn = SendMessage(rtfMain.hWnd, EM_LINEINDEX, ByVal lngCurLine - 1, 0&)
    CurrentColumn = (rtfMain.SelStart) - CurrentColumn
End Property
Public Property Let CurrentColumn(ByVal New_CurrentColumn As Long)
    Err.Raise 382
End Property
'////////////////////////////////////////////////////////////
Public Property Get CurrentLine() As Long
Attribute CurrentLine.VB_Description = "Returns the current line"
Attribute CurrentLine.VB_MemberFlags = "400"
    CurrentLine = 1 + rtfMain.GetLineFromChar(rtfMain.SelStart)
End Property
Public Property Let CurrentLine(ByVal New_CurrentLine As Long)
    Err.Raise 382
End Property
'////////////////////////////////////////////////////////////
Public Property Get TextRTF() As String
Attribute TextRTF.VB_MemberFlags = "400"
    TextRTF = rtfMain.TextRTF
End Property
Public Property Let TextRTF(ByVal New_TextRTF As String)
    If Ambient.UserMode = False Then Err.Raise 387
    Mode = vbwrtf
    rtfMain.TextRTF = New_TextRTF
    PropertyChanged "TextRTF"
End Property
'////////////////////////////////////////////////////////////
Public Property Get SelRTF() As String
    SelRTF = rtfMain.SelRTF
End Property
Public Property Let SelRTF(ByVal New_SelRTF As String)
Attribute SelRTF.VB_MemberFlags = "400"
    If Ambient.UserMode = False Then Err.Raise 387
    Mode = vbwrtf
    rtfMain.SelRTF = New_SelRTF
    PropertyChanged "SelRTF"
End Property
'////////////////////////////////////////////////////////////
'Public Property Get Border() As Boolean
'    Border = BorderStyle
'End Property
'Public Property Let Border(ByVal bState As Boolean)
'    BorderStyle() = Abs(bState)
'    PropertyChanged "Border"
'End Property
'////////////////////////////////////////////////////////////
Public Property Get Font() As Font
    Set Font = rtfMain.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set rtfMain.Font = New_Font
    Set rtfTemp.Font = New_Font
    PropertyChanged "Font"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Colour_Comment() As OLE_COLOR
    Colour_Comment = m_Colour_Comment
End Property
Public Property Let Colour_Comment(ByVal New_Colour_Comment As OLE_COLOR)
    m_Colour_Comment = New_Colour_Comment
    PropertyChanged "Colour_Comment"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Colour_Keyword() As OLE_COLOR
    Colour_Keyword = m_Colour_Keyword
End Property
Public Property Let Colour_Keyword(ByVal New_Colour_Keyword As OLE_COLOR)
    m_Colour_Keyword = New_Colour_Keyword
    PropertyChanged "Colour_Keyword"
End Property
'////////////////////////////////////////////////////////////
Public Property Get Colour_Text() As OLE_COLOR
    Colour_Text = m_Colour_Text
End Property
Public Property Let Colour_Text(ByVal New_Colour_Text As OLE_COLOR)
    m_Colour_Text = New_Colour_Text
    PropertyChanged "Colour_Text"
End Property
Public Property Get Version() As String
    Version = "0.50"
End Property
'////////////////////////////////////////////////////////////
'// Private Procedures
'////////////////////////////////////////////////////////////

Private Sub SaveCursorPos()
    '// Save the sel
    m_lngSelPos = rtfMain.SelStart
    m_lngSelLen = rtfMain.SelLength
End Sub
'////////////////////////////////////////////////////////////
Private Sub RestoreCursorPos()
    '// Save the sel
    rtfMain.SelStart = m_lngSelPos
    rtfMain.SelLength = m_lngSelLen
End Sub
'////////////////////////////////////////////////////////////
Private Sub LockMain()
If bMainLocked Then Exit Sub
    '// Lock the text box to prevent changes
    'SendMessage rtfTemp.hwnd, WM_SETREDRAW, False, 0&
    LockWindowUpdate rtfMain.hWnd
    'SendMessage rtfMain.hwnd, EM_HIDESELECTION, True, 0&
    'SendMessage rtfTemp.hwnd, EM_HIDESELECTION, True, 0&
    rtfMain.Locked = True
    bMainLocked = True
End Sub
'////////////////////////////////////////////////////////////
Private Sub UnLockMain()
If Not bMainLocked Then Exit Sub

    '// Lock the text box to prevent changes
    rtfMain.Locked = False
    If bBusy = True Then Exit Sub
    'SendMessage rtfMain.hwnd, EM_HIDESELECTION, False, 0&
    LockWindowUpdate 0
     bMainLocked = False
End Sub
'////////////////////////////////////////////////////////////
Private Function GetLineText(rtf As Object) As String
    Dim lCurrLine As Long
    Dim lStartPos As Long
    Dim lEndLen As Long
    Dim lSelLen As Long
    Dim strSearch As String
    Dim sText As String
    Dim lCommentStart As Long
    Dim lCommentEnd As Long
    Static bInComment As Boolean
   ' LockMain
    rtf.SelLength = 0
    '// Get current line
    lCurrLine = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, rtf.SelStart + rtf.SelLength, 0&)
    Select Case Mode
'    Case vbwVB, vbwVBScript
'        '// Set the start pos at the beginning of the line
'        rtf.SelStart = SendMessage(rtf.hWnd, EM_LINEINDEX, lCurrLine, 0&)
'        '// Get the length of the line
'        lSelLen = SendMessage(rtf.hWnd, EM_LINELENGTH, rtf.SelStart + rtf.SelLength, 0&)
'        If Err Then
'            '// Line does not exist
'            GetLineText = -1
'            Exit Function
'        End If
    
    Case Else
        sText = rtf.Text
        lCommentStart = InStrRev(Left$(sText, rtfMain.SelStart), "/*")
        If lCommentStart <> 0 Then
            lCommentEnd = InStr(lCommentStart, sText, "*/")
        Else
            lCommentEnd = -1
        End If
        If lCommentEnd > rtfMain.SelStart + 1 Then
            '// Set the start pos at the beginning of the line
            rtf.SelStart = lCommentStart - 1
            If lCommentEnd = lCommentStart + 1 Then
                lSelLen = Len(sText) - rtf.SelStart
            Else
                '// Get the length of the line
                lSelLen = lCommentEnd + 1 - rtf.SelStart
            End If
            bInComment = True
        Else
            '// Set the start pos at the beginning of the line
            rtf.SelStart = SendMessage(rtf.hWnd, EM_LINEINDEX, lCurrLine, 0&)
            If rtf.SelStart < lCommentEnd Then
                rtf.SelStart = lCommentEnd + 1
            End If
            '// Get the length of the line
            lSelLen = SendMessage(rtf.hWnd, EM_LINELENGTH, rtf.SelStart + rtf.SelLength, 0&)
            If Err Then
                '// Line does not exist
                GetLineText = -1
                Exit Function
            End If
        End If
    End Select
   ' UnLockMain
    '// Select the line
    rtf.SelLength = lSelLen
    GetLineText = rtf.SelText
End Function
'////////////////////////////////////////////////////////////
Private Sub ColourCurLine()
    Dim strText As String
    Dim intSelLen As Long
    'LockMain
    '// Copy the text into our tempory text box, ColourVB it and
    '// return the text to the original text box
    rtfTemp.Text = GetLineText(rtfMain)
    '// delete line
    rtfMain.SelText = ""
    If rtfTemp.Text = "" Then
        '// if there is no line, exit
        rtfMain.SelRTF = ""
        GoTo TheEnd
    End If
    blnAll = False
    RunColorize
    '// move text back. Select text so we do not get
    '// trailing return in RTF code
    rtfTemp.SelStart = 0
    rtfTemp.SelLength = TempCharacterCount
    rtfMain.SelRTF = rtfTemp.SelRTF
    '// Restore positions
TheEnd:
    RestoreCursorPos
    rtfMain.SelColor = m_Colour_Text
    m_blnChanged = False
    UnLockMain
End Sub
Private Sub RunColorize(Optional bIncHTML As Boolean = False)
    Select Case Mode
    Case VBWHTML
        If bIncHTML Then ColorizeWordsHTML
    End Select
End Sub
Private Sub Colour()

End Sub
Public Sub ColorizeWordsHTML()
    Dim sBuffer    As String
    Dim sTmpWord   As String
    Dim nJ         As Long
    Dim nStartPos  As Long
    Dim nSelLen    As Long
    Dim nWordPos   As Long
    Dim blnInTag   As Boolean
    Dim blnInQuote As Boolean
    Dim nEndPos    As Long
    Dim nEndTag    As Long
    Dim sCode As String
    Dim nSubPos As Long
    Dim nNextSpace As Long
    Dim blnScript As Boolean
    Dim sTemp As String
    tmrUpdate.Enabled = True
    RaiseEvent ProgressStart
    m_CancelColour = False
    With rtfTemp
        'Screen.MousePointer = vbhourglass
        '// Apply the font and ColourVB styles
        SetTempSelection 0, -1
        sBuffer = .Text
        If InStr(1, sBuffer, "<") = 0 Then
            .SelColor = m_Colour_Text
        Else
            .SelColor = &HBF0000 'm_Colour_Text
        End If
        .SelFontName = rtfMain.Font
        .SelFontSize = rtfMain.Font.Size
        '// Loop through every character in the text and check
        '// for different letters and characters
        
        nBufferLen = Len(sBuffer)
        For i = 1 To nBufferLen
            If m_CancelColour Then GoTo TheEnd
            DoEvents
            Select Case Mid$(sBuffer, i, 1)
                Case "<"
                    '// in tag
                    blnInTag = True
                    If Mid$(sBuffer, i + 1, 3) = "!--" Then
                        nEndTag = InStr(i, sBuffer, "-->") ' + 2
                        .SelStart = i - 1 ' - 2
                        .SelLength = nEndTag - (i - 3)
                        .SelColor = &H7F7F7F
                        i = nEndTag
                        '// we are out of tag
                        blnInTag = False
                    ElseIf LCase$(Mid$(sBuffer, i + 1, 6)) = "script" Then
                        
                        
                        'nEndTag = InStr(i + 1, sBuffer, ">")
                        'i = nEndTag - 1
                        blnScript = True
                        GoTo CaseElse
                    Else
CaseElse:
                        nEndTag = InStr(i + 1, sBuffer, ">")
                        nEndPos = InStr(i + 1, sBuffer, "=")
                        
                        'If nEndTag = 0 Then
                        '    '// no end tag
                        '    i = nBufferLen
                        'Else
                            If nEndPos <> 0 And nEndPos < nEndTag Then
                                '// = found.
                                DoEvents
                            Else
                                If nEndTag <> 0 Then
                                    '// no =, skip to just before end of tag
                                    i = nEndTag - 1
                                Else
                                    
                                End If
                            End If
                        'End If
                    End If
                Case ">"
                    '// out of tag
                    
                    If blnScript Then
                        nEndPos = InStr(i, sBuffer, "</script>", vbTextCompare)
                        If nEndPos = 0 Then nEndPos = nBufferLen + 1
                        
                        .SelStart = i
                        .SelLength = nEndPos - (i + 1)
                        .SelColor = &H80&
                        i = nEndPos
                        blnScript = False
                        blnInTag = True
                    Else
                        '// skip to next tag
                        nEndPos = InStr(i + 1, sBuffer, "<")
                        .SelStart = i
                        If nEndPos <> 0 Then
                            nEndPos = nEndPos - 1
                        Else
                            nEndPos = nBufferLen
                        End If
                        .SelLength = nEndPos - .SelStart
                        .SelColor = 0
                        i = nEndPos
                        blnInTag = False
                    End If
                Case "="
                    sTemp = Mid$(sBuffer, i + 1, 1)
                    Select Case sTemp
                    Case Chr(34), "'"
                        sCode = sTemp
                    Case Else
                        sCode = " "
                    End Select
                    '// find end of quote
                    'nEndTag = InStr(i + 2, sBuffer, ">")
                    nEndPos = InStr(i + 2, sBuffer, sCode)
                    If nEndTag < nEndPos And nEndTag <> 0 Then
                        'If sCode <> " " Then
                        '    '// tag before end of quote
                        '    GoTo SyntaxError
                        'Else
                            nEndPos = nEndTag
                        'End If
                    End If
                    If nEndPos = 0 Then nEndPos = nBufferLen
                    .SelStart = i
                    .SelLength = nEndPos - (i)
                    .SelColor = 0 '// colour it black
                    i = nEndPos
            End Select
            DoEvents
        Next i
TheEnd:
        '// sel at start
        .SelStart = 0
        Screen.MousePointer = vbDefault
        m_CancelColour = False
        tmrUpdate.Enabled = False
        RaiseEvent ProgressComplete
    End With
End Sub
Private Function RemoveBlock(strString As String)
    Dim lLineCount As Long
    Dim lResult As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim lLine As Long
    m_blnBusy = True
    SendMessage rtfMain.hWnd, WM_SETREDRAW, False, 0&
    SelectBlock
    rtfTemp.TextRTF = rtfMain.SelRTF
    lLineCount = SendMessage(rtfTemp.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lLine = 0 To lLineCount - 1
        lStart = SendMessage(rtfTemp.hWnd, EM_LINEINDEX, lLine, 0&)
        If strString = Space(m_Indent) Then
            If Mid$(rtfTemp.Text, lStart + 1, Len(strString)) = strString Then
                rtfTemp.SelStart = lStart ' - 1
                rtfTemp.SelLength = Len(strString)
                rtfTemp.SelText = ""
            End If
        Else
            lResult = GetFirstChar(lStart + 1, strString)
            If lResult >= 0 Then
                rtfTemp.SelStart = lStart + lResult - 1
                rtfTemp.SelLength = Len(strString)
                rtfTemp.SelText = ""
            End If
        End If
    Next
    If strString = "'" Or strString = "//" Then
        RunColorize
    End If
    rtfTemp.SelStart = 0
    rtfTemp.SelLength = TempCharacterCount
    lStart = rtfMain.SelStart
    rtfMain.SelRTF = rtfTemp.SelRTF
    lEnd = rtfMain.SelStart - lStart
    rtfMain.SelStart = lStart
    rtfMain.SelLength = lEnd
    SendMessage rtfMain.hWnd, WM_SETREDRAW, True, 0&
    rtfMain.Refresh
    m_blnBusy = False
    m_blnChanged = True
End Function
Public Sub CommentBlock()
    Select Case Mode
'    Case vbwVB, vbwVBScript
'        InsertBlock ("'")
'    Case vbwJava, vbwJavaScript, vbwC
'        InsertBlock ("//")
    End Select
End Sub
Public Sub UncommentBlock()
    Select Case Mode
'    Case vbwVB, vbwVBScript
'        RemoveBlock ("'")
'    Case vbwJava, vbwJavaScript, vbwC
'        RemoveBlock ("//")
    End Select
End Sub
Public Sub Indent()
    If IsNotCode Then
        InsertBlock (Chr(vbKeyTab))
        
    Else
        InsertBlock (Space(m_Indent))
    End If
End Sub
Public Sub Outdent()
    If IsNotCode Then
        RemoveBlock (Chr(vbKeyTab))
    Else
        RemoveBlock (Space(m_Indent))
    End If
End Sub
Private Sub SelectBlock()
    Dim lStart As Long
    Dim lEnd As Long
    Dim lNewEnd As Long
    Dim lNewStart As Long
    With rtfMain
        '// ensure we have selected the whole line
        lEnd = .SelLength
        lStart = .SelStart
        '// go to beginning of line
        lNewStart = SendMessage(.hWnd, EM_LINEINDEX, .GetLineFromChar(.SelStart), 0&)
        '// extend selection as needed
        'lNewEnd = SendMessage(.hwnd, EM_LINEINDEX, .GetLineFromChar(.SelStart + lEnd), 0&) + SendMessage(.hwnd, EM_LINELENGTH, .GetLineFromChar(.SelStart + lEnd), 0&)  '(lStart - .SelStart) + lEnd
        lNewEnd = lEnd + (lStart - lNewStart)
        .SelStart = lNewStart
        .SelLength = lNewEnd
        DoEvents
        'lEnd = .SelLength
        '.SelStart .SelLength - lineindex
        '.SelLength = SendMessage(.hwnd, EM_LINEINDEX, .GetLineFromChar(.SelStart + .SelLength), 0&) + SendMessage(.hwnd, EM_LINELENGTH, .GetLineFromChar(.SelStart + .SelLength), 0&)
    End With
End Sub
Private Function InsertBlock(strString As String)
    Dim lLineCount As Long
    Dim lResult As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim i As Long
    SendMessage rtfMain.hWnd, WM_SETREDRAW, False, 0&
    m_blnBusy = True
    SelectBlock
    rtfTemp.TextRTF = rtfMain.SelRTF
    With rtfTemp
        
        lLineCount = SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)
        For i = 0 To lLineCount - 1
            .SelStart = SendMessage(.hWnd, EM_LINEINDEX, i, 0&) '+ 1
            .SelText = strString
        Next
        .SelStart = 0
        .SelLength = TempCharacterCount
        lStart = rtfMain.SelStart
        rtfMain.SelRTF = .SelRTF
    End With
    lEnd = rtfMain.SelStart - lStart

    rtfMain.SelStart = lStart
    rtfMain.SelLength = lEnd
    '// comment
'    If strString = "'" Or strString = "//" Then
'        Select Case Mode
'        Case vbwVB, vbwVBScript
'            rtfMain.SelColor = m_Colour_Comment
'        Case vbwJava, vbwJavaScript, vbwC
'            rtfMain.SelColor = m_Colour_Comment
'        End Select
'    End If
    SendMessage rtfMain.hWnd, WM_SETREDRAW, True, 0&
    rtfMain.Refresh
    m_blnChanged = True
    m_blnBusy = False
End Function
Private Function GetFirstChar(lStart As Long, sString As String, Optional ByVal sText As String) As Long
    Dim lLen As Long
    Dim i As Long
    lLen = Len(sString)
    If lLen = 0 Then lLen = 1
    If IsMissing(sText) Or sText = "" Then sText = rtfTemp.Text
    For i = lStart To Len(sText)
        Select Case Mid$(sText, i, lLen)
        Case sString
            GetFirstChar = i - lStart + 1
            Exit Function
        Case " ", Chr(vbKeyTab)
            
        Case Else
            GetFirstChar = -1
            '// first letter
            Exit Function
        End Select
    Next
    GetFirstChar = -2
End Function
Private Function GetIndent(strLine As String) As String
Dim i As Long
    For i = 1 To Len(strLine)
        Select Case Mid$(strLine, i, 1)
        Case " ", Chr(vbKeyTab)
            GetIndent = GetIndent & Mid$(strLine, i, 1)
        Case Else
            '// first letter
            Exit Function
        End Select
    Next
End Function
Private Function GetLineText2(rtf As Object) As String
    Dim lCurLine As Long
    Dim lStart As Long
    Dim lLen As Long
    Dim strSearch As String
    '// Get current line
    lCurLine = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, rtf.SelStart, 0&)
    '// Set the start pos at the beginning of the line
    lStart = SendMessage(rtf.hWnd, EM_LINEINDEX, lCurLine, 0&) + 1
    If Err Then
        '// Line does not exist
        GetLineText2 = ""
        Exit Function
    End If
    '// Get the length of the line
    lLen = SendMessage(rtf.hWnd, EM_LINELENGTH, lStart, 0&) ' + 1
    '// Select the line
    GetLineText2 = Mid$(rtf.Text, lStart, lLen)
End Function
Private Function GetLineText3(rtf As Object, lStart As Long) As String
    Dim lCurLine As Long
    Dim lLen As Long
    Dim strSearch As String
    '// Get current line
    lCurLine = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, rtf.SelStart, 0&)
    '// Set the start pos at the beginning of the line
    lStart = SendMessage(rtf.hWnd, EM_LINEINDEX, lCurLine, 0&) + 1
    If Err Then
        '// Line does not exist
        GetLineText3 = ""
        Exit Function
    End If
    '// Get the length of the line
    lLen = SendMessage(rtf.hWnd, EM_LINELENGTH, lStart, 0&) ' + 1
    '// Select the line
    GetLineText3 = Mid$(rtf.Text, lStart, lLen)
End Function
'////////////////////////////////////////////////////////////
'// Private Events
'////////////////////////////////////////////////////////////



Private Sub rtfMain_Change()
    Modified = True
    m_blnChanged = True
    'If Not m_blnBusy Then RaiseEvent Change
End Sub
'Private Sub rtfMain_Click()
'    RaiseEvent Click
'End Sub
'Private Sub rtfMain_DblClick()
'    RaiseEvent DblClick
'End Sub
'Private Sub rtfMain_GotFocus()
'    'm_blnChanged = True
'End Sub
Private Function IsNotCode() As Boolean
'    If m_Mode = vbwText Or m_Mode = vbwrtf Then IsNotCode = True
End Function
Private Function IsVB() As Boolean
'    If m_Mode = vbwVB Or m_Mode = vbwVBScript Then IsVB = True
End Function
Private Function IsMoveKey(KeyCode As Integer) As Boolean
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
        IsMoveKey = True
    End Select
End Function
Private Sub rtfMain_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim blnCancel As Boolean
    Dim strIndent As String
    Dim lNextCloseTag As Long
    Dim lNextOpenTag As Long
    Dim lLastCloseTag As Long
    Dim sText As String
    Dim lStart As Long
    Dim sTemp As String
    blnJustOutTag = False
    blnJustDeletingTag = False
    On Error GoTo ErrHandler
    RaiseEvent KeyDown(KeyCode, Shift)
    'If Not IsMoveKey(KeyCode) Then GetStatus
    GetStatus
    'SetStatusBar Str(Status)
    Select Case Shift
    Case 0
        Select Case KeyCode
        Case vbKeyReturn
            If m_AutoIndent = False Or IsNotCode Then Exit Sub
            m_blnBusy = True
            '// autoindent
            sTemp = GetLineText2(rtfMain)
            strIndent = GetIndent(sTemp)
            If GetFirstChar(1, "", sTemp) = -1 Then
                m_blnBusy = False
            End If
            rtfMain.SelText = vbCrLf & strIndent
            KeyCode = 0
            m_blnBusy = False
           '
        Case vbKeyTab
            If Not IsVB Then Exit Sub
            
            If rtfMain.SelLength > 0 Then
                Indent
            Else
                rtfMain.SelText = Space(m_Indent)
            End If
            
            KeyCode = 0
        Case vbKeyBack, vbKeyDelete
            
            If IsVB Then
                If rtfMain.SelStart - m_Indent + 1 <= 0 Then Exit Sub
                sText = GetLineText3(rtfMain, lStart)
                If rtfMain.SelStart + 1 - lStart - m_Indent + 1 <= 0 Then Exit Sub
                
                If Mid$(sText, rtfMain.SelStart + 1 - lStart - m_Indent + 1, m_Indent) = Space(m_Indent) Then
                    m_blnBusy = True
                    LockWindowUpdate (rtfMain.hWnd)
                    rtfMain.SelStart = rtfMain.SelStart - m_Indent
                    rtfMain.SelLength = m_Indent
                    rtfMain.SelText = ""
                    LockWindowUpdate 0
                    KeyCode = 0
                    
                    m_blnBusy = False
                End If
            ElseIf m_Mode = VBWHTML Then
                sText = rtfMain.Text
                Select Case Status
                Case inTag
                    Select Case Mid$(sText, rtfMain.SelStart, 1)
                    Case ">", "<"
                        blnJustDeletingTag = True
                    End Select
            
                End Select
            End If
        End Select
    Case vbCtrlMask
        Select Case KeyCode
        
        Case vbKeyV
            RaiseEvent BeforePaste(blnCancel)
            If Not blnCancel Then Paste
            KeyCode = 0
        Case vbKeyC
            RaiseEvent BeforeCopy(blnCancel)
            If Not blnCancel Then Copy
            KeyCode = 0
        Case vbKeyX
            RaiseEvent BeforeCut(blnCancel)
            If Not blnCancel Then Cut
            KeyCode = 0
        End Select
    Case vbShiftMask
        If KeyCode = 16 Then Exit Sub
        Select Case KeyCode
        Case vbKeyTab
            If Not IsVB Then Exit Sub
            Outdent
            KeyCode = 0
        '// open tag - <
        Case 188
            
            DoEvents
            If Mode <> VBWHTML Then Exit Sub
            m_blnBusy = True
            sText = rtfMain.Text
            SaveCursorPos
            '// ok
            lNextCloseTag = InStr(rtfMain.SelStart + 1, sText, ">")
            LockMain
            If lNextCloseTag <> 0 Then
                rtfMain.SelLength = lNextCloseTag - rtfMain.SelStart
                rtfMain.SelColor = &HBF0000
                ColourTag rtfMain.SelStart, lNextCloseTag
            End If
            RestoreCursorPos
            rtfMain.SelLength = 0
            rtfMain.SelColor = &HBF0000
            UnLockMain
            m_blnBusy = False
        '// close tag
        Case 190
            blnJustOutTag = True
        Case 50
            Select Case Status
            Case inTag
                If blnInQuote Then
                    rtfMain.SelText = """"
                    rtfMain.SelColor = &HBF0000
                    KeyCode = 0
                    blnInQuote = False
                Else
                    rtfMain.SelColor = vbBlack
                    blnInQuote = True
                End If
            End Select
        End Select
    End Select
    Exit Sub
ErrHandler:
'    ErrHandler Err, Error, "ColourVB.KeyDown"
    m_blnBusy = False
    UnLockMain
End Sub
Private Sub rtfMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub rtfMain_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler
    RaiseEvent KeyUp(KeyCode, Shift)
    Dim lNextOpenTag As Long
    Dim lLastCloseTag As Long
    Dim lNextCloseTag As Long
    Dim lLastOpenTag As Long
    Dim nEndPos As Long
    Dim nEndTag As Long
    Dim i As Long
    Dim sTemp As String
    Dim sCode As String
    Dim OldStatus As Status
    '// save the old status
    OldStatus = Status
    GetStatus
'    SetStatusBar Str(Status)
    If OldStatus <> Status Then
        If IsMoveKey(KeyCode) Then Exit Sub
        '// status has changed
'        Select Case Mode
'        Case vbwJava, vbwJavaScript, vbwC
'            Select Case OldStatus
'            Case InComment
'
'                '// we are now out of a comment
'                '// colour from cursor point
'                sText = rtfMain.Text
'                If Mid$(sText, rtfMain.SelStart + 1, 1) = "*" Then
'                    lLastOpenTag = rtfMain.SelStart
'                ElseIf Mid$(sText, rtfMain.SelStart, 1) = "/" Then
'                    lLastOpenTag = rtfMain.SelStart - 1
'                Else
'                    Exit Sub
'                End If
'
'               ' lLastOpenCommentTag = InStrRev(Left$(sText, rtfMain.SelStart + 1), "/*")
'               ' lLastCloseCommentTag = InStrRev(Left$(sText, rtfMain.SelStart + 1), "*/")
'                'If lLastCloseCommentTag = 0 Then lLastCloseCommentTag = CharacterCount
'                '// in comment
'
'
'
'                m_blnBusy = True
'                lNextCloseTag = InStr(rtfMain.SelStart + 1, sText, "*/") + 1
'                If lNextCloseTag = 1 Then lNextCloseTag = CharacterCount + 1
'                SaveCursorPos
'                LockMain
''                If lLastOpenCommentTag > lLastCloseCommentTag Then
''                    Status = InComment
''                    If rtfMain.SelLength = 0 Then rtfMain.SelColor = m_Colour_Comment
''                Else
''                    Status = OutComment
''                    If rtfMain.SelLength = 0 Then rtfMain.SelColor = 0
''                End If
'
'                rtfMain.SelStart = lLastOpenTag
'                rtfMain.SelLength = lNextCloseTag - rtfMain.SelStart
'                rtfTemp.Text = rtfMain.SelText
'                ColorizeWordsJava
'                rtfTemp.SelStart = 0
'                rtfTemp.SelLength = TempCharacterCount
'                rtfMain.SelRTF = rtfTemp.SelRTF
'                RestoreCursorPos
'                UnLockMain
'                m_blnBusy = False
'            Case OutComment
'                '// we are now in comment
'                '// colour green
'
'                sText = rtfMain.Text
'                If rtfMain.SelStart < 3 Then
'                    lLastOpenTag = 0
'                Else
'                    If Mid$(sText, rtfMain.SelStart - 1, 2) = "/*" Then
'                        lLastOpenTag = rtfMain.SelStart - 2
'                    ElseIf Mid$(sText, rtfMain.SelStart, 2) = "/*" Then
'                        lLastOpenTag = rtfMain.SelStart - 1
'                    Else
'                        Exit Sub
'                    End If
'                End If
'                m_blnBusy = True
'                LockMain
'                lNextCloseTag = InStr(rtfMain.SelStart + 1, sText, "*/") + 1
'                If lNextCloseTag = 1 Then lNextCloseTag = CharacterCount
'                SaveCursorPos
'                rtfMain.SelStart = lLastOpenTag
'                rtfMain.SelLength = lNextCloseTag - lLastOpenTag
'                rtfMain.SelColor = m_Colour_Comment
'                RestoreCursorPos
'                    'rtfTempText = rtfMain.SelText
'                    'ColorizeWordsJava
'                m_blnBusy = False
'                    UnLockMain
'                'lLastOpenTag = InStrRev(Left$(sText, rtfMain.SelStart), "/*")
'
'                'lNextOpenTag = InStrRev
'            End Select
'        End Select
    End If
    Exit Sub
    Select Case KeyCode
    Case 190
        If Mode <> VBWHTML Then Exit Sub
        If Status = OutTag And blnJustOutTag = True Then
            m_blnBusy = True
            sText = rtfMain.Text
            SaveCursorPos
            LockMain
            rtfMain.SelColor = vbBlack
            
            '// colour tag
'            lLastOpenTag = InStrRev(Left$(sText, rtfMain.SelStart), "<")
            rtfMain.SelStart = lLastOpenTag - 1
            rtfMain.SelLength = m_lngSelPos - rtfMain.SelStart
            rtfMain.SelColor = &HBF0000
            
            DoEvents
            ColourTag lLastOpenTag, m_lngSelPos
            
            rtfMain.SelStart = m_lngSelPos
            lNextOpenTag = InStr(rtfMain.SelStart - 1, sText, "<")
            If lNextOpenTag <> 0 Then
                rtfMain.SelLength = lNextOpenTag - 1 - rtfMain.SelStart
                rtfMain.SelColor = vbBlack
            End If
            RestoreCursorPos
            UnLockMain
            m_blnBusy = False
        End If
    Case vbKeyBack, vbKeyDelete
        If Mode <> VBWHTML Then Exit Sub
        If blnJustDeletingTag = False Then Exit Sub
        On Error Resume Next
        m_blnBusy = True
        sText = rtfMain.Text
        '// we are deleting a tag
        If Status = OutTag Then
            
            SaveCursorPos
            m_blnBusy = True
            LockMain
            lNextOpenTag = InStr(rtfMain.SelStart + 1, sText, "<")
            If lNextOpenTag = 0 Then lNextOpenTag = Len(sText)
            rtfMain.SelLength = lNextOpenTag - 1 - rtfMain.SelStart
            rtfMain.SelColor = vbBlack
'            lLastCloseTag = InStrRev(Left$(sText, rtfMain.SelStart), ">")
            rtfMain.SelStart = lLastCloseTag
            
            rtfMain.SelLength = m_lngSelPos - rtfMain.SelStart
            rtfMain.SelColor = vbBlack
            RestoreCursorPos
            UnLockMain
            m_blnBusy = False
        ElseIf Status = inTag Then
            LockMain
            SaveCursorPos
            lLastOpenTag = InStrRev(Left$(sText, rtfMain.SelStart), "<")

            lNextCloseTag = InStr(rtfMain.SelStart - 1, sText, ">")
            lNextOpenTag = InStr(rtfMain.SelStart - 1, sText, "<")
            rtfMain.SelStart = lLastOpenTag - 1
            rtfMain.SelLength = m_lngSelPos - rtfMain.SelStart
            rtfMain.SelColor = &HBF0000
            If lNextCloseTag < lNextOpenTag Then
                ColourTag lLastOpenTag, m_lngSelPos
            Else
                rtfMain.SelStart = lLastOpenTag - 1
                rtfMain.SelLength = m_lngSelPos - rtfMain.SelStart
                rtfMain.SelColor = vbBlack
            End If
            RestoreCursorPos
            UnLockMain
        End If
        blnJustDeletingTag = False
        m_blnBusy = False
    End Select
        Exit Sub
ErrHandler:
'    ErrHandler Err, Error, "ColourVB.KeyUp"
    m_blnBusy = False
    UnLockMain
End Sub
Private Sub ColourTag(ByVal lStart As Long, ByVal lEnd As Long)
    Dim sTemp As String
    Dim sCode As String
    Dim nEndPos As Long
    Dim nEndTag As Long
    Dim i As Long
    If lStart = 0 Then lStart = 1
    For i = lStart To lEnd
        Select Case Mid$(sText, i, 1)
        Case "="
            sTemp = Mid$(sText, i + 1, 1)
            Select Case sTemp
            Case Chr(34), "'"
                sCode = sTemp
            Case Else
                sCode = " "
            End Select
            '// find end of quote
            nEndTag = InStr(i + 2, sText, ">")
            nEndPos = InStr(i + 2, sText, sCode)
            If nEndTag < nEndPos Then
                'If sCode <> " " Then
                '    '// tag before end of quote
                '    GoTo SyntaxError
                'Else
                    nEndPos = nEndTag
                'End If
            End If
            If nEndPos = 0 Then
                nEndPos = lEnd - 1
            End If
            rtfMain.SelStart = i
            rtfMain.SelLength = nEndPos - (i)
            rtfMain.SelColor = 0 '// colour it black
            i = nEndPos
        End Select
    Next
End Sub
'Private Sub rtfMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent MouseDown(Button, Shift, x, y)
'End Sub
'Private Sub rtfMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent MouseMove(Button, Shift, x, y)
'End Sub
'Private Sub rtfMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent MouseUp(Button, Shift, x, y)
'End Sub
'
'Private Sub rtfMain_OLECompleteDrag(Effect As Long)
'    RaiseEvent OLECompleteDrag(Effect)
'End Sub
'
'Private Sub rtfMain_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
'End Sub
'
'Private Sub rtfMain_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
'    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, state)
'End Sub
'
'Private Sub rtfMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
'End Sub
'
'Private Sub rtfMain_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
'    RaiseEvent OLESetData(Data, DataFormat)
'End Sub
'
'Private Sub rtfMain_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
'    RaiseEvent OLEStartDrag(Data, AllowedEffects)
'End Sub


Private Sub rtfMain_SelChange()
On Error GoTo ErrHandler
    Static blnMoved As Boolean
    
    If m_blnBusy Or bBusy Then Exit Sub
    RaiseEvent SelChange
'    Select Case Mode
'    Case vbwVB, vbwVBScript, vbwJava, vbwJavaScript, vbwC
'        '// Save current position
'        SaveCursorPos
'        If m_blnChanged = False Then
'            m_lngLastLine = SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
'            m_lngLastLinePos = rtfMain.SelStart + rtfMain.SelLength
'            Exit Sub
'        End If
'        m_blnBusy = True
'        If Not SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, m_lngLastLinePos, 0&) = SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&) Then
'
'            LockMain
'            '// Go back to the last line we were on
'            rtfMain.SelStart = m_lngLastLinePos
'            rtfMain.SelLength = 0
'            DoEvents
'            '// Colour it
'            blnAll = False
'            ColourCurLine
'        End If
'        m_lngLastLine = SendMessage(rtfMain.hWnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
'        m_lngLastLinePos = rtfMain.SelStart + rtfMain.SelLength
'        m_blnBusy = False
'        m_blnChanged = False
'    End Select
    Exit Sub
ErrHandler:
'    ErrHandler Err, Error, "ColourVB.SelChange"
End Sub
Private Sub tmrUpdate_Timer()
    RaiseEvent ProgressChange(Int((i / nBufferLen) * 100))
End Sub
Private Sub GetStatus()
On Error GoTo ErrHandler
    Dim lLastOpenCommentTag As Long
    Dim lLastCloseCommentTag As Long
    Dim sText As String
    sText = rtfMain.Text
    Select Case Mode
    Case VBWHTML
        
        Dim lLastOpenTag As Long
        Dim lLastCloseTag As Long
        Dim lLastOpenScriptTag As Long
        Dim lLastCloseScriptTag As Long
        Dim sQuote As String
        Dim bInQuote As Boolean
        Dim i As Long
        Dim lResult As Long
        Dim lStart As Long
        Dim lLen As Long
        lStart = rtfMain.SelStart
        lLen = CharacterCount
        '// get positions
        lLastOpenTag = InStrRev(Left$(sText, lStart), "<")
        lLastCloseTag = InStrRev(Left$(sText, lStart), ">")
        If lLastCloseTag = 0 Then lLastCloseTag = lLen
        lLastOpenCommentTag = InStrRev(Left$(sText, lStart), "<!--")
        lLastCloseCommentTag = InStrRev(Left$(sText, lStart), "-->")
        If lLastCloseCommentTag = 0 Then lLastCloseCommentTag = lLen
        lLastOpenScriptTag = InStrRev(Left$(sText, lStart), "<script", , vbTextCompare)
        lLastCloseScriptTag = InStrRev(Left$(sText, lStart), "</script>", , vbTextCompare)
        'If lLastCloseScriptTag = 0 Then lLastCloseScriptTag = lLen
        
        '// check for script first
        If lLastOpenScriptTag > lLastCloseScriptTag Then
            Status = InScript
            If rtfMain.SelLength = 0 Then rtfMain.SelColor = &H80&
            Exit Sub
        End If
        '// then for comment
        If lLastOpenCommentTag > lLastCloseCommentTag Then
            Status = InComment
            If rtfMain.SelLength = 0 Then rtfMain.SelColor = &H7F7F7F
            Exit Sub
        End If
        '// then for normal tag
        If lLastOpenTag > lLastCloseTag Then
            '// we are in a tag
            Status = inTag
            DoEvents
            '// see if we are in a quote
            sQuote = Mid$(sText, lLastOpenTag, lStart + 1 - lLastOpenTag)
            bInQuote = False
            i = 1
            Do
                lResult = InStr(i, sQuote, """")
                If lResult = 0 Then Exit Do
                bInQuote = Not bInQuote
                i = lResult + 1
            Loop
            blnInQuote = bInQuote
            If bInQuote Then
'                SetStatusBar "InTag, InQuote"
                If rtfMain.SelLength = 0 Then rtfMain.SelColor = vbBlack
            Else
'                SetStatusBar "InTag, OutQuote"
                If rtfMain.SelLength = 0 Then rtfMain.SelColor = &HBF0000
            End If
        Else
'            SetStatusBar "OutTag, OutQuote"
            Status = OutTag
            If rtfMain.SelLength = 0 Then rtfMain.SelColor = vbBlack
            '// we are not in a tag
            DoEvents
        End If
'    Case vbwJava, vbwJavaScript, vbwC
'        lLastOpenCommentTag = InStrRev(Left$(sText, rtfMain.SelStart + 1), "/*")
'        lLastCloseCommentTag = InStrRev(Left$(sText, rtfMain.SelStart + 1), "*/")
'        'If lLastCloseCommentTag = 0 Then lLastCloseCommentTag = CharacterCount
'        '// in comment
'        If lLastOpenCommentTag > lLastCloseCommentTag Then
'            Status = InComment
'            If rtfMain.SelLength = 0 Then rtfMain.SelColor = m_Colour_Comment
'        Else
'            Status = OutComment
'            If rtfMain.SelLength = 0 Then rtfMain.SelColor = 0
'        End If
    End Select
    Exit Sub
ErrHandler:
'    ErrHandler Err, Error, "ColourVB.KeyDown"
End Sub
'////////////////////////////////////////////////////////////
'// Subclassing
'////////////////////////////////////////////////////////////

Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '// scroll events
    Select Case iMsg
    Case WM_VSCROLL
        RaiseEvent VScroll
    Case WM_HSCROLL
        RaiseEvent HScroll
    End Select
End Function
Private Sub pAttachMessages()
    On Error Resume Next
    AttachMessage Me, rtfMain.hWnd, WM_VSCROLL
    AttachMessage Me, rtfMain.hWnd, WM_HSCROLL
End Sub
Private Sub pDetachMessages()
    On Error Resume Next
    DetachMessage Me, rtfMain.hWnd, WM_VSCROLL
    DetachMessage Me, rtfMain.hWnd, WM_HSCROLL
End Sub

'////////////////////////////////////////////////////////////
'// UserControl
'////////////////////////////////////////////////////////////

Private Sub UserControl_Initialize()
    pAttachMessages
    Mode = VBWHTML
End Sub
Private Sub UserControl_Terminate()
    pDetachMessages
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    rtfMain.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
'// Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    With rtfMain
        .Text = m_def_Text
        .TabStop = False
        .ToolTipText = m_def_ToolTipText
'        .RightMargin = m_def_RightMargin
        .Locked = m_def_Locked
        .OLEDropMode = rtfOLEDropNone
        .OLEDragMode = rtfOLEDragManual
        .Enabled = True
        .BulletIndent = m_def_BulletIndent
        .BackColor = &H80000005
        .AutoVerbMenu = m_def_AutoVerbMenu
    End With
    Set Font = Ambient.Font
    Mode = 1
    
    m_Colour_Comment = m_def_Colour_Comment
    m_Colour_Keyword = m_def_Colour_Keyword
    m_Colour_Text = m_def_Colour_Text
    
    BorderStyle = Abs(m_def_Border)
    m_Saved = m_def_Saved
    m_Modified = m_def_Modified
    m_CancelColour = m_def_CancelColour
    m_AutoIndent = m_def_AutoIndent
    m_Indent = m_def_Indent
    m_IgnoreVBHeader = True
End Sub
'// Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With rtfMain
        Set .MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
        .Text = PropBag.ReadProperty("Text", m_def_Text)
        .TabStop = PropBag.ReadProperty("TabStop", False)
        '.ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
        '.RightMargin = PropBag.ReadProperty("RightMargin", m_def_RightMargin)
        .Locked = PropBag.ReadProperty("Locked", m_def_Locked)
        '.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
        '.OLEDragMode = PropBag.ReadProperty("OLEDragMode", m_def_OLEDragMode)
        .Enabled = PropBag.ReadProperty("Enabled", True)
        '.BulletIndent = PropBag.ReadProperty("BulletIndent", m_def_BulletIndent)
        '.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
        '.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", m_def_AutoVerbMenu)
    End With
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Mode = PropBag.ReadProperty("Mode", 0)
    
    m_Colour_Comment = PropBag.ReadProperty("Colour_Comment", m_def_Colour_Comment)
    m_Colour_Keyword = PropBag.ReadProperty("Colour_Keyword", m_def_Colour_Keyword)
    m_Colour_Text = PropBag.ReadProperty("Colour_Text", m_def_Colour_Text)
    
    BorderStyle = Abs(PropBag.ReadProperty("Border", m_def_Border))
    'If Mode = vbwVB Then ColourVB
    m_Saved = PropBag.ReadProperty("Saved", m_def_Saved)
    m_Modified = PropBag.ReadProperty("Modified", m_def_Modified)
    m_CancelColour = PropBag.ReadProperty("CancelColour", m_def_CancelColour)
    m_AutoIndent = PropBag.ReadProperty("AutoIndent", m_def_AutoIndent)
    m_Indent = PropBag.ReadProperty("Indent", m_def_Indent)
    m_FileName = PropBag.ReadProperty("FileName", "")
    m_IgnoreVBHeader = PropBag.ReadProperty("IgnoreVBHeader", True)
End Sub

'// Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With rtfMain
        Call PropBag.WriteProperty("MouseIcon", .MouseIcon, Nothing)
        'Call PropBag.WriteProperty("AutoVerbMenu", .AutoVerbMenu, m_def_AutoVerbMenu)
        'Call PropBag.WriteProperty("BackColor", .BackColor, &H80000005)
        'Call PropBag.WriteProperty("BulletIndent", .BulletIndent, m_def_BulletIndent)
        Call PropBag.WriteProperty("Enabled", .Enabled, True)
        Call PropBag.WriteProperty("Locked", .Locked, m_def_Locked)
        'Call PropBag.WriteProperty("MultiLine", .MultiLine, m_def_MultiLine)
        'Call PropBag.WriteProperty("OLEDragMode", .OLEDragMode, m_def_OLEDragMode)
        'Call PropBag.WriteProperty("OLEDropMode", OLEDropMode, 0)
        'Call PropBag.WriteProperty("RightMargin", .RightMargin, m_def_RightMargin)
        Call PropBag.WriteProperty("Text", .Text, m_def_Text)
        'Call PropBag.WriteProperty("ToolTipText", .ToolTipText, m_def_ToolTipText)
        Call PropBag.WriteProperty("TabStop", .TabStop, False)
    End With
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Mode", Mode, 0)
    
    Call PropBag.WriteProperty("Colour_Comment", m_Colour_Comment, m_def_Colour_Comment)
    Call PropBag.WriteProperty("Colour_Keyword", m_Colour_Keyword, m_def_Colour_Keyword)
    Call PropBag.WriteProperty("Colour_Text", m_Colour_Text, m_def_Colour_Text)
    
    Call PropBag.WriteProperty("Border", BorderStyle, m_def_Border)
    Call PropBag.WriteProperty("Saved", m_Saved, m_def_Saved)
    Call PropBag.WriteProperty("Modified", m_Modified, m_def_Modified)
    Call PropBag.WriteProperty("CancelColour", m_CancelColour, m_def_CancelColour)
    Call PropBag.WriteProperty("AutoIndent", m_AutoIndent, m_def_AutoIndent)
    Call PropBag.WriteProperty("Indent", m_Indent, m_def_Indent)
    Call PropBag.WriteProperty("FileName", m_FileName, "")
    Call PropBag.WriteProperty("IgnoreVBHeader", m_IgnoreVBHeader, True)
End Sub

Public Property Get FontBackColour() As OLE_COLOR
Dim tCF2 As CHARFORMAT2
Dim lR As Long
    tCF2.dwMask = CFM_BACKCOLOR
    tCF2.cbSize = Len(tCF2)
    lR = SendMessage(rtfMain.hWnd, EM_GETCHARFORMAT, SCF_SELECTION, tCF2)
    FontBackColour = tCF2.crBackColor
End Property
Public Property Let FontBackColour(ByVal oColor As OLE_COLOR)
Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If oColor = -1 Then
      tCF2.dwMask = CFM_BACKCOLOR
      tCF2.dwEffects = CFE_AUTOBACKCOLOR
      tCF2.crBackColor = -1
   Else
      tCF2.dwMask = CFM_BACKCOLOR
      tCF2.crBackColor = TranslateColor(oColor)
   End If
   tCF2.cbSize = Len(tCF2)
   lR = SendMessage(rtfMain.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, tCF2)
End Property


Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function


'// Sets the View Mode
Public Sub SetViewMode(ByVal eViewMode As ERECViewModes)
   Select Case eViewMode
   Case ercWYSIWYG
      On Error Resume Next
      SendMessageLong rtfMain.hWnd, EM_SETTARGETDEVICE, Printer.hDC, Printer.Width
   Case ercWordWrap
      SendMessageLong rtfMain.hWnd, EM_SETTARGETDEVICE, 0, 0
   Case ercDefault
      SendMessageLong rtfMain.hWnd, EM_SETTARGETDEVICE, 0, 1
   End Select
End Sub

Public Function GetLineFromChar(lChar As Long) As Long
    GetLineFromChar = rtfMain.GetLineFromChar(lChar)
End Function
Public Sub GetPosFromChar(ByVal lIndex As Long, ByRef xPixels As Long, ByRef yPixels As Long)
Dim lxy As Long
   lxy = SendMessageLong(rtfMain.hWnd, EM_POSFROMCHAR, lIndex, 0)
   xPixels = (lxy And &HFFFF&)
   yPixels = (lxy \ &H10000) And &HFFFF&
End Sub
Public Sub Mark()
    If FontBackColour = &H8080FF Then
        FontBackColour = vbWhite
    Else
        FontBackColour = &H8080FF
    End If
    RaiseEvent SelChange
End Sub
Public Sub ClearMarks()
    SaveCursorPos
    LockMain
    rtfMain.SelStart = 0
    rtfMain.SelLength = CharacterCount
    FontBackColour = vbWhite
    RestoreCursorPos
    UnLockMain
End Sub
Public Sub SetSelection(ByVal lStart As Long, ByVal lEnd As Long)
Dim tCR As CharRange
   tCR.cpMin = lStart
   tCR.cpMax = lEnd
   SendMessage rtfMain.hWnd, EM_EXSETSEL, 0, tCR
End Sub
Public Sub SetTempSelection(ByVal lStart As Long, ByVal lEnd As Long)
Dim tCR As CharRange
   tCR.cpMin = lStart
   tCR.cpMax = lEnd
   SendMessage rtfTemp.hWnd, EM_EXSETSEL, 0, tCR
End Sub
Public Function CharacterCount() As Long
    CharacterCount = SendMessageLong(rtfMain.hWnd, WM_GETTEXTLENGTH, 0, 0)
End Function
Public Function TempCharacterCount() As Long
    TempCharacterCount = SendMessageLong(rtfTemp.hWnd, WM_GETTEXTLENGTH, 0, 0)
End Function
