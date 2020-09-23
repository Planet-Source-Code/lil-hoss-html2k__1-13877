VERSION 5.00
Object = "{F1A41040-2044-11D4-B705-0008C72B926D}#1.0#0"; "HTMLTEXT.OCX"
Begin VB.Form frmDoc 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   Icon            =   "frmDoc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   7620
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   7080
   End
   Begin VB.PictureBox picLines 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6925
      Left            =   0
      ScaleHeight     =   6930
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   0
      Width           =   420
   End
   Begin htmlText.htmSyntaxBox rtfHTML 
      Height          =   6975
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12303
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDoc.frx":0442
      Text            =   ""
   End
   Begin VB.CommandButton cmdDummy 
      Caption         =   "Dummy Button"
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1695
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    HTML2K
'    Copyright (C) 2000 Matt Wunch
'
'    The author of this program offers no warranties, either
'    expressed or implied, including but not limited to any
'    implied warranties of merchantability or fitness for a
'    particular purpose, regarding this material and makes
'    such material available solely on an "as is" basis.
'    It is not warranted that the operation of the program
'    will be uninterrupted or error free.
'
'    In no event shall the author be liable to anyone for
'    special, collateral, incidental, or consequential damages
'    in connection with or arising out of the use of these
'    materials.

Option Explicit

Private WithEvents cP As cPopupMenu
Attribute cP.VB_VarHelpID = -1
Private Const mcWEBSITE = -&H8000&

' ****************************************************************************************
' ****************************************************************************************
Dim LineCountChange As Integer
Dim FirstLine As Long
Dim FirstLineNow As Long


' ****************************************************************************************
' ****************************************************************************************

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Const KeyLessThan = 60 '188

Dim pt          As POINTAPI
Dim lngStart    As Long

Public Drive As String
Public Dir As String
Public File As String
Public Ext As String

Public Modified As Boolean
Public UnSaved As Boolean
Public Filename As String
Public DocKey As String
Public TempName As String

' Used in multi-level undo/redo scheme
Public trapUndo As Boolean         ' Locks the document while
                                    ' undo/redo is performed
Public UndoStack As New Collection
Public RedoStack As New Collection

Dim NotChangingKeys() As Integer

Dim LineChanged As Boolean
Dim CurrentStart As Integer
Dim CurrentLen As Integer

' AskSave = 1 if user canceled the request
Private Function AskSave() As Integer
    If Modified Then
        Dim R As VbMsgBoxResult
        R = MsgBox("Save changes to " & Me.Caption & "?", _
            vbQuestion Or vbYesNoCancel, "Just a moment...")
        Select Case R
        Case vbYes
            If UnSaved Then
                frmMain.SaveFileAs
            Else
                rtfHTML.SaveFile Filename
            End If
            AskSave = 0
        Case vbNo
            AskSave = 0
        Case vbCancel
            AskSave = 1
            Exit Function
        End Select
    End If
    
    ' If user wishes to discard, then remove temporary
    ' file
    On Error Resume Next
    If Len(Me.TempName) > 0 Then
        TempFiles.Remove TempName
        Kill TempName
        FileDumpCol TempFiles, AppPath & "tempdoc.his"
        Me.TempName = ""
    End If
End Function

Private Sub Form_Activate()
    Set frmMain.CurrentDoc = Me
    frmMain.UpdateUndoFunctions
    rtfHtml_SelChange
    rtfHTML.Refresh
    rtfHTML.AutoColorize = True
    
    DrawNumbers
End Sub

Private Sub Form_Load()
    Dim DocCount%
    Dim DocTitle As String
    DocCount = frmMain.Documents.Count
    Set frmMain.CurrentDoc = Me

    'Add all of the keys that don't cause the
    'line to change.
    ReDim NotChangingKeys(1 To 17)

    NotChangingKeys(1) = vbKeyShift
    NotChangingKeys(2) = vbKeyControl
    NotChangingKeys(3) = vbKeyMenu
    NotChangingKeys(4) = vbKeyPause
    NotChangingKeys(5) = vbKeyCapital
    NotChangingKeys(6) = vbKeyEscape
    NotChangingKeys(7) = vbKeyPageUp
    NotChangingKeys(8) = vbKeyPageDown
    NotChangingKeys(9) = vbKeyEnd
    NotChangingKeys(10) = vbKeyHome
    NotChangingKeys(11) = vbKeyLeft
    NotChangingKeys(12) = vbKeyUp
    NotChangingKeys(13) = vbKeyRight
    NotChangingKeys(14) = vbKeyDown
    NotChangingKeys(15) = vbKeyPrint
    NotChangingKeys(16) = vbKeyInsert
    NotChangingKeys(17) = vbKeyNumlock

    Set cP = New cPopupMenu
    ' Make sure the ImageList has icons before setting
    ' this if it is a MS ImageList:
    cP.ImageList = frmMain.ilsIcons
    ' Make sure you set this up before trying any menus
    '
    cP.hWndOwner = Me.hWnd
    
    ' Cool!
    cP.GradientHighlight = True
    
    ' Create some menus and store them:
    CreateMenus

    trapUndo = True

    If Len(Filename) = 0 Then
        UnSaved = True
        Modified = False
        GetColors
        Me.Caption = "Untitled: " & (DocCount + 1)
    Else
        GetColors
        rtfHTML.Colorize
        rtfHTML.LoadFile Filename
        Me.Caption = Filename
        rtfHTML.Colorize
        Modified = False
        UnSaved = False
    End If
    
    File = ""
    rtfHtml_Change
    rtfHtml_SelChange
    DocKey = CStr(FreeNum)
    frmMain.Documents.Add Me, DocKey
    FreeNum = FreeNum + 1
    frmMain.AllDocs
    Modified = False
    DisableSaves
End Sub

Private Sub GetColors()
    On Error Resume Next
    rtfHTML.BackColor = GetSetting(ThisApp, "Options", "Background Color")
    rtfHTML.CommentColor = GetSetting(ThisApp, "Options", "Comment Color")
    rtfHTML.EntityColor = GetSetting(ThisApp, "Options", "Entity Color")
    rtfHTML.PropNameColor = GetSetting(ThisApp, "Options", "Property Color")
    rtfHTML.PropValColor = GetSetting(ThisApp, "Options", "Value Color")
    rtfHTML.TagColor = GetSetting(ThisApp, "Options", "Tag Color")
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    rtfHTML.Width = Me.ScaleWidth - picLines.Width - 60
    rtfHTML.Height = Me.ScaleHeight
    picLines.Height = Me.ScaleHeight
    picLines.Visible = True
    DrawNumbers
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frmMain.ExitProcess Then
        ' user is not exiting whole program, just
        ' this document
        Cancel = AskSave
        If Cancel = 1 Then Exit Sub
    End If
    
    frmMain.Documents.Remove DocKey
    If frmMain.Documents.Count = 0 Then
        Set frmMain.CurrentDoc = Nothing
        frmMain.NoDocs
    End If
End Sub

Private Sub rtfHtml_Change()
    Dim LineCount As Long
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim objElement, objElement2
    Dim C%, l&

    'remove all redo items because of the change
    For C% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next C%

    'set the values of the new element
    newElement.SelStart = rtfHTML.SelStart
    newElement.TextLen = Len(rtfHTML.Text)
    newElement.Text = rtfHTML.Text

' *******************************************************************************
' *******************************************************************************
'    // Get number of lines in Rtftext
    LineCount = SendMessage(rtfHTML.hWnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    If LineCount = LineCountChange Then
        GoTo Skip:    '// Line count is still the same
      Else
        DrawNumbers '// new Line count is required
    End If
' *******************************************************************************
' *******************************************************************************

Skip:

'    add it to the undo stack
    UndoStack.Add item:=newElement
    If UndoStack.Count > 100 Then UndoStack.Remove 1
    
'    enable controls accordingly
    frmMain.UpdateUndoFunctions

    Modified = True
    UpdateDisplay
    EnableSaves
End Sub

Private Sub rtfHtml_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        InsertSurroundTag ("http://www."), (".com")
    ElseIf KeyCode = vbKeyF5 Then
        rtfHTML.Colorize
        rtfHTML.Refresh
    End If
End Sub

Private Sub rtfHtml_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        
        Dim iIndex As Long
        With cP
           .Restore "vbAccelerator"
           iIndex = .ShowPopupMenu(X, Y, X, Y, X, Y)
           If (iIndex > 0) Then
'              Status "Selected Item=" & iIndex
              If (.ItemKey(iIndex) = "Web") Then
'                 mnuHelp_Click 0
              ElseIf (.ItemKey(iIndex) = "Channel") Then
'                 mnuHelp_Click 1
              ElseIf (.ItemData(iIndex) = mcWEBSITE) Then
                 Screen.MousePointer = vbHourglass
                 ShellEx .ItemKey(iIndex)
                 Screen.MousePointer = vbDefault
              End If
           End If
        End With

'        PopupMenu frmMain.mnuEdit, , X, Y
    End If
End Sub

Private Sub rtfHtml_SelChange()
    Dim Ln As Long
    
    If trapUndo = False Then Exit Sub

    Ln = rtfHTML.SelLength
    With frmMain
        .mnuEditCut.Enabled = Ln
        .mnuEditCopy.Enabled = Ln
        .mnuEditDelete.Enabled = Ln
        .mnuHTMLSourceCompact.Enabled = Ln
        .mnuHTMLSourceSimple.Enabled = Ln
        .mnuHTMLSourceHierachal.Enabled = Ln
        .Toolbar1.Buttons("CUT").Enabled = Ln
        .Toolbar1.Buttons("COPY").Enabled = Ln
        .Toolbar1.Buttons("DELETE").Enabled = Ln
    End With

    If Me.rtfHTML = "" Then
        frmMain.mnuEditSelectAll.Enabled = False
    Else
        frmMain.mnuEditSelectAll.Enabled = True
    End If
    
    UpdateDisplay
    frmMain.UpdateUndoFunctions
End Sub

' Insert a tag surrounding the selected text, or, if there is
' none, insert the tag set and put the cursor in between the
' opening and closing tags.
'
Public Sub InsertSurroundTag(BeginTag$, EndTag$)
    Dim T$
    T$ = rtfHTML.SelText

    rtfHTML.SelText = ""
    rtfHTML.SelText = BeginTag & T$ & EndTag
    If Len(T$) = 0 Then
        rtfHTML.SelStart = rtfHTML.SelStart - Len(EndTag)
    End If
End Sub

Public Sub SetModified(M As Boolean)
    Modified = M
    With frmMain
        .Toolbar1.Buttons("SAVE").Enabled = M
        .mnuFileSave.Enabled = M
        .mnuFileSaveAs.Enabled = M
    End With
End Sub

Public Sub SaveTemp()
Dim a%, B%, C$, f$
    If Len(Me.TempName) > 0 Then
        rtfHTML.SaveFile Me.TempName
        Exit Sub
    End If
    
    Randomize Timer
    
    ' Generate random filename
    a = CInt(Rnd * (Rnd * 100))
    B = CInt((vbKeyZ - vbKeyA + 1) * Rnd + vbKeyA)
    C$ = Chr$(B)
    f$ = AppPath & C$ & CStr(a) & ".tmp.html"
    
    ' Save the file
    rtfHTML.SaveFile f$
    
    ' Remember it so we can recover it after a
    ' crash if necessary
    TempFiles.Add f$, f$
    FileDumpCol TempFiles, AppPath & "tempdoc.his"
    Me.TempName = f$
End Sub

Public Sub Undo()
Dim chg$, X&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object
Dim f As frmDoc
    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            cmdDummy.SetFocus   'change focus of form
            X& = SendMessage(rtfHTML.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            rtfHTML.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            rtfHTML.SelLength = objElement.TextLen - objElement2.TextLen
            rtfHTML.SelText = ""
            X& = SendMessage(rtfHTML.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            rtfHTML.SelStart = objElement2.SelStart
            rtfHTML.SelLength = 0
            rtfHTML.SelText = chg$
            rtfHTML.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                rtfHTML.SelLength = Len(chg$)
            Else
                rtfHTML.SelStart = rtfHTML.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    frmMain.UpdateUndoFunctions
    trapUndo = True
    rtfHTML.SetFocus
End Sub

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(rtfHTML.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            rtfHTML.SelStart = objElement.SelStart
            rtfHTML.SelLength = Len(rtfHTML.Text) - objElement.TextLen
            rtfHTML.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(rtfHTML.Text, objElement.Text, objElement.SelStart + 1)
            rtfHTML.SelStart = objElement.SelStart - Len(chg$)
            rtfHTML.SelLength = 0
            rtfHTML.SelText = chg$
            rtfHTML.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                rtfHTML.SelLength = Len(chg$)
            Else
                rtfHTML.SelStart = rtfHTML.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    frmMain.UpdateUndoFunctions
    trapUndo = True
    rtfHTML.SetFocus
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
    Dim tempParam$
    Dim d&
    
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub CloseDoc()
    Unload Me
End Sub




' **********************************************************************************
' **********************************************************************************
' **********************************************************************************
Private Sub Timer1_Timer()
    DoEvents
    FirstLine = SendMessage(rtfHTML.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary
    DoEvents
    If Not FirstLineNow = FirstLine Then DrawNumbers '// I can't hook to a scrollbar so I used a sucker-timer
    DoEvents
End Sub

Sub DrawNumbers()
  Dim LineCount As Long '// How many lines in total
  Dim i As Long      '// Just an integer
  Dim TempBuf As String
  Static WidthCount As Integer
'// Get number of lines in Rtftext
    LineCount = SendMessage(rtfHTML.hWnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    '// Same lines ?
    LineCountChange = LineCount

    '// Get first visible line in rtfText
    FirstLine = SendMessage(rtfHTML.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary

    picLines.Cls '// Clear the PicLines
    picLines.CurrentY = 40  '// Move the .top text by 40 twips

    '// Print the number of each line on a picture
    For i = 0 To LineCount - FirstLine
        picLines.CurrentY = picLines.CurrentY + 7.49 '// Where on Y
        picLines.CurrentX = 20 '-2                   '// Where on X
        picLines.Print i + FirstLine + 1             '// print the number
    Next
    picLines.Refresh
    'LineCountChange = LineCount '// Remember the last line count
    FirstLineNow = FirstLine     '// Is the first visible line still the same ?
End Sub



Private Sub CreateMenus()
Dim i As Long
Dim J As Long
Dim iIndex As Long
Dim lIcon As Long
Dim sKey As String
Dim sCap As String
   
   ' Create the demo menu:
   With cP
      .Clear
      For i = 1 To 10
       If (i = 6) Or (i = 7) Then sKey = "CHECK" Else sKey = ""
         iIndex = .AddItem("Test " & i, , i, , i + 3, ((i = 6) Or (i = 7)), ((i Mod 3) <> 0), sKey)
         If (i = 5) Then
            For J = 1 To 30
               sCap = "SubMenu Test" & J
               If ((J - 1) Mod 10) = 0 And J > 1 Then
                  sCap = "|" & sCap
               End If
               .AddItem sCap, , , iIndex, J + 10
            Next J
         End If
         If (i = 4) Or (i = 5) Then
            .AddItem "-"
         End If
      Next i
      .Store "Demo"
      
      ' Create the edit menu:
      .Clear
      .AddItem "Cu&t" & vbTab & "Ctrl+X", , , , frmMain.ilsIcons.ListImages("CUT").Index - 1, , , "Cut"
      .AddItem "&Copy" & vbTab & "Ctrl+C", , , , frmMain.ilsIcons.ListImages("COPY").Index - 1, , , "Copy"
      .AddItem "&Paste" & vbTab & "Ctrl+V", , , , frmMain.ilsIcons.ListImages("PASTE").Index - 1, , False, "Paste"
      .Store "Edit"
      
      ' Create the vbAccelerator menu:
      .Clear
      .AddItem "Special Functions"
      .Header(1) = True
      .AddItem "&HTML2K on the Web..." & vbTab & "F1", , , , 58, , , "Web"
      .Default(2) = True
'      lIcon = frmMain.ilsIcons.ListImages("Web").Index - 1
      .AddItem "Add vbAccelerator Active &Channel...", , mcWEBSITE, , lIcon, , , "Channel"
      .AddItem "-Other sites"
      i = .AddItem("VB Sites", , , , lIcon)
      .AddItem "-VB Sites", , , i
      .AddItem "Goffredo's VB Page", , mcWEBSITE, i, lIcon, , , "http://www.cs.utexas.edu/users/gglaze/vb.htm"
      .AddItem "Advanced Visual Basic WebBoard", , mcWEBSITE, i, lIcon, , , "http://webboard.duke.net:8080/~avb/"
      .AddItem "VBNet", , mcWEBSITE, 2, lIcon, , , "http://www.mvps.org/mvps"
      .AddItem "CCRP", , mcWEBSITE, 2, lIcon, , , "http://www.mvps.org/ccrp"
      .AddItem "DevX", , mcWEBSITE, i, lIcon, , , "http://www.devx.com/"
      i = .AddItem("Technology", , , , lIcon)
      .AddItem "-Games", , , i
      .AddItem "Dave's Classics", , mcWEBSITE, i, lIcon, , , "http://www.davesclassics.com/"
      .AddItem "Future Gamer", , mcWEBSITE, i, lIcon, , , "http://www.futuregamer.com/"
      .AddItem "-Web Site Building", , , i
      .AddItem "Builder.com", , mcWEBSITE, i, lIcon, , , "http://www.builder.com/"
      .AddItem "The Web Design Resource", , mcWEBSITE, i, lIcon, , , "http://www.pageresource.com/"
      .AddItem "Web Review", , mcWEBSITE, i, lIcon, , , "http://www.webreview.com/"
      .AddItem "-Downloads", , , i
      .AddItem "Tucows", , mcWEBSITE, i, lIcon, , , "http://tucows.cableinet.net/"
      .AddItem "WinFiles.com", , mcWEBSITE, i, lIcon, , , "http://www.winfiles.com/"
      i = .AddItem("Searching and Other", , , , lIcon)
      .AddItem "-Pick'n'Mix", , , i
      .AddItem "The SCHWA Corporation", , mcWEBSITE, i, lIcon, , , "http://www.theschwacorporation.com/"
      .AddItem "Art Cars", , mcWEBSITE, i, lIcon, , , "http://www.artcars.com/"
      .AddItem "The Onion", , mcWEBSITE, i, lIcon, , , "http://www.theonion.com/"
      .AddItem "Virtues of a Programmer", i, mcWEBSITE, i, lIcon, , , "http://www.hhhh.org/wiml/virtues.html"
      .AddItem "-Search", , , i
      .AddItem "HotBot", , mcWEBSITE, i, lIcon, , , "http://www.hotbot.com/"
      .AddItem "DogPile", , mcWEBSITE, i, lIcon, , , "http://www.dogpile.com/"
      .Store "vbAccelerator"
      
      .Clear
      .AddItem "First Check", , , , , True, , "Check1"
      .AddItem "Second Check", , , , , , , "Check2"
      .AddItem "Third Check", , , , , , , "Check3"
      .AddItem "-"
      i = .AddItem("First Option", , , , , , , "Option1")
      .RadioCheck(i) = True
      .AddItem "Second Option", , , , , , , "Option2"
      .AddItem "Third Option", , , , , , , "Option3"
      .AddItem "Fourth Option", , , , , , , "Option4"
      .AddItem "-"
      .AddItem "&vbAccelerator on the Web...", , , , lIcon, , , "Web"
      .Store "CheckTest"
      
      .Clear
      .AddItem "&Back" & vbTab & "Alt+Left Arrow", , , , , , , "mnuAccel(0)"
      .AddItem "&Next" & vbTab & "Alt+Right Arrow", , , , , , , "mnuAccel(1)"
      .AddItem "-"
      .AddItem "&Home Page" & vbTab & "Alt+Home", , , , , , , "mnuAccel(3)"
      .AddItem "&Search the Web", , , , , , , "mnuAccel(4)"
      .AddItem "-"
      .AddItem "&Mail", , , , , , , "mnuAccel(6)"
      .AddItem "&News", , , , , , , "mnuAccel(7)"
      .AddItem "My &Computer", , , , , , , "mnuAccel(8)"
      .AddItem "A&ddress Book", , , , , , , "mnuAccel(9)"
      .AddItem "Ca&lendar", , , , , , , "mnuAccel(10)"
      .AddItem "&Internet Call", , , , , , , "mnuAccel(11)"
      .Store "AccelTest"
   End With
   
End Sub
