Attribute VB_Name = "modFunctions"
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

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26


Sub UpdateDisplay()
    Dim Row As Single, Col As Single
    Dim cPos As Long ', lpMinPos As Long, lpMaxPos As Long
    ' Display the line number that the cursor is currently on.
    CurrentLine = SendMessageLong(frmMain.CurrentDoc.rtfHTML.hwnd, EM_LINEFROMCHAR, -1&, 0&) + 1
    ' Display the total number of lines.
    TotalLines = SendMessageLong(frmMain.CurrentDoc.rtfHTML.hwnd, EM_LINETOTAL, 0&, 0&)
    CaretPosition frmMain.CurrentDoc.rtfHTML, Col, Row
    cPos = GetScrollPos(frmMain.CurrentDoc.rtfHTML.hwnd, SB_HORZ)

    frmMain.sbStatusBar.Panels("row").Text = "ln " & CurrentLine
    frmMain.sbStatusBar.Panels("col").Text = "col " & Col
    frmMain.sbStatusBar.Panels("total").Text = TotalLines
    frmMain.sbStatusBar.Refresh
End Sub

Public Function HighlightText(ctl As TextBox)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
    ctl.SetFocus
End Function

Public Sub CaretPosition(Ctrl As htmSyntaxBox, Col As Single, Row As Single)
    Dim wColNumber As Long, dwLineNumber As Long, dwLineIndex As Long
    Dim dwGetSel As Long, dwStart As Long, dwEnd As Long

    dwLineNumber = SendMessage(Ctrl.hwnd, EM_LINEFROMCHAR, -1, 0&)

    ' Send the EM_GETSEL message to the edit control.
    ' The low-order word of the return value is the character
    ' position of the caret relative to the first character in the
    ' edit control.
    dwGetSel = SendMessage(Ctrl.hwnd, EM_GETSEL, 0, 0&)
    dwStart = dwGetSel And &HFFFF&
    dwEnd = dwGetSel / &HFFFF&

    If dwStart < dwEnd Then dwStart = dwEnd

    ' Send the EM_LINEINDEX message with the value of -1 in wParam.
    ' The return value is the absolute number of characters
    ' that precede the first character in the line containing
    ' the caret.

    dwLineIndex = SendMessage(Ctrl.hwnd, EM_LINEINDEX, -1, 0&)

    ' Subtract the LineIndex from the start of the selection,
    ' and then add 1 (since the column is zero-based).
    ' This result is the column number of the caret position.
    wColNumber = dwStart - dwLineIndex

    Row = dwLineNumber + 1
    Col = wColNumber + 1
End Sub

Public Sub UpdateMRU()
    Dim i%
    
    On Error GoTo MRU_Error
    If MRUFiles.Count > 0 Then
        
        ' Remove the bottom item if necessary
        If MRUFiles.Count > 4 Then
            While MRUFiles.Count > 4
                MRUFiles.Remove 5
            Wend
        End If
        
        ' Set menu items
        For i = 1 To MRUFiles.Count
            frmMain.mnuMRU(i - 1).Caption = "&" & _
                CStr(i) & " " & MRUFiles(i)
            frmMain.mnuMRU(i - 1).Visible = True
        Next i
        frmMain.mnuMRUSep.Visible = True
        
        ' Hide unused menu items
        If MRUFiles.Count < 4 Then
            For i = MRUFiles.Count + 1 To 3
                frmMain.mnuMRU(i - 1).Visible = False
            Next i
        End If
        
        With frmMain.Toolbar1.Buttons("OPEN").ButtonMenus
            .Clear
            For i = 1 To MRUFiles.Count
                .Add , MRUFiles(i), MRUFiles(i)
            Next i
        End With
    Else
        For i = 0 To 3
            frmMain.mnuMRU(i).Visible = False
        Next
        frmMain.mnuMRUSep.Visible = False
        
        ' Clear the toolbutton dropdown menu
        frmMain.Toolbar1.Buttons("OPEN").ButtonMenus.Clear
    End If
    Exit Sub
MRU_Error:
    ErrHandler vbObjectError, "Error Updating the MRU List", "UpdateMRU", , , True
End Sub

Public Sub MRURegPoke(key$, Value$)
    SaveSetting ThisApp, "MRU List", key, Value
End Sub

Public Function MRURegPeek(key$, Optional Default) As String
    If IsMissing(Default) Then
        MRURegPeek = GetSetting(ThisApp, "MRU List", key)
    Else
        MRURegPeek = GetSetting(ThisApp, "MRU List", key, Default)
    End If
End Function

Public Sub RegPoke(key$, Value$)
    SaveSetting ThisApp, "Options", key, Value
End Sub

Public Function RegPeek(key$, Optional Default) As String
    If IsMissing(Default) Then
        RegPeek = GetSetting(ThisApp, "Options", key)
    Else
        RegPeek = GetSetting(ThisApp, "Options", key, Default)
    End If
End Function

' Checks for the existence of a string in a list
Public Function InList(l As Variant, S As String) As Boolean
Dim i%, C As New Collection, T$
    
    T$ = TypeName(l)
    
    ' If we have a combobox or listbox here, copy it over
    ' to a collection. This way we can handle all kinds
    ' of requests.
    '
    If T$ = "ComboBox" Or T$ = "ListBox" Then
        If l.ListCount = 0 Then
            InList = False
            Exit Function
        End If
        
        For i = 0 To l.ListCount - 1
            C.Add l.List(i)
        Next i
    ElseIf T$ = "Collection" Then
        Set C = l
    Else
        InList = False
        Exit Function
    End If
    
    If C.Count = 0 Then
        InList = False
        Exit Function
    Else
        ' Do a non-case-sensitive search
        For i = 1 To C.Count
            If (S = C(i)) Or ((Len(S) = Len(C(i))) And _
                InStr(1, S, C(i), 1)) Then
                    InList = True
                    Exit Function
            End If
        Next i
    End If
    
    InList = False
End Function

Public Function CutText()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmMain.CurrentDoc.rtfHTML.SelText
    ' Delete the selected text.
    frmMain.CurrentDoc.rtfHTML.SelText = ""
End Function

Public Function SelectAllText()
    ' Use SelStart & SelLength to select the text.
    On Error Resume Next
    frmMain.CurrentDoc.rtfHTML.SelStart = 0
    frmMain.CurrentDoc.rtfHTML.SelLength = Len(frmMain.CurrentDoc.rtfHTML.Text)
End Function

Public Function CopyText()
    Clipboard.SetText frmMain.CurrentDoc.rtfHTML.SelText
End Function

Public Function PasteText()
    frmMain.CurrentDoc.rtfHTML.SelText = Clipboard.GetText
End Function

' Modify a string based on whether UCaseTags is true
Public Function CaseTag(Tag As String) As String
    CaseTag = IIf(UcaseTags, UCase(Tag), LCase(Tag))
End Function

Public Sub LoadOptions()
    Dim i As Integer
    ShowSplash = RegPeek("ShowSplash", True)
    StartNewFile = RegPeek("StartNewFile", False)
    UcaseTags = RegPeek("UCaseTags", False)
    ItalicComments = RegPeek("ItalicComments", False)
    AutosaveInterval = RegPeek("AutosaveInterval", 15)
    AutoIndent = RegPeek("AutoIndent", True)
    FlatToolbars = RegPeek("FlatToolbars", True)
    
    If FlatToolbars = True Then
        frmMain.Toolbar1.Style = tbrFlat
    Else
        frmMain.Toolbar1.Style = tbrStandard
    End If
    
    If StartNewFile = True Then
        LoadNewDoc
        frmMain.CurrentDoc.WindowState = vbMaximized
        frmMain.CurrentDoc.rtfHTML.SetFocus
    End If
    
    Dim f$
    For i = 1 To 4
        f$ = MRURegPeek("MRUFile" & CStr(i), "")
        If Len(f$) > 0 Then
            MRUFiles.Add f$, f$
        End If
    Next i
    UpdateMRU
End Sub

Public Sub SaveOptions()
    RegPoke "ShowSplash", frmSettings.chkShowSplash.Value
    RegPoke "StartNewFile", frmSettings.chkNewFile.Value
    RegPoke "UCaseTags", frmSettings.chkUpperCaseTag.Value
    RegPoke "AutosaveInterval", frmSettings.txtAutosave.Text
    RegPoke "Comment Color", frmSettings.htmlOption.CommentColor
    RegPoke "Entity Color", frmSettings.htmlOption.EntityColor
    RegPoke "Property Color", frmSettings.htmlOption.PropNameColor
    RegPoke "Value Color", frmSettings.htmlOption.PropValColor
    RegPoke "Tag Color", frmSettings.htmlOption.TagColor
    RegPoke "Background Color", frmSettings.htmlOption.BackColor
    RegPoke "ItalicComments", frmSettings.chkItalic.Value
    RegPoke "AutoIndent", frmSettings.chkAutoIndent.Value
    RegPoke "FlatToolbars", frmSettings.chkFlatToolbars.Value
    RegPoke "ShowLineNumbers", frmSettings.chkLineNumbers.Value
    If FlatToolbars = True Then
        frmMain.Toolbar1.Style = tbrFlat
    Else
        frmMain.Toolbar1.Style = tbrStandard
    End If
End Sub

Public Sub FileDumpCol(C As Collection, Filename$)
    Dim i%
    
    Open Filename For Output As #1
    For i = 1 To C.Count
        If Len(C.item(i)) > 0 Then Print #1, C.item(i)
    Next i
    Close #1
End Sub

' FileFillCol and FileDumpCol provide simple line-based
' IO for saving collections to text files. Each item
' in a collection is stored as a line.
'
Public Sub FileFillCol(C As Collection, Filename$)
Dim a$
    On Error GoTo NoFile
    Open Filename For Input As #1
    
    While Not EOF(1)
        Line Input #1, a$
        If Len(a$) > 0 Then C.Add a$, a$
    Wend
    Close #1
NoFile:
End Sub

'// Function OpenURL (and any other file!)
'//
Function OpenURL(ByVal sFile As String, Optional vArgs As Variant, Optional vShow As Variant, Optional vInitDir As Variant, Optional vVerb As Variant, Optional vhWnd As Variant) As Long
    '// Fill any empty optional arguments
    If IsMissing(vArgs) Then vArgs = vbNullString
    If IsMissing(vShow) Then vShow = vbNormalFocus
    If IsMissing(vInitDir) Then vInitDir = vbNullString
    If IsMissing(vVerb) Then vVerb = vbNullString
    If IsMissing(vhWnd) Then vhWnd = 0
    '// Call the dll
    OpenURL = ShellExecute(0, vbNullString, sFile, vbNullString, vbNullString, vbNormalFocus)
    'MsgBox Shell(sFile)
End Function

Public Sub LoadCodeClips()
    Dim i$, cClip As New clsCodeClip
    If Len(Dir(AppPath & "clips.dat")) = 0 Then Exit Sub

    Open AppPath & "clips.dat" For Input As #1
    While Not EOF(1)
        Line Input #1, i$
        If i$ = "***---***" Then
            CodeClips.Add cClip, cClip.Name
            Set cClip = New clsCodeClip
        Else
            If Len(cClip.Name) = 0 Then
                cClip.Name = i$
            Else
                cClip.Code = cClip.Code & i$
            End If
        End If
    Wend
    Close #1
End Sub

Public Sub SaveCodeClips()
    Dim Clip As clsCodeClip
    Open AppPath & "clips.dat" For Output As #1
    For Each Clip In CodeClips
        Print #1, Clip.Name
        Print #1, Clip.Code
        Print #1, "***---***"
    Next Clip
    Close #1
End Sub

'  This function lets the computer make an informed
' guess at what you want while you are typing it, based
' on some kind of list that contains values you have used
' in the past.  Used in image link, hyperlink, and other places.
'
Public Sub AutoComplete(l As ComboBox)
    Dim strText As String
    Dim strLength As Integer
    Dim i%

    strText = LCase$(l.Text)
    strLength = Len(strText)
    If strLength > 0 Then
        For i = 0 To l.ListCount - 1
            If strText = LCase$(Left$(l.List(i), strLength)) Then
                ' We found something
                If Len(l.List(i)) > strLength Then
                    l.SelText = Mid$(l.List(i), strLength + 1)
                    l.SelStart = strLength
                    l.SelLength = Len(Mid$(l.List(i), strLength + 1))
                End If
                Exit For
            End If
        Next i
    End If
End Sub

Public Sub AddAllFilesInDir(strDir As String, cboCombo As ComboBox)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, lRet As Long
    On Error Resume Next
    '// get a file handle
    lFileHdl = FindFirstFile(App.Path & "\" & strDir & "\*.html", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            DoEvents
            '// if it is a file
            sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
            If sTemp <> "." And sTemp <> ".." Then cboCombo.AddItem Left$(sTemp, Len(sTemp) - 5)
            '// based on the file handle iterate through all files and dirs
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    '// close the file handle
    lRet = FindClose(lFileHdl)
End Sub

'// Removes trailing nulls from a string
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function ReplaceSubString(str As String, ByVal substr As String, ByVal newsubstr As String)
    Dim Pos As Double
    Dim startPos As Double
    Dim new_str As String

    startPos = 1
    Pos = InStr(str, substr)
    Do While Pos > 0
        new_str = new_str & Mid$(str, startPos, Pos - startPos) & newsubstr
        startPos = Pos + Len(substr)
        Pos = InStr(startPos, str, substr)
    Loop
    new_str = new_str & Mid$(str, startPos)
    ReplaceSubString = new_str
End Function

Public Function CompactFormat(target As String) As String
    Dim a As String
    
    a = ReplaceSubString(target, vbCrLf, "")
    a = ReplaceSubString(a, Chr(9), " ")
    a = ReplaceSubString(a, "     ", " ")
    a = ReplaceSubString(a, "    ", " ")
    a = ReplaceSubString(a, "   ", " ")
    a = ReplaceSubString(a, "  ", " ")
    a = Clean(a)

    CompactFormat = a
End Function

Public Function SimpleFormat(target As String) As String
    SimpleFormat = ReplaceSubString(CompactFormat(target), "><", ">" & vbCrLf & "<")
End Function

Public Function HierarchalFormat(target As String) As String
    target = ReplaceSubString(target, vbCrLf, "")
    target = ReplaceSubString(target, vbTab, "")
    
    target = Eformat(target)
    
    HierarchalFormat = Clean(target)
End Function

Private Function Clean(targ As String) As String
    targ = ReplaceSubString(targ, " >", ">")
    targ = ReplaceSubString(targ, "< ", "<")
    targ = ReplaceSubString(targ, "> <", "><")
    'targ = ReplaceSubString(targ, "> ", ">")
    'targ = ReplaceSubString(targ, " <", "<")
    Clean = targ
End Function

Private Function Eformat(str As String) As String
    On Error Resume Next

    Dim startPos As Double
    Dim endPos As Double
    Dim indentationLevel As Double
    Dim new_str As String

    indentationLevel = 0
    startPos = 0
    endPos = 0

    If (Mid$(str, 1, 1) <> "<") Then
        Dim tempEnd As Double
        tempEnd = InStr(1, str, "<")
        If tempEnd = 0 Then
            tempEnd = Len(str)
        End If
        new_str = Mid$(str, 1, tempEnd)
    End If

    Do
        DoEvents
        If InStr(startPos + 1, str, "</") <> 0 And InStr(startPos + 1, str, "</") <= InStr(startPos + 1, str, "<") Then
            startPos = InStr(startPos + 1, str, "</")
            endPos = InStr(startPos + 1, str, "<")
            If endPos = 0 Then
                endPos = Len(str) + 1
            End If
            indentationLevel = indentationLevel - 1
            new_str = new_str & vbCrLf & String(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)
        Else
            startPos = InStr(startPos + 1, str, "<")
            endPos = InStr(startPos + 1, str, "<")
            If endPos = 0 Then
                endPos = Len(str) + 1
            End If
            new_str = new_str & vbCrLf & String(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)
            Dim tagName As String
            tagName = LCase(returnNameOfTag(returnNextTag(str, startPos)))
            If tagName <> "br" And tagName <> "hr" And tagName <> "img" And tagName <> "meta" And tagName <> "applet" And tagName <> "p" And tagName <> "!--" And tagName <> "input" And tagName <> "!doctype" And tagName <> "area" Then
            'If isPairedTag(tagName) Then
                indentationLevel = indentationLevel + 1
            End If
        End If
    Loop While startPos > 0
    Eformat = new_str
End Function

Public Function returnNameOfTag(ByRef str As Tag) As String
    On Error Resume Next
    
    Dim endPos As Double
    Dim Start As Double
    
    Start = 2
    endPos = InStr(1, str.Text, " ")
    If Mid$(str.Text, 2, 3) = "!--" Then
        endPos = 5
    ElseIf endPos = 0 Then
        endPos = InStr(1, str.Text, ">")
    End If

    returnNameOfTag = Mid$(str.Text, Start, endPos - Start)
End Function

Public Function returnNextTag(ByRef str As String, ByVal Start As Double) As Tag
    On Error Resume Next

    Dim endPos As Double

    Start = InStr(Start + 1, str, "<")
    endPos = InStr(Start + 1, str, ">")
    returnNextTag.Text = Mid$(str, Start, endPos - Start + 1)
    returnNextTag.Start = Start
    returnNextTag.length = endPos - Start
End Function

' Function for getting parts of a filename
' Handles local filenames ONLY!
'
Public Function GetFilePart(Filename$, FilePart As kFilePart) As String
Static LastF$
Static Dr$, P$, E$, f$

Dim StartP%, LastS%, EndP%
Dim StartE%, EndE%

    If Filename <> LastF Then
            
        P$ = ""
        Dr$ = ""
        E$ = ""
        
        If Filename Like "[a-zA-Z]:\*" Then
            Dr = Left$(Filename, 2)
        End If
        
        If Filename Like "*\*" Then
            If Len(Dr$) > 0 Then
                StartP = 3
                LastS = 1
            Else
                StartP = 1
                LastS = 1
            End If
            
            Do
                LastS = InStr(LastS, Filename, "\")
                If LastS = 0 Then
                    P$ = Mid$(Filename, StartP, (EndP - StartP) + 1)
                Else
                    EndP = LastS
                    LastS = LastS + 1
                End If
            Loop Until Len(P$) > 0
        End If
        
        If Filename Like "*.*" Then
            LastS = 1
            Do
                LastS = InStr(LastS, Filename, ".")
                If LastS = 0 Then
                    E$ = Right$(Filename, Len(Filename) - EndE)
                Else
                    EndE = LastS
                    LastS = LastS + 1
                End If
            Loop Until Len(E$) > 0
        End If
        
        If Len(P) = 0 And Len(Dr) = 0 Then
            f$ = Filename
        Else
            f$ = Right$(Filename, Len(Filename) - EndP)
        End If
    End If
    
    Select Case FilePart
        Case kfpDrive
            GetFilePart = Dr$
        Case kfpDrivePath
            GetFilePart = Dr$ & P$
        Case kfpPath
            GetFilePart = P$
        Case kfpPathFile
            GetFilePart = P$ & f$
        Case kfpExtension
            GetFilePart = E$
        Case kfpFilename
            GetFilePart = f$
    End Select
    LastF = Filename
End Function

' Extracts the RGB values from a VBA Long integer
' representing the combined values
'
Public Sub ExtractRGB(ColorVal, R, G, B)
    R = ColorVal And &HFF
    G = (ColorVal \ &H100) And &HFF
    B = (ColorVal \ &H10000) And &HFF
End Sub

'Here is a small function to change button to flat:-
Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True
'    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
'    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

'/////////////////////////////////////////////////////////////////////////
'// Disables the X
Public Sub RemoveMenus(f As Form)
    Dim lMenu As Long
    '// Get the form's system menu handle.
    lMenu = GetSystemMenu(f.hwnd, False)
    DeleteMenu lMenu, 6, MF_BYPOSITION
    DeleteMenu lMenu, 5, MF_BYPOSITION
End Sub
Public Sub FlattenCombo(ctl As ComboBox, f As Form)
    Dim ctlx As Control

    For Each ctlx In f.Controls
        If TypeOf ctl Is ComboBox Then
            m_iCount = m_iCount + 1
            ReDim Preserve m_cFlatten(1 To m_iCount) As cFlatControl
            Set m_cFlatten(m_iCount) = New cFlatControl
            m_cFlatten(m_iCount).Attach ctl
        End If
    Next ctlx
End Sub

Public Function ShowProperties(Filename As String, OwnerhWnd As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    Dim R As Long
    
    'Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    'call the API
    R = ShellExecuteEX(SEI)
    'return the instance handle as a sign of success
    ShowProperties = SEI.hInstApp
End Function

Public Sub ToDoIndent()
' AUTO INDENT
'   This will implement the auto-indent feature
        
    ' Variables to perform auto-indent
    Dim i As Long
    Dim addString As String
    Dim message As String
        
    ' Get the previous line, and save it in message
    CurrentLine = SendMessage(frmFileView.txtFile.hwnd, EM_LINEFROMCHAR, -1, 0)
    getLine frmFileView.txtFile.hwnd, CurrentLine, message

    ' Initialize addString
    addString = vbCrLf
    ' Walk along the message and count the spaces / tabs
    For i = 1 To Len(message)
        ' Do the spaces
        If Mid(message, i, 1) = Chr(32) Then
            addString = addString + " "
        ' Do the tabs
        ElseIf Mid(message, i, 1) = vbTab Then
            addString = addString + vbTab
        Else
            ' Hit something else, so exit the for loop and
            ' add the indented text
            Exit For
        End If
    Next i
    ' Insert text the text
    frmFileView.txtFile.SelText = addString
End Sub

Public Sub getLine(ByVal hwnd As Long, ByVal whichLine As Integer, Line As String)
' GET LINE
'   Used to get a line from a text box
    ' Declare variables
    Dim length As Long, bArr() As Byte, bArr2() As Byte, lc As Long
    ' Find the start of the line I want
    lc = SendMessage(hwnd, EM_LINEINDEX, whichLine, ByVal 0&)
    ' Get the length of the line I want
    length = SendMessage(hwnd, EM_LINELENGTH, lc, ByVal 0&)
    ' If line exists
    If length > 0 Then
        ' Get the storage space
        ReDim bArr(length + 1) As Byte, bArr2(length - 1) As Byte
        ' Get and copy the line (bytes)
        Call RtlMoveMemory(bArr(0), length, 2)
        Call SendMessage(hwnd, EM_GETLINE, whichLine, bArr(0))
        Call RtlMoveMemory(bArr2(0), bArr(0), length)
        ' Convert the bytes to a string
        Line = StrConv(bArr2, vbUnicode)
    Else
        ' Line does not exist, set it equal to nothing
        Line = ""
    End If
End Sub


Public Function ShellEx(ByVal sFile As String, Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, Optional ByVal sParameters As String = "", Optional ByVal sDefaultDir As String = "", Optional sOperation As String = "open", Optional Owner As Long = 0) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function
