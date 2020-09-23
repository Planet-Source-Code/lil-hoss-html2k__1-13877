Attribute VB_Name = "modHTMLColoring"
'Option Explicit
'

'
'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type
'
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
'Public Const EM_CHARFROMPOS& = &HD7
'Public Const EM_GETFIRSTVISIBLELINE = &HCE '(0&,pt)
'Public Const EM_FMTLINES = &HC8
'Public Const EM_GETLINE = &HC4 '(line num,pt)=len
'Public Const EM_GETLINECOUNT = &HBA
'Public Const EM_LINEINDEX = &HBB '(char num,0&)


'Public subs.
'Public Function GetFirstLinePos(Line As Integer, Start As Integer, rtf As RichTextBox) As Integer
'Dim i As Integer, C As Integer
'
'For i = Start To 0 Step -1
'    C = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, _
'    i, 0&)
'
'    If C < Line Then
'        GetFirstLinePos = i + 1
'
'        Exit For
'    ElseIf i = 0 Then
'        GetFirstLinePos = 0
'    End If
'Next i
'
'End Function
'
'Public Function GetLastLinePos(Line As Integer, Start As Integer, rtf As RichTextBox) As Integer
'Dim i As Integer, C As Integer
'
'For i = Start To Len(rtf.Text)
'    C = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, _
'    i, 0&)
'
'    If C > Line Then
'        GetLastLinePos = i - 1
'
'        Exit For
'    ElseIf i = Len(rtf.Text) Then
'        GetLastLinePos = Len(rtf.Text)
'    End If
'Next i
'
'End Function

'Public Sub ColorTags(iStart As Integer, iEnd As Integer, rtf As RichTextBox, Optional Color As Long = vbBlue, Optional ErrorColor As Long = vbRed, Optional CommentColor As Long = &H8000&, Optional ParamColor As Long = &H800080, Optional IncludeColor As Long = &H40C0&)
'Dim iFirst As Integer
'Dim iLast As Integer
'Dim tmp As Long
'Dim I As Integer, C As Integer, t As Integer
'Dim OldStart As Integer
'On Error Resume Next
''Turn refreshing off.
'tmp = LockWindowUpdate(rtf.hwnd)
'
'OldStart = iStart
'
'rtf.SelStart = iStart
'rtf.SelLength = iEnd - iStart
'rtf.SelColor = vbBlack
'rtf.SelLength = 0
'
'iStart = InStr(iStart + 1, rtf.Text, "<")
'
'If iStart > 0 Then _
'    rtf.SelStart = iStart - 1
'
'iFirst = iStart - 1
''iFirst = InStr(iFirst + 1, rtf.Text, "<") - 1
'
'C = OldStart + 1
'I = InStr(OldStart + 1, rtf.Text, "<")
'
'Do
'    C = InStr(C + 1, rtf.Text, ">")
'
'    If C < I Then
'        rtf.SelStart = C - 1
'        rtf.SelLength = 1
'
'        rtf.SelColor = ErrorColor
'
'        If iStart > 0 Then
'            rtf.SelStart = iStart - 1
'        Else
'            rtf.SelStart = 0
'        End If
'        rtf.SelLength = 0
'
'    Else
'        Exit Do
'    End If
'Loop
'
'
'iLast = InStr(iFirst + 1, rtf.Text, ">")
'
'I = iLast
'C = iLast
'Do
'    'A ">" without "<".
'    I = InStr(I + 1, rtf.Text, "<")
'    C = InStr(C + 1, rtf.Text, ">")
'
'    If (C < I And C > 0) And (C < iEnd) Or _
'    (C > 0 And I = 0) And (C < iEnd) Then
'        rtf.SelStart = C - 1
'        rtf.SelLength = 1
'        rtf.SelColor = ErrorColor
'        rtf.SelLength = 0
'
'        rtf.SelStart = iFirst
'    End If
'
'    If I = 0 Then
'        I = iLast
'    End If
'Loop Until (C = 0)
'
'
'I = 0
'C = 0
'
'Do Until iFirst = -1
'    iLast = InStr(iFirst + 1, rtf.Text, ">")
'    rtf.SelStart = iFirst
'
'    'A "<" without ">"
'    tmp = InStr(iFirst + 2, rtf.Text, "<")
'
'    If tmp < iLast And tmp > 0 Or iLast = 0 Then
'        If tmp = 0 Then
'            rtf.SelLength = Len(rtf.Text)
'        Else
'            rtf.SelLength = tmp - iFirst - 1
'        End If
'
'        rtf.SelColor = ErrorColor
'    Else
'        rtf.SelLength = iLast - iFirst
'
'        If Mid(rtf.Text, iFirst + 1, 4) = "<!--" And _
'        Mid(rtf.Text, iLast - 2, 3) = "-->" Then
'
'            tmp = InStr(iFirst, rtf.Text, "#include", vbTextCompare)
'
'            If tmp > 0 Then
'                If Trim(Mid(rtf.Text, iFirst + 5, tmp - iFirst - 5)) = "" Then
'                    rtf.SelColor = IncludeColor
'                Else
'                    rtf.SelColor = CommentColor
'                End If
'            Else
'                rtf.SelColor = CommentColor
'            End If
'        Else
'            rtf.SelColor = Color
'        End If
'
'        'Color the parameters.
'        t = iFirst
'
'        Do
'            tmp = InStr(t + 1, rtf.Text, "=")
'
'            If tmp > 0 And tmp < iLast Then
'                For I = tmp + 1 To iLast
'                    If Mid(rtf.Text, I, 1) <> " " And _
'                    Mid(rtf.Text, I, 1) <> vbCr And _
'                    Mid(rtf.Text, I, 1) <> vbLf Then
'                        Exit For
'                    End If
'                Next I
'
'                If I >= iLast Then
'                    'A '=' without a parameter.
'                    rtf.SelStart = tmp - 1
'                    rtf.SelLength = 1
'                    rtf.SelColor = ErrorColor
'                    Exit Do
'                End If
'
'                For C = I + 1 To iLast
'                    If Mid(rtf.Text, C, 1) = """" And _
'                    Mid(rtf.Text, I, 1) = """" Then
'                        Exit For
'                    ElseIf Mid(rtf.Text, C, 1) = " " And _
'                    Mid(rtf.Text, I, 1) <> """" Or _
'                    Mid(rtf.Text, C, 1) = vbCr And _
'                    Mid(rtf.Text, I, 1) <> """" Or _
'                    Mid(rtf.Text, C, 1) = vbLf And _
'                    Mid(rtf.Text, I, 1) <> """" Then
'                        Exit For
'                    End If
'                Next C
'
'                If C >= iLast And _
'                Mid(rtf.Text, I, 1) = """" Then
'                    'A parameter starting with
'                    ''"' and doesn't end with one.
'
'                    rtf.SelStart = I - 1
'                    rtf.SelLength = iLast - I
'                    rtf.SelColor = ErrorColor
'
'                    Exit Do
'                End If
'
'                'Color the parameter.
'                rtf.SelStart = I - 1
'                rtf.SelLength = C - I + 1
'
'                If rtf.SelColor = CommentColor Then
'                    Exit Do
'                End If
'
'                rtf.SelColor = ParamColor
'
'                t = tmp + 1
'            Else
'                Exit Do
'            End If
'        Loop
'    End If
'
'    iFirst = rtf.Find("<", iFirst + 1, , rtfNoHighlight)
'
'    If iFirst > iEnd Then Exit Do
'Loop
'rtf.SelStart = iStart
'
'
''Allow repainting (Refreshing).
'tmp = LockWindowUpdate(0)
'
'rtf.SelStart = OldStart
'End Sub


'Public Function InStrRev(Start As Integer, SearchIn As String, SearchFor As String, Optional MatchCase As Boolean = False) As Integer
'Dim i As Integer
'
'For i = Start To 1 Step -1
'    If (UCase(Mid(SearchIn, i, _
'    Len(SearchFor))) = UCase(SearchFor)) And _
'    MatchCase = False Then
'        'Found match.
'        InStrRev = i
'
'        Exit For
'    ElseIf (Mid(SearchIn, i, _
'    Len(SearchFor)) = SearchFor) And _
'    MatchCase = True Then
'        'Found match.
'        InStrRev = i
'
'        Exit For
'    End If
'Next i
'
'End Function


