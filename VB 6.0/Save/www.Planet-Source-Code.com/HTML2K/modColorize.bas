Attribute VB_Name = "modColorize"
'    --------------------------------------------------------------------------
'    EzColorTest HTML Editor Color Coding Test
'    Copyright (C) 2000  Eric Banker
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'
'    Note: this version is a bit faster but is still really slow. I'm going to
'    be moving this code over to inserting RTF codes shortly which will speed it
'    up greatly.
'    --------------------------------------------------------------------------

Option Explicit

' These hold the color information
Public m_TextCol As String
Public m_AttribCol As String
Public m_TagCol As String
Public m_CommentCol As String
Public m_AspCol As String
Public m_TableCol As String
Public m_ColdCol As String

' These are for the color coding stuff
Public cInTag As Boolean
Public cInComment As Boolean
Public cInASP As Boolean
Public cInCold As Boolean
Public cInTable As Boolean
Public cTypedIn As Boolean
Public cInAttribQuote As Boolean, cInAttrib As Boolean

' These are special for cold fusion keypress color coding
Public cPrev As String
Public cBefPrev As String

' This stops the program from displaying the line and column while color coding
Public StopLine As Boolean

' ##########################################################################################
' These below are the color coding functions. These handle all color coding for the program.
' ##########################################################################################

' Call this when you load a form with code in it. It color codes the entire document

Public Sub HtmlHighlight()
On Error Resume Next
    frmDoc.trapUndo = False
    StopLine = True
    StopLines = True
    
    ' Color Html and asp
    HtmlColorCode
    
    ' Move back to the start of the thing
    frmDoc.rtfHTML.SelStart = 0
    
    StopLines = False
    StopLine = False
    frmDoc.trapUndo = True
End Sub

' Colorizes HTML while typing
' --------------------------------------------------------------

Public Function KeyPressEvent(KeyAscii As Integer) As Integer
    frmDoc.trapUndo = False
      
    Dim cChar As String
    
        cChar = Chr$(KeyAscii)
        
        If cInTag = False And cInAttrib = False And cInComment = False And cInASP = False And cInCold = False Then
            frmDoc.rtfHTML.SelColor = m_TextCol
        End If
        
        If cChar = "<" And (cInASP = False And cInComment = False And cInCold = False) Then
            frmDoc.rtfHTML.SelColor = m_TagCol
            cInTag = True
        End If
        
        If cInTag = True And (cInAttrib = True Or cInAttribQuote = True) Then
            frmDoc.rtfHTML.SelColor = m_AttribCol
        End If

        If cChar = "=" And cInTag = True Then
            cInAttrib = True
        End If

        If cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = True Then
            cInAttrib = False
            cInAttribQuote = False
        ElseIf cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = False Then
            cInAttribQuote = True
        End If

        If cChar = Chr$(32) And (cInAttribQuote = False And cInTag = True) Then
            frmDoc.rtfHTML.SelColor = m_TagCol
            cInAttrib = False
        End If

        If cChar = "!" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart, 1) = "<" Then

            frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 1
            frmDoc.rtfHTML.SelLength = 1
            frmDoc.rtfHTML.SelColor = m_CommentCol
            frmDoc.rtfHTML.SelText = "<!--"

            cInTag = False
            cInAttrib = False
            cInASP = False
            cInComment = True

            KeyAscii = 0
        End If
        
        If cChar = "%" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart, 1) = "<" Then

            frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 1
            frmDoc.rtfHTML.SelLength = 1
            frmDoc.rtfHTML.SelColor = m_AspCol
            frmDoc.rtfHTML.SelText = "<%"

            cInTag = False
            cInAttrib = False
            cInASP = True
            cInComment = False

            KeyAscii = 0
        End If
        
        If cPrev = "c" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart - 1, 1) = "<" Then
            If cChar = "f" Then
                frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 2
                frmDoc.rtfHTML.SelLength = 2
                frmDoc.rtfHTML.SelColor = m_AspCol
                frmDoc.rtfHTML.SelText = "<cf"
            
                cInTag = False
                cInAttrib = False
                cInASP = False
                cInCold = True
                cInComment = False

                KeyAscii = 0
            Else
                ' do nothing we are not in a cf tag so color normal
            End If
        End If
        
        If cBefPrev = "/" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart - 2, 1) = "<" Then
            If cPrev = "c" Then
                If cChar = "f" Then
                    frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 3
                    frmDoc.rtfHTML.SelLength = 3
                    frmDoc.rtfHTML.SelColor = m_AspCol
                    frmDoc.rtfHTML.SelText = "</cf"
            
                    cInTag = False
                    cInAttrib = False
                    cInASP = False
                    cInCold = True
                    cInComment = False

                    KeyAscii = 0
                End If
            End If
        End If
        
        If cPrev = "C" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart - 1, 1) = "<" Then
            If cChar = "F" Then
                frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 2
                frmDoc.rtfHTML.SelLength = 2
                frmDoc.rtfHTML.SelColor = m_AspCol
                frmDoc.rtfHTML.SelText = "<CF"
            
                cInTag = False
                cInAttrib = False
                cInASP = False
                cInCold = True
                cInComment = False

                KeyAscii = 0
            Else
                ' do nothing we are not in a cf tag so color normal
            End If
        End If
        
        If cBefPrev = "/" And Mid$(frmDoc.rtfHTML.Text, frmDoc.rtfHTML.SelStart - 2, 1) = "<" Then
            If cPrev = "C" Then
                If cChar = "F" Then
                    frmDoc.rtfHTML.SelStart = frmDoc.rtfHTML.SelStart - 3
                    frmDoc.rtfHTML.SelLength = 3
                    frmDoc.rtfHTML.SelColor = m_AspCol
                    frmDoc.rtfHTML.SelText = "</CF"
            
                    cInTag = False
                    cInAttrib = False
                    cInASP = False
                    cInCold = True
                    cInComment = False

                    KeyAscii = 0
                End If
            End If
        End If
        
        If cChar = ">" Then
            If cInComment = False And cInASP = False And cInCold = True Then
                frmDoc.rtfHTML.SelColor = m_AspCol
            ElseIf cInComment = False And cInASP = True And cInCold = False Then
                frmDoc.rtfHTML.SelColor = m_AspCol
            ElseIf cInComment = True And cInASP = False And cInCold = False Then
                frmDoc.rtfHTML.SelColor = m_CommentCol
            ElseIf cInComment = False And cInASP = False And cInCold = False Then
                frmDoc.rtfHTML.SelColor = m_TagCol
            End If
            
            cInTag = False
            cInASP = False
            cInCold = False
            cInAttrib = False
            cInComment = False
        End If

    KeyPressEvent = KeyAscii
    
    ' This keeps track of the previous 2 keys for color coding
    ' different types of tags
    cBefPrev = cPrev
    cPrev = Chr$(KeyAscii)
    
    ' make sure we re-enable undo/redo
    frmDoc.trapUndo = True
    
    Exit Function
    
ErrExit:
    Exit Function
End Function

' Insert text w/tag coloring if necessary

Public Sub InsertTag(Tag$, StopAsp As Boolean)
Dim S As Long
    S = frmDoc.rtfHTML.SelStart
    
    frmDoc.trapUndo = False
    If Len(frmDoc.rtfHTML.SelText) > 0 Then frmDoc.rtfHTML.SelText = ""
    frmDoc.trapUndo = True
    
    frmDoc.rtfHTML.SelText = Tag$
    
    If StopAsp = True Then
        frmDoc.trapUndo = False
        HtmlColorCode S, S + Len(Tag), True
        frmDoc.trapUndo = True
    Else
        frmDoc.trapUndo = False
        HtmlColorCode S, S + Len(Tag), False
        frmDoc.trapUndo = True
    End If
End Sub

' Insert Asp code with asp coloring, This is no longer needed as I have made ASP color coding
' and all others in one function

Public Sub InsertAspTag(Tag$)
Dim U As Long
    U = frmDoc.rtfHTML.SelStart
    If Len(frmDoc.rtfHTML.SelText) > 0 Then frmDoc.rtfHTML.SelText = ""
    frmDoc.rtfHTML.SelText = Tag$

    frmDoc.trapUndo = False
    ASPColorCode U, U + Len(Tag)
    frmDoc.trapUndo = True
End Sub

' This function determines whether the caret is currently outside a tag.

Public Function IsOutsideTag()
On Error Resume Next

' These give me the location of start and
' end tags for all color code items
Dim LastGT As Long, LastLT As Long
Dim EndTag As Long, StartTag As Long
Dim EndASP As Long, StartASP As Long
Dim LastASP As Long, LastASPT As Long
Dim EndComment As Long, StartComment As Long
Dim LastComment As Long, LastCommentT As Long

' Here are some text variables
Dim txt$, Start As Long, Start2 As Long

' These are the variables that tell me where i am at
' and which item to color code
Dim InMainTag As Boolean, InEndTag As Boolean
Dim InASPTag As Boolean, InEndASP As Boolean
Dim InCommentTag As Boolean, InEndComment As Boolean
    
    ' Get back the current text
    txt = frmDoc.rtfHTML.Text
    Start = frmDoc.rtfHTML.SelStart
    
    ' This checks to see if it's an asp tag
    EndASP = InStr(Start + 1, txt, "%>", vbBinaryCompare)
    StartASP = InStr(Start + 1, txt, "<%", vbBinaryCompare)
    
    If StartASP > EndASP Then
        InASPTag = True
    Else
        InASPTag = False
    End If
    
    LastASP = RevInStr(txt, "<%", Start + 1, vbBinaryCompare)
    LastASPT = RevInStr(txt, "%>", Start + 1, vbBinaryCompare)
        
    If LastASP < LastASPT Then
        InEndASP = True
    Else
        InEndASP = False
    End If
    
    ' This checks to see if it's a comment
    EndComment = InStr(Start + 1, txt, "->", vbBinaryCompare)
    StartComment = InStr(Start + 1, txt, "<!-", vbBinaryCompare)
    
    If StartComment > EndComment Then
        InCommentTag = True
    Else
        InCommentTag = False
    End If
    
    LastComment = RevInStr(txt, "<!--", Start + 1, vbBinaryCompare)
    LastCommentT = RevInStr(txt, "->", Start + 1, vbBinaryCompare)
        
    If LastComment < LastCommentT Then
        InEndComment = True
    Else
        InEndComment = False
    End If
    
    ' This checks to see if it's an html attribute
    EndTag = InStr(Start + 1, txt, ">", vbBinaryCompare)
    StartTag = InStr(Start + 1, txt, "<", vbBinaryCompare)
       
    If StartTag > EndTag Then
        InMainTag = True
    Else
        InMainTag = False
    End If
        
    LastLT = RevInStr(txt, "<", Start + 1, vbBinaryCompare)
    LastGT = RevInStr(txt, ">", Start + 1, vbBinaryCompare)
        
    If LastLT < LastGT Then
        InEndTag = True
    Else
        InEndTag = False
    End If
    
    ' This code takes the info above and sets the flags right for
    ' color coding
        
    If InASPTag = True Or InEndASP = True Then
        cTypedIn = True
        cInASP = True
        
    ' This is if Comment is true then color code comment
    ElseIf InCommentTag = True Or InEndComment = True Then
        cTypedIn = True
        cInComment = True
        
    ' This is if Html is true then color code HTML tags
    ElseIf InMainTag = True Or InEndTag = True Then
        cTypedIn = True
        cInTag = True
        
    ' Nothing is true so don't color code anything
    Else
        cTypedIn = False
        cInAttrib = False
        cInAttribQuote = False
        cInCold = False
        cInASP = False
        cInComment = False
        cInTag = False
    End If
End Function

' ##########################################################################################
' These are the main color coding functions. These are not called ever by the user.
' ##########################################################################################

' This is the main color coding function. This does everything html, comments, and attributes. It also calls
' the ASP color coding function if nessasary

Public Function HtmlColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1, Optional StopAsp As Boolean = False)
On Error GoTo ErrHandler
    ' These are the variables for the tags for ColorCoding
    Dim CommentOpenTag As String
    Dim CommentCloseTag As String

    Dim oldselstart As Long, oldsellen As Long
    
    ' These are place holders for the color coding
    Dim tag_open As Long
    Dim tag_close As Long
    Dim Curr As String
    
    With frmDoc.rtfHTML
    
    Dim strTextMain As String
    strTextMain = .Text
    
    frmDoc.trapUndo = False
    
    ' Find out where the cursor is
    oldselstart = .SelStart
    oldsellen = .SelLength
    
    If endchar = -1 Then endchar = Len(strTextMain)
    If startchar = 0 Then startchar = 1

    ' These are the close tags for colorcoding
    tag_close = startchar
    
    ' Now lets loop through the tags and color code it
    Do
        ' See where the next tag starts. if any
        tag_open = InStr(tag_close, strTextMain, "<", vbBinaryCompare)
        
        'If so, then color it...
        If tag_open <> 0 Then  'Found a tag
            
            'Find the end of the next tag...
            tag_close = InStr(tag_open, strTextMain, ">", vbBinaryCompare)

            'Get the current HTML tag...
            Curr = Mid$(strTextMain, tag_open, tag_close - tag_open + 1)
            
            If tag_close <> 0 Then
                ' Comments
                If Left$(Curr, 3) = "<!-" Then
                    tag_close = InStr(tag_open, strTextMain, "->", vbBinaryCompare) + 1
                    .SelStart = tag_open - 1
                    .SelLength = tag_close - tag_open + 1
                    .SelColor = m_CommentCol
                ElseIf Left$(Curr, 1) = "<" Then
                    cycleAttrib Curr, tag_open, tag_close
                End If
            End If
            
            If tag_close = 0 Or tag_close >= endchar Then
                ' If we are coloring tags and it's over the end tag then
                ' get me out of this loop and don't color anymore
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    ' Color ASP Stuff only if we need to. We have a special function for coloring ASP tags so we won't
    ' worry if this deals with it or not.
    If StopAsp = False Then
        If Right(OpenFilename, 3) = "asp" Or Right(OpenFilename, 3) = "asa" Or OpenFilename = "" Then
            ASPColorCode startchar, endchar
        Else
            ' don't color code
        End If
    End If
    
    If Right(OpenFilename, 3) = "cfm" Or OpenFilename = "" Then
        CFColorCode startchar, endchar
    Else
        ' Don't color code
    End If
    
    .SelStart = oldselstart
    .SelLength = oldsellen
    .SetFocus
    
    frmDoc.trapUndo = True
    
    End With
    Exit Function
    
ErrHandler:
    Exit Function
End Function

' This function colorizes ASP code but is no longer needed

Private Function ASPColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1)
On Error GoTo ErrHandler
    Dim oldselstart As Long, oldsellen As Long

    ' These are place holders for the color coding
    Dim tag_open As Long
    Dim tag_close As Long
    Dim Curr As String

    With frmDoc.rtfHTML

    Dim strText As String
    strText = .Text

    ' don't allow undo to see the color changes
    frmDoc.trapUndo = False

    ' Find out where the cursor is
    oldselstart = .SelStart
    oldsellen = .SelLength

    If endchar = -1 Then endchar = Len(strText)
    If startchar = 0 Then startchar = 1
    
    ' These are the close tags for colorcoding
    tag_close = startchar

    ' Now lets loop through the tags and color code it
    Do
        ' See where the next tag starts. if any
        tag_open = InStr(tag_close, strText, "<%", vbBinaryCompare)

        'If so, then color it...
        If tag_open <> 0 Then  'Found a tag

            'Find the end of the next tag...
            tag_close = InStr(tag_open, strText, "%>", vbBinaryCompare)

            'Get the current HTML tag...
            Curr = Mid$(strText, tag_open, tag_close - tag_open + 1)

            If tag_close <> 0 Then
                Select Case Left$(Curr, 2)
                    Case "<%"
                        ' It's asp
                        tag_close = InStr(tag_open, strText, "%>", vbBinaryCompare) + 1
                        .SelStart = tag_open - 1
                        .SelLength = tag_close - tag_open + 1
                        .SelColor = m_AspCol
                    Case Else
                        ' it's not an asp tag so do nothing
                End Select
            End If

            If tag_close = 0 Or tag_close >= endchar Then
                ' If we are coloring tags and it's over the end tag then
                ' get me out of this loop and don't color anymore
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    ' reset the cursor position
    .SelStart = oldselstart
    .SelLength = oldsellen
    .SetFocus

    ' reinit the undo stuff
    frmDoc.trapUndo = True

    End With
    Exit Function

ErrHandler:
    Exit Function
End Function

' This function colorizes Cold Fusion code but is no longer needed

Private Function CFColorCode(Optional startchar As Long = 1, Optional endchar As Long = -1)
On Error GoTo ErrHandler
    Dim oldselstart As Long, oldsellen As Long

    ' These are place holders for the color coding
    Dim tag_open As Long
    Dim tag_close As Long
    Dim Curr As String

    With frmDoc.rtfHTML

    Dim strText As String
    strText = .Text

    ' don't allow undo to see the color changes
    frmDoc.trapUndo = False

    ' Find out where the cursor is
    oldselstart = .SelStart
    oldsellen = .SelLength

    If endchar = -1 Then endchar = Len(strText)
    If startchar = 0 Then startchar = 1
    
    ' These are the close tags for colorcoding
    tag_close = startchar

    ' Now lets loop through the tags and color code it
    Do
        ' See where the next tag starts. if any
        tag_open = InStr(tag_close, strText, "<", vbBinaryCompare)

        'If so, then color it...
        If tag_open <> 0 Then  'Found a tag

            'Find the end of the next tag...
            tag_close = InStr(tag_open, strText, ">", vbBinaryCompare)

            'Get the current HTML tag...
            Curr = Mid$(strText, tag_open, tag_close - tag_open + 1)

            If tag_close <> 0 Then
                If Left$(Curr, 3) Like "<CF" Or Left$(Curr, 3) Like "<cf" Then
                    ' It's asp
                    tag_close = InStr(tag_open, strText, ">", vbBinaryCompare) + 1
                    .SelStart = tag_open - 1
                    .SelLength = tag_close - tag_open
                    .SelColor = m_AspCol
                End If
                If Left$(Curr, 4) Like "</CF" Or Left$(Curr, 4) Like "</cf" Then
                        tag_close = InStr(tag_open, strText, ">", vbBinaryCompare) + 1
                        .SelStart = tag_open - 1
                        .SelLength = tag_close - tag_open + 1
                        .SelColor = m_AspCol
                End If
            End If

            If tag_close = 0 Or tag_close >= endchar Then
                ' If we are coloring tags and it's over the end tag then
                ' get me out of this loop and don't color anymore
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    ' reset the cursor position
    .SelStart = oldselstart
    .SelLength = oldsellen
    .SetFocus

    ' reinit the undo stuff
    frmDoc.trapUndo = True

    End With
    Exit Function

ErrHandler:
    Exit Function
End Function

' This cycles through the html and comes back with the right tag colors for the tag and all of it's
' attributes I am not using it right now because it's really slow. There are various places in the
' program where there are comments set to uncomment or comment to enable attribute color coding.
' Please look that over if you wish to enable it

Private Function cycleAttrib(CurrTag As String, opentag As Long, closetag As Long)

    Dim fPos As Long, sPos As Long, qPos As Long, qnPos As Long, aPos As Long, tBeg As Long, tEnd As Long
    Dim isFirstCycle As Boolean
    Dim eTag As String
    Dim sPosTxt As String
    Dim LeftOver As Long
    Dim EndTag As Long, QuotePos As Long, QuoteEndPos As Long

    With frmDoc.rtfHTML

    eTag = CurrTag
    isFirstCycle = True

    Do While Len(eTag) > 0
        fPos = InStr(1, eTag, "=")

        If (fPos = 0 And isFirstCycle = True) Then
            ' This just checks to see if it's a basic html tag w/ no attributes and if so colors that
            ' without going through the rest of the junk.
            .SelStart = opentag - 1
            .SelLength = closetag - opentag + 1
            .SelColor = m_TagCol
            Exit Function
        ' It looks like we have an attribute. Here comes the hard part...
        ElseIf fPos <> 0 Then 'Put in the color info...
            If Left$(eTag, 1) = "<" Then
                ' This brings back the entire tag. something like:
                ' <img src="blah.jpg" onclick="blah">
                ' and then color codes the entire thing
                tBeg = opentag
                tEnd = opentag + fPos

                ' Color Code the entire tag first
                .SelStart = tBeg - 1
                .SelLength = closetag - tBeg + 1
                .SelColor = m_TagCol

                ' This brings back the text that is past the attribute. in the previous example:
                ' "blah.jpg" onclick="blah">
                eTag = Mid$(eTag, fPos + 1)
                LeftOver = closetag - Len(eTag)
            End If
        End If

        'Find the first instance of a space in the
        'part of the tag that we have left...
        sPos = InStr(1, eTag, Chr$(32), vbBinaryCompare)

        'Gets the text up to the next space...
        sPosTxt = Mid$(eTag, 1, sPos)

        'Checks to see if there's a quote in the text...
        qPos = InStr(1, sPosTxt, Chr$(34), vbBinaryCompare)

        'If there's a quote found, then we need to find
        'its end...
        If qPos <> 0 Then
            'Look for the next quote...
            qnPos = InStr(2, eTag, Chr$(34), vbBinaryCompare)

            If qnPos <> 0 Then
                sPosTxt = Mid$(eTag, 1, qnPos)
            End If
        End If

        LeftOver = closetag - Len(eTag)
        .SelStart = LeftOver
        .SelLength = Len(sPosTxt)
        .SelColor = m_AttribCol

        'Truncates the tag so there's no attrib value left...
        eTag = Mid$(eTag, Len(sPosTxt) + 1)

        'Find the next position of an equal sign...
        sPos = InStr(1, eTag, "=")

        'If there's no =, then we know we're on the last
        'attrib value, so we need to put in some final
        'info...all that's left is something like:
        '"#ffffff">
        If sPos = 0 Then
            'Put in the attrib color before the ">"
            'if it's the last attribute...
            eTag = Mid$(eTag, 1, Len(eTag) - 1)

            'Insert the RTF info...
            'bef = bef & infoRTF & AttribInfo & eTag
            .SelStart = LeftOver
            .SelLength = Len(eTag)
            .SelColor = m_AttribCol

            'Truncate the end...
            sPos = Len(eTag)
            Exit Do
        End If

        'Truncates the tag appropriately...
        eTag = Mid$(eTag, sPos + 1)
        isFirstCycle = False

        'If there's nothing left, then we need to exit
        'the loop so it doesn't loop infinitely...
        If sPos = 0 And qPos = 0 Then Exit Do
    Loop

    End With
    Exit Function
End Function
