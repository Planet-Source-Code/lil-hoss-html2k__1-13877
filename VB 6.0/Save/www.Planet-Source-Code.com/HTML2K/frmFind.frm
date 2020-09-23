VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   0
      Top             =   960
   End
   Begin VB.ComboBox cboReplaceStr 
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox cboSearchStr 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
      Begin VB.CheckBox chkMatchWord 
         Caption         =   "&Whole Word Only"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "&Match Case"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraScope 
      Caption         =   "Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
      Begin VB.OptionButton optSelected 
         Caption         =   "&Selected"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optWholeText 
         Caption         =   "Al&l Text"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2145
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Replace:"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "frmFind"
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

' Search options
Private MatchCase As Boolean, MatchWord As Boolean

' For autocomplete of combo boxes, not saved across
' sessions
Private SearchHis As Collection, ReplaceHis As Collection

' Indicates whether the form is in Find-only state or
' Find/Replace
'
Public Replace As Boolean
Private Locked As Boolean
Private Const WM_SYSCOMMAND = &H112
Private Const MOUSE_MOVE = &HF012

' These are positional constants for shuffling things
' around when user switches to Replace mode
'
Const REPLACE_HEIGHT = 3120
Const REPLACE_OPT_TOP = 1275
Const REPLACE_CLOSE_TOP = 1800

Const FIND_HEIGHT = 2790
Const FIND_OPT_TOP = 840
Const FIND_CLOSE_TOP = 1320

' Subclassed text box allows us to listen in on events
Private WithEvents rtbText As RichTextBox
Attribute rtbText.VB_VarHelpID = -1

Private Wrapped As Boolean
Private LastLine As Long


Private Sub rtbText_SelChange()
    If Locked Then Exit Sub
    
    ' Update scope options
    If rtbText.SelLength > 0 Then
        optSelected.Enabled = True
    Else
        optWholeText.Value = True
        optSelected.Enabled = False
    End If
End Sub

Private Sub cboSearchStr_Change()
    LastLine = -1
End Sub

Private Sub chkMatchCase_Click()
    MatchCase = IIf(chkMatchCase.Value = 1, True, False)
End Sub

Private Sub chkMatchWord_Click()
    MatchWord = IIf(chkMatchWord.Value = 1, True, False)
End Sub

Private Sub cmdClose_Click()
    frmMain.CurrentDoc.rtfHTML.SetFocus
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    DoFind
    UpdateLists
    LastSearch = cboSearchStr.Text
End Sub

Private Sub cmdReplace_Click()
Dim i%
    ' Switch to Replace mode if not in it, then exit sub
    If Not Replace Then
        Replace = True
        UpdateReplaceStatus
        Exit Sub
    End If
    
    If Len(cboSearchStr.Text) = 0 Then Exit Sub
    
    ' Replace next occurrence
    i = DoFind
    If i = -1 Then Exit Sub
    
    frmMain.CurrentDoc.rtfHTML.SelText = cboReplaceStr.Text
    frmMain.CurrentDoc.rtfHTML.SelStart = i
    frmMain.CurrentDoc.rtfHTML.SelLength = Len(cboReplaceStr.Text)
    
    UpdateLists
    LastSearch = cboSearchStr.Text
End Sub

Private Sub cmdReplaceAll_Click()
' Replace all occurrences
Dim Result%, C%, ST$, Compensate As Integer
Dim StartL As Long, EndL As Long
    If frmMain.Documents.Count = 0 Then
        MsgBox "No documents open for searching!", vbCritical, _
            "Ex Nihilo Error"
        Exit Sub
    End If

    If optSelected.Value = True Then
        StartL = frmMain.CurrentDoc.rtfHTML.SelStart
        EndL = StartL + frmMain.CurrentDoc.rtfHTML.SelLength
        
        If Len(cboReplaceStr.Text) > Len(cboSearchStr.Text) Then
            Compensate = 1
        ElseIf Len(cboSearchStr.Text) > Len(cboReplaceStr.Text) Then
            Compensate = 2
        End If
    Else
        StartL = 0
        EndL = Len(frmMain.CurrentDoc.rtfHTML.Text)
        Compensate = 0
    End If
    
    While Result <> -1
        Result = frmMain.CurrentDoc.rtfHTML.Find(cboSearchStr.Text, StartL, EndL, CreateFlags)
        
        If Result <> -1 Then
            frmMain.CurrentDoc.rtfHTML.SelText = ""
            frmMain.CurrentDoc.rtfHTML.SelText = cboReplaceStr.Text
            C = C + 1
            ' move past string for next search
            StartL = Result + Len(cboReplaceStr.Text)
            
            ' If the search string is longer or shorter than the replacement
            ' string, the ending character index will have to be changed each time
            ' a replacement is made.
            '
            If Compensate = 1 Then
                EndL = EndL + (Len(cboReplaceStr.Text) - Len(cboSearchStr.Text))
            ElseIf Compensate = 2 Then
                EndL = EndL - (Len(cboSearchStr.Text) - Len(cboReplaceStr.Text))
            End If
        End If
    Wend
    
    If C = 0 Then
        StatusBar1.Panels(1).Text = "No matches found"
    Else
        StatusBar1.Panels(1).Text = C & " replacements made."
    End If

    UpdateLists
    LastSearch = cboSearchStr.Text
End Sub

Private Sub Form_Load()
    UpdateReplaceStatus
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Set SearchHis = New Collection
    Set ReplaceHis = New Collection
'    Set rtbText = frmMain.CurrentDoc.rtfHTML
    If frmMain.CurrentDoc.rtfHTML.SelLength > 0 Then
        optSelected.Value = True
        If Not frmMain.CurrentDoc.rtfHTML.SelText Like "*" & vbCrLf & "*" Then
            cboSearchStr.Text = frmMain.CurrentDoc.rtfHTML.SelText
        End If
    Else
        optWholeText.Value = True
        optSelected.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
Static FindStr$, ReplaceStr$
    If Me.ActiveControl Is Nothing Then Exit Sub
    
    If Me.ActiveControl = cboSearchStr Or Replace = False Then
        If FindStr <> cboSearchStr.Text Then
            AutoComplete cboSearchStr
            FindStr = cboSearchStr.Text
            Exit Sub
        End If
    ElseIf Me.ActiveControl = cboReplaceStr Then
        If ReplaceStr <> cboReplaceStr.Text Then
            AutoComplete cboReplaceStr
            ReplaceStr = cboReplaceStr.Text
        End If
    End If
End Sub

' Performs a basic find based on selected options and
' returns the result (matching position or -1)
'
Private Function DoFind() As Integer
Dim Result As Long
Dim l As Long, R$
Dim StartL As Long, EndL As Long
    If frmMain.Documents.Count = 0 Then
        MsgBox "No documents open for searching!", vbCritical, _
            "Ex Nihilo Error"
        Exit Function
    End If
    
'    Set rtbText = frmMain.CurrentDoc.rtfHTML
    StartL = frmMain.CurrentDoc.rtfHTML.SelStart
        
    If Wrapped Then
        ' Start at the top and go down
        Locked = True
        Result = frmMain.CurrentDoc.rtfHTML.Find(cboSearchStr.Text, 0, , CreateFlags)
        If Result = -1 Then
            StatusBar1.Panels(1).Text = "No matches found"
        Else
            l = frmMain.CurrentDoc.rtfHTML.GetLineFromChar(Result)
            If LastLine = l Then
                R$ = "Only m"
            Else
                R$ = "M"
            End If
            StatusBar1.Panels(1).Text = R$ & "atch found on line " & l
            LastLine = l
        End If
        Wrapped = False
        DoFind = Result
        Locked = False
    Else
        If optSelected.Value = True Then
            EndL = StartL + frmMain.CurrentDoc.rtfHTML.SelLength
        Else
            StartL = StartL + 1
            EndL = Len(frmMain.CurrentDoc.rtfHTML.Text)
        End If
        
        ' Go down
        Locked = True
        Result = frmMain.CurrentDoc.rtfHTML.Find(cboSearchStr.Text, StartL, EndL, CreateFlags)
        If Result = -1 Then
            ' If only searching selected text then
            ' call it quits
            If optSelected.Value = True Then
                StatusBar1.Panels(1).Text = "No matches found in selected text"
                DoFind = Result
            ' Otherwise wrap around to beginning
            Else
                Wrapped = True
                DoFind = DoFind()    ' Recursively call to search again from top
            End If
        Else
            DoFind = Result
            l = frmMain.CurrentDoc.rtfHTML.GetLineFromChar(Result)
            If LastLine = l Then
                R$ = "Another m"
            Else
                R$ = "M"
            End If
            StatusBar1.Panels(1).Text = R$ & "atch found on line " & frmMain.CurrentDoc.rtfHTML.GetLineFromChar(Result)
            LastLine = l
            Locked = False
        End If
    End If
End Function

Private Function CreateFlags() As Integer
Dim FindFlags%
    FindFlags = 0
    If MatchCase And MatchWord Then
'        FindFlags = rtfMatchCase Or rtfWholeWord
    ElseIf MatchWord Then
'        FindFlags = rtfWholeWord
    ElseIf MatchCase Then
'        FindFlags = rtfMatchCase
    End If
    CreateFlags = FindFlags
End Function

' Adds new items to Find and Replace lists for use in
' AutoComplete.  Called whenever a new search or replace
' is performed.
Private Sub UpdateLists()
    If Not InList(cboSearchStr, cboSearchStr.Text) And _
        Len(cboSearchStr.Text) > 0 Then
            cboSearchStr.AddItem cboSearchStr.Text
    End If
    
    If Not Replace Then Exit Sub
    
    If Not InList(cboReplaceStr, cboReplaceStr.Text) And _
        Len(cboReplaceStr.Text) > 0 Then
            cboReplaceStr.AddItem cboReplaceStr.Text
    End If
End Sub

' Shuffle things around depending on whether this is
' a Find or a Find/Replace dialog
'
Private Sub UpdateReplaceStatus()
    If Not Replace Then
        cmdReplace.Caption = "&Replace..."
        cmdReplaceAll.Visible = False
        cboReplaceStr.Visible = False
        Label2.Visible = False
        
        cmdClose.Top = FIND_CLOSE_TOP
        fraOptions.Top = FIND_OPT_TOP
        fraScope.Top = FIND_OPT_TOP
        Me.Height = FIND_HEIGHT
    Else
        cmdReplace.Caption = "&Replace"
        frmFind.Caption = "Find/Replace"
        cmdReplaceAll.Visible = True
        cboReplaceStr.Visible = True
        Label2.Visible = True
        
        cmdClose.Top = REPLACE_CLOSE_TOP
        fraOptions.Top = REPLACE_OPT_TOP
        Me.Height = REPLACE_HEIGHT
        fraScope.Top = REPLACE_OPT_TOP
    End If
End Sub
