VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New/Template"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox files 
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "O&ther..."
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdgDialog 
      Left            =   1320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstTemplate 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmNew"
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo OKErr
    Me.Hide
    If lstTemplate.Text = "(none)" Then
        Dim Dc As frmDoc
        Set Dc = New frmDoc
        Dc.Show
    Else
        Dim strFile As String
        strFile = App.Path & "\TemplateS\" & lstTemplate.List(lstTemplate.ListIndex) & ".html"
        frmMain.OpenDoc strFile, False
    End If
    Unload Me
    Exit Sub
OKErr:
    ErrHandler vbObjectError, "Error Loading Template", "Template Load Error", , , True
    Screen.MousePointer = 0
End Sub

Private Sub cmdOther_Click()
Dim T$
    On Error GoTo Canceled
    With cdgDialog
        .Filter = Filter
        .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
        .ShowOpen
        
        T$ = InputBox("Enter a name for this template:", "New Template", _
            GetFilePart(.Filename, kfpFilename))
        
        If LCase$(Right$(T$, 5)) <> ".html" And LCase$(Right$(T$, 4)) <> ".htm" Then
            T$ = T$ & ".html"
        End If
        
        FileCopy .Filename, AppPath & "templates\" & T$
        
        lstTemplate.AddItem T$
        Templates.Add AppPath & "templates\" & T$, AppPath & "templates\" & T$
    End With
Canceled:
End Sub

Private Sub cmdRemove_Click()
Dim i%, Name$
    i = lstTemplate.ListIndex
    Name = lstTemplate.Text
    
    lstTemplate.RemoveItem i
    Templates.Remove i
End Sub

Private Sub Form_Load()
    Dim i%
    
    lstTemplate.AddItem "(none)"
    
''    For i = 1 To Templates.Count
''        lstTemplate.AddItem GetFilePart(Templates(i), kfpFilename)
''    Next i
    LoadTemplates
    lstTemplate.ListIndex = 0
End Sub

Private Sub LoadTemplates()
' This function looks in the code directory and lists all files
' that have a .cbf extension : Code Bank File
    ' Variables
    Dim i As Long
    Dim intEntries As Integer
    ' Clear existing files
    ' Set the directory to the app.path / code directory

'    defFolder = App.Path & "\code\"
'    files.Path = defFolder
'    files.Path = App.Path & "\Code\"

    If Right(App.Path, 1) = "\" Then
        files.Path = App.Path & "\Templates\"
    Else
        files.Path = App.Path & "\Templates\"
    End If
    ' Set the file pattern
    files.Pattern = "*.html"
    ' Refresh list
    files.Refresh
    ' Display the files
    For i = 0 To files.ListCount - 1
        ' Add to the list box
        lstTemplate.AddItem Left(files.List(i), Len(files.List(i)) - 5)
        ' Add up the entries
        intEntries = intEntries + 1
    Next i
End Sub

Private Sub lstTemplate_Click()
    cmdRemove.Enabled = True
End Sub

Private Sub lstTemplate_DblClick()
    cmdOK_Click
End Sub
