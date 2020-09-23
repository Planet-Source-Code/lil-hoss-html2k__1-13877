VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F1A41040-2044-11D4-B705-0008C72B926D}#1.0#0"; "htmlText.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   6480
      TabIndex        =   32
      Top             =   1320
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      WordWrap        =   0   'False
      MouseIcon       =   "frmSettings.frx":0442
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmSettings.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "UpDown1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAutosave"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkNewFile"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkUpperCaseTag"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkShowSplash"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkCreateBackup"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkLineNumbers"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkWordWrap"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboWordWrap"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkAutoIndent"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkFlatToolbars"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Color"
      TabPicture(1)   =   "frmSettings.frx":047A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Picture4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Picture1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Picture2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Picture3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Picture5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Picture6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkItalic"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "htmlOption"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Templates"
      TabPicture(2)   =   "frmSettings.frx":0496
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTemplateName"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtTemplate"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboTemplate"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Advanced"
      TabPicture(3)   =   "frmSettings.frx":04B2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin htmlText.htmSyntaxBox htmlOption 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   40
         Top             =   1920
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmSettings.frx":04CE
         Text            =   ""
      End
      Begin VB.CheckBox chkFlatToolbars 
         Caption         =   "Flat Toolbars"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoIndent 
         Caption         =   "Autoindent text"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italicize comments"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   4680
         Width           =   1695
      End
      Begin VB.PictureBox Picture6 
         Height          =   255
         Left            =   -70560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   34
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cboWordWrap 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSettings.frx":04EA
         Left            =   3960
         List            =   "frmSettings.frx":04F7
         TabIndex        =   31
         Top             =   1750
         Width           =   1935
      End
      Begin VB.CheckBox chkWordWrap 
         Caption         =   "Word Wrap"
         Height          =   195
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkLineNumbers 
         Caption         =   "Show Line Numbers"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox chkCreateBackup 
         Caption         =   "Create Backup File When Saving"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   28
         Top             =   840
         Width           =   2775
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   -70560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   -70560
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkShowSplash 
         Caption         =   "Show Splash Screen"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkUpperCaseTag 
         Caption         =   "Uppercase Tags"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkNewFile 
         Caption         =   "New File on Startup"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtAutosave 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "1"
         Top             =   3240
         Width           =   270
      End
      Begin VB.Frame Frame1 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   5775
         Begin VB.CommandButton cmdClearMRU 
            Caption         =   "Clear MRU List"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdViewErrorLog 
            Caption         =   "Display &Error Log"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton cmdClearRegistry 
            Caption         =   "Clear &Registry"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Empties the Most Recently Used list."
            Height          =   195
            Left            =   2520
            TabIndex        =   33
            Top             =   1200
            Width           =   2580
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Displays the error log created by HTML2K."
            Height          =   195
            Left            =   2520
            TabIndex        =   10
            Top             =   1920
            Width           =   3000
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Clears all entries within the system registry."
            Height          =   195
            Left            =   2520
            TabIndex        =   9
            Top             =   480
            Width           =   2970
         End
      End
      Begin VB.ComboBox cboTemplate 
         Height          =   315
         ItemData        =   "frmSettings.frx":0522
         Left            =   -73320
         List            =   "frmSettings.frx":052C
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtTemplate 
         Height          =   3855
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   960
         Width           =   4935
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1710
         TabIndex        =   15
         Top             =   3240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtAutosave"
         BuddyDispid     =   196626
         OrigLeft        =   1920
         OrigTop         =   2160
         OrigRight       =   2160
         OrigBottom      =   2415
         Max             =   20
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Background Color"
         Height          =   195
         Left            =   -72000
         TabIndex        =   35
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Property Value"
         Height          =   195
         Left            =   -71760
         TabIndex        =   27
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Property Name"
         Height          =   195
         Left            =   -71760
         TabIndex        =   26
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tags"
         Height          =   195
         Left            =   -74520
         TabIndex        =   25
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Left            =   -74520
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   -74520
         TabIndex        =   23
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Autosave every"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "min."
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   3240
         Width           =   285
      End
      Begin VB.Label lblTemplateName 
         AutoSize        =   -1  'True
         Caption         =   "Template Name:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   4
         Top             =   625
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   6480
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
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

' WORD WRAP
'Private Sub mnuViewWrap0_Click()
'    mnuViewWrap0.Checked = True
'    mnuViewWrap1.Checked = False
'    SendMessageLong frmMain.CurrentDoc.rtfHtml.hWnd, EM_SETTARGETDEVICE, 0, 0
'End Sub
'
' NO WRAP
'Private Sub mnuViewWrap1_Click()
'    mnuViewWrap0.Checked = False
'    mnuViewWrap1.Checked = True
'    SendMessageLong frmMain.CurrentDoc.rtfHtml.hWnd, EM_SETTARGETDEVICE, 0, 1
'End Sub


Private Sub cboTemplate_Click()
    Static strLastTemplate As String
''    If bIgnoreClick Then Exit Sub
'    If SaveTemplate(strLastTemplate) = False Then
'        bIgnoreClick = True
'        cboTemplateName.Text = strLastTemplate
'        bIgnoreClick = False
'        Exit Sub
'    End If
    '// load the file
    Dim intFileNum As Integer
    intFileNum = FreeFile
    Open App.Path & "\templates\" & cboTemplate.Text & ".html" For Input As intFileNum
    txtTemplate = Input(LOF(intFileNum), intFileNum)
    Close #intFileNum
    strLastTemplate = cboTemplate.Text
'    blnTemplateChanged = False
End Sub

Private Sub chkAutoIndent_Click()
    EnableApply
End Sub

Private Sub chkFlatToolbars_Click()
    EnableApply
End Sub

Private Sub chkItalic_Click()
    EnableApply
    If chkItalic.Value = "1" Then
        htmlOption.CommentItalic = True
        htmlOption.Colorize
    Else
        htmlOption.CommentItalic = False
        htmlOption.Colorize
    End If
End Sub

Private Sub chkLineNumbers_Click()
    EnableApply
End Sub

Private Sub chkNewFile_Click()
    EnableApply
End Sub

Private Sub chkShowSplash_Click()
    EnableApply
End Sub

Private Sub chkUpperCaseTag_Click()
    EnableApply
End Sub

Private Sub chkWordWrap_Click()
    cboWordWrap.Enabled = IIf(chkWordWrap.Value = 1, True, False)
    If chkWordWrap.Value Then cboWordWrap.SetFocus
End Sub

Private Sub cmdApply_Click()
    ApplyChanges
    cmdApply.Enabled = False
End Sub

Private Sub ApplyChanges()
    StartNewFile = IIf(chkNewFile.Value = 1, True, False)
    ShowLineNumbers = IIf(chkLineNumbers.Value = 1, True, False)
    ShowSplash = IIf(chkShowSplash.Value = 1, True, False)
    UcaseTags = IIf(chkUpperCaseTag.Value = 1, True, False)
    AutosaveInterval = CInt(txtAutosave.Text)
    ItalicComments = IIf(chkItalic.Value = 1, True, False)
    AutoIndent = IIf(chkAutoIndent.Value = 1, True, False)
    FlatToolbars = IIf(chkFlatToolbars.Value = 1, True, False)
    
    SaveOptions
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearMRU_Click()
    Select Case MsgBox("Are you sure you want to empty the MRU list?", 68, "Confirm MRU Clear")
        Case vbYes
            DeleteSetting "HTML2K", "MRU List"
            MsgBox "The MRU list has been emptied!", vbOKOnly + vbInformation, "Done"
            Exit Sub
        Case vbNo
            MsgBox "Nothing was done.", vbOKOnly + vbInformation
    End Select
End Sub

Private Sub cmdClearRegistry_Click()
    Select Case MsgBox("Are you sure you want to delete the registry settings for HTML2K?", 68, "Confirm Registry Setting Deletion")
        Case vbYes
            DeleteSetting "HTML2K"
            cmdClearMRU.Visible = False
            MsgBox "HTML2K settings have been deleted!", vbOKOnly + vbInformation, "Done"
            Exit Sub
        Case vbNo
            MsgBox "No settings have been deleted.", vbOKOnly + vbInformation
    End Select
End Sub

Private Sub cmdOK_Click()
    ApplyChanges
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdViewErrorLog_Click()
    On Error GoTo ViewErrorLogErr
    OpenURL App.Path & "\Errorlog.txt"
    Exit Sub
ViewErrorLogErr:
    ErrHandler vbObjectError, "Error loading Errorlog.txt", "frmSettings.cmdViewErrorLog", , "Unable to load file", True
End Sub

Private Sub Form_Load()
    Dim txt As String
    txt = "<!DOCTYPE HTML PUBLIC """"-//W3C//DTD HTML 3.2//EN"""">" & vbCrLf
    txt = txt & "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf
    txt = txt & "<TITLE>(Type a title for your page here)</TITLE>" & vbCrLf
'    txt = txt & "<META NAME=""""GENERATOR"""" CONTENT=""""HTML2K"""">" & vbCrLf
    txt = txt & "</HEAD>" & vbCrLf
    txt = txt & "<BODY BACKGROUND="""""""" BGCOLOR=""""#ffffff"""" TEXT=""""#000000"""">" & vbCrLf
    txt = txt & "<!-- End Of Display Scrolling Message -->" & vbCrLf
    txt = txt & "</BODY>" & vbCrLf & "</HTML>"
    
    LoadValues
    
    htmlOption.Text = txt
    htmlOption.Colorize
    
    Picture1.BackColor = &H404040
    Picture2.BackColor = &H8000&
    Picture3.BackColor = &HFF0000
    Picture4.BackColor = &H800000
    Picture5.BackColor = &H80&
    LoadValues
    ListTemplates
    cboWordWrap.ListIndex = 0
    cmdApply.Enabled = False
End Sub

Private Function ListTemplates()
    cboTemplate.Clear
    AddAllFilesInDir "Templates", cboTemplate
    cboTemplate.ListIndex = 0
End Function

Private Sub LoadValues()
    On Error Resume Next
    chkShowSplash.Value = GetSetting(ThisApp, "Options", "ShowSplash")
    chkNewFile.Value = GetSetting(ThisApp, "Options", "StartNewFile")
    chkUpperCaseTag.Value = GetSetting(ThisApp, "Options", "UcaseTags")
    txtAutosave.Text = GetSetting(ThisApp, "Options", "AutosaveInterval")
    If AutosaveInterval = "" Then
        txtAutosave.Text = "5"
    End If
    chkItalic.Value = GetSetting(ThisApp, "Options", "ItalicComments")
    chkAutoIndent.Value = GetSetting(ThisApp, "Options", "AutoIndent")
    chkFlatToolbars.Value = GetSetting(ThisApp, "Options", "FlatToolbars")
    chkLineNumbers.Value = GetSetting(ThisApp, "Options", "ShowLineNumbers")

    On Error Resume Next
    htmlOption.BackColor = GetSetting(ThisApp, "Options", "Background Color")
    Picture6.BackColor = htmlOption.BackColor
    htmlOption.CommentColor = GetSetting(ThisApp, "Options", "Comment Color")
    Picture2.BackColor = htmlOption.CommentColor
    htmlOption.EntityColor = GetSetting(ThisApp, "Options", "Entity Color")
    Picture1.BackColor = htmlOption.EntityColor
    htmlOption.PropNameColor = GetSetting(ThisApp, "Options", "Property Color")
    Picture4.BackColor = htmlOption.PropNameColor
    htmlOption.PropValColor = GetSetting(ThisApp, "Options", "Value Color")
    Picture5.BackColor = htmlOption.PropValColor
    htmlOption.TagColor = GetSetting(ThisApp, "Options", "Tag Color")
    Picture3.BackColor = htmlOption.TagColor
    htmlOption.Locked = True
End Sub

Private Sub Picture1_Click()
    GetColor Picture1
    EnableApply
End Sub

Private Sub Picture2_Click()
    GetColor Picture2
    EnableApply
End Sub

Private Sub Picture3_Click()
    GetColor Picture3
    EnableApply
End Sub

Private Sub Picture4_Click()
    GetColor Picture4
    EnableApply
End Sub

Private Sub Picture5_Click()
    GetColor Picture5
    EnableApply
End Sub

Private Sub Picture6_Click()
    GetColor Picture6
    EnableApply
End Sub

Private Sub UpDown1_Change()
    EnableApply
End Sub

Private Sub EnableApply()
    cmdApply.Enabled = True
End Sub

Private Function GetColor(pic As PictureBox)
'    On Error GoTo ColorError
    dlgColors.Color = pic.BackColor
    dlgColors.Flags = &H1&
    dlgColors.ShowColor
    pic.BackColor = dlgColors.Color
    
    htmlOption.EntityColor = Picture1.BackColor
    htmlOption.CommentColor = Picture2.BackColor
    htmlOption.TagColor = Picture3.BackColor
    htmlOption.PropNameColor = Picture4.BackColor
    htmlOption.PropValColor = Picture5.BackColor
    htmlOption.BackColor = Picture6.BackColor
    htmlOption.Colorize
    
'ColorError:
'    MsgBox "An error has occurred!"
End Function
