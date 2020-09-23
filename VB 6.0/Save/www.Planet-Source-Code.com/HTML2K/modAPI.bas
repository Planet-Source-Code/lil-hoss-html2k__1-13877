Attribute VB_Name = "modAPI"
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

Public Const ThisApp = "HTML2K"

Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line
Public Const EM_LINETOTAL = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public Const SB_HORZ = 0
Public Const EM_GETSEL = &HB0
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const SB_GETRECT As Long = (WM_USER + 10)
Public Const IMAGE_CURSOR = 2
Public Const LR_LOADFROMFILE = &H10
Public Const DI_NORMAL As Long = &H3
Public Const ERROR_NO_MORE_FILES = 18&
Public Const BS_FLAT = &H8000&
Public Const GWL_STYLE = (-16)
Public Const WS_CHILD = &H40000000
Public Const MF_BYPOSITION = &H400&
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETLINE = &HC4

Public TitreVar As String
Public LfVar As String
Public MsgVar As String
Public RepVar As Integer
Public m_cFlatten() As cFlatControl
Public m_iCount As Long
Public Templates As New Collection
Public TempFiles As New Collection
Public MRUFiles As New Collection
Public CodeClips As New Collection
Public xFile As String
Public xTitle As String
Public HTMLFile As Boolean
Public OpenFilename As String


Public Enum Status
    inTag = 0
    OutTag = 1
    InComment = 2
    OutComment = 3
    InScript = 4
End Enum
Public Enum ModeConstants
    VBWHTML = 1
End Enum
'Public Enum ERECViewModes
'   ercDefault = 0
'   ercWordWrap = 1
'   ercWYSIWYG = 2
'End Enum
Public Type CharRange
  cpMin As Long     '// First character of range (0 for start of doc)
  cpMax As Long     '// Last character of range (-1 for end of doc)
End Type
'// Word Wrap View Types
Public Enum ERECViewModes
   ercDefault = 0
   ercWordWrap = 1
   ercWYSIWYG = 2
End Enum
Public Enum ShowYesNoResult
    yes = 1
    YesToAll = 2
    no = 3
    NoToAll = 4
    None = 5
    Cancel = -1
End Enum
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
   X As Long
   Y As Long
End Type
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional parameter
    lpClass As String 'Optional parameter
    hkeyClass As Long 'Optional parameter
    dwHotKey As Long 'Optional parameter
    hIcon As Long 'Optional parameter
    hProcess As Long 'Optional parameter
End Type

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long

Public Type Tag
    Text As String
    Start As Double
    length As Double
End Type

Public Enum kFilePart
    kfpDrive
    kfpPath
    kfpDrivePath
    kfpExtension
    kfpFilename
    kfpPathFile
End Enum

Public Enum EShellShowConstants
     essSW_HIDE = 0
     essSW_MAXIMIZE = 3
     essSW_MINIMIZE = 6
     essSW_SHOWMAXIMIZED = 3
     essSW_SHOWMINIMIZED = 2
     essSW_SHOWNORMAL = 1
     essSW_SHOWNOACTIVATE = 4
     essSW_SHOWNA = 8
     essSW_SHOWMINNOACTIVE = 7
     essSW_SHOWDEFAULT = 10
     essSW_RESTORE = 9
     essSW_SHOW = 5
End Enum

Public ReadyToClose As Boolean
Public ShowSplash As Boolean
Public StartNewFile As Boolean
Public StatusPanel As Panel
Public AutosaveInterval As Integer
Public ItalicComments As Boolean
Public FlatToolbars As Boolean
Public AutoIndent As Boolean
Public UcaseTags As Boolean
Public TotalLines As Integer
Public CurrentLine As Integer
Public ShowLineNumbers As Boolean
Public FreeNum As Integer
Public LastSearch As String     ' Last search text

Public Const Filter = "Web Documents |*.htm;*.html|All Files (*.*)|*.*"   ' Common Dialog Filter constant


Public Sub LoadNewDoc()
    Dim f As New frmDoc
    Set f = New frmDoc
    f.Show
End Sub

Public Sub SaveDoc()
    Dim filenum As Integer
    'if form is empty or has not changed then exit
    If Left$(frmMain.CurrentDoc.Caption, 8) = "Untitled" Then
        'if form is newly created call File Save As
        SaveDocAs
    Else
        frmMain.CurrentDoc.rtfHTML.SaveFile (frmMain.CurrentDoc.Caption)
        frmMain.CurrentDoc.Modified = False
        DisableSaves
    End If
End Sub

Public Sub EnableSaves()
    frmMain.mnuFileSave.Enabled = True
    frmMain.mnuFileSaveAll.Enabled = True
    frmMain.mnuFileSaveAs.Enabled = True
    frmMain.mnuFileSaveAsTemplate.Enabled = True
    frmMain.Toolbar1.Buttons(3).Enabled = True
End Sub

Public Sub DisableSaves()
    frmMain.mnuFileSave.Enabled = False
    frmMain.mnuFileSaveAll.Enabled = False
    frmMain.mnuFileSaveAs.Enabled = False
    frmMain.mnuFileSaveAsTemplate.Enabled = False
    frmMain.Toolbar1.Buttons(3).Enabled = False
End Sub

Public Sub SaveDocAs()
'    On Error GoTo SaveError
    With frmMain.dlgFiles
        .DialogTitle = "Save file..."
        .Filter = Filter
        .ShowSave
    End With
    
    frmMain.CurrentDoc.rtfHTML.SaveFile (frmMain.dlgFiles.Filename)
    frmMain.CurrentDoc.Caption = frmMain.dlgFiles.Filename
    frmMain.CurrentDoc.Filename = frmMain.dlgFiles.Filename
    frmMain.CurrentDoc.Modified = False
    DisableSaves
    
'SaveError:
End Sub

Sub ErrHandler(Optional lngErrNum As Long = 0, Optional strErrorText As String = "", Optional strSource As String = "<Unknown>", Optional blnMustExit As Boolean = False, Optional strExtra As String = Empty, Optional bNoError As Boolean = False)
    Dim szErrMsg As String
    Dim intFileNum As Integer
    Load ErrMessage
    With ErrMessage
        If Dir(App.Path & "\Errorlog.txt") = Empty Then
            '// create file
            intFileNum = FreeFile
            Open App.Path & "\Errorlog.txt" For Output As intFileNum
            Close #intFileNum
        End If
        '// Write to log
        szErrMsg = "Err " & lngErrNum & ": " & strErrorText & vbCrLf
        If strExtra <> "" Then szErrMsg = szErrMsg & strExtra & vbCrLf
        szErrMsg = szErrMsg & "Last DLL Error: " & Err.LastDllError & vbCrLf
        szErrMsg = szErrMsg & "Error Source: " & Err.Source & vbCrLf
        szErrMsg = szErrMsg & "Procedure: " & strSource & vbCrLf
        szErrMsg = szErrMsg & Now & vbCrLf
        '// set the details text
        .txtDetails = szErrMsg
        .SetForm False '// ok, exit buttons
        .WritetoLog szErrMsg
        If blnMustExit Then .cmdOK.Enabled = False '// disable ok if critical
        .txtMsg.Text = .ChangeErrMsg(lngErrNum, strErrorText, strSource)
        .Show vbModal
        Unload ErrMessage
    End With
End Sub

Public Function AppPath() As String
    Dim a$
    
    a$ = App.Path
    AppPath = a$ & IIf(Right$(a$, 1) = "\", "", "\")
End Function
