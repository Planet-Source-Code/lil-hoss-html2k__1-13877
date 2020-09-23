VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "CPOPMENU.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "HTML2K [Beta Preview 1.0]"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAutoSave 
      Interval        =   60000
      Left            =   2160
      Top             =   480
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Create a new document (Ctlr+N)."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open a document (Ctrl+O)."
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save current document (Ctrl+S)."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CUT"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COPY"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PASTE"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UNDO"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DELETE"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "REDO"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WEB"
            Object.ToolTipText     =   "View Web Document with default viewer"
            ImageIndex      =   59
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TABLE"
            Object.ToolTipText     =   "Create a table."
            ImageIndex      =   27
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "TH"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "TR"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "TD"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "CAPTION"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "TABLE"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LIST"
            Object.ToolTipText     =   "Create a list."
            ImageIndex      =   29
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "LI"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "DT"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "DD"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "MENU"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FRAME"
            Object.ToolTipText     =   "Create frames."
            ImageIndex      =   31
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "IFRAMES"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "FRAMESET"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "FRAME"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "NOFRAME"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BOLD"
            ImageIndex      =   49
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ITALIC"
            ImageIndex      =   50
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UNDERLINE"
            ImageIndex      =   51
         EndProperty
      EndProperty
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   9360
         Pattern         =   "*.htm;*.html"
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   1440
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   61
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28FC
            Key             =   "TILEHOR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A56
            Key             =   "CASCADE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BB0
            Key             =   "NEWWINDOW"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D8A
            Key             =   "ARRANGEICONS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3324
            Key             =   "TILEVERT"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":347E
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35D8
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3732
            Key             =   "CLOSEALL"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":388C
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39E6
            Key             =   "SAVEAS"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B40
            Key             =   "H1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D1A
            Key             =   "H2"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EF4
            Key             =   "H3"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40CE
            Key             =   "H4"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42A8
            Key             =   "H5"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4482
            Key             =   "H6"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":465C
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47B6
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4910
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A6A
            Key             =   "REDO"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BC4
            Key             =   "UNDO"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D1E
            Key             =   "SEND"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5170
            Key             =   "PAGESETUP"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52CA
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5424
            Key             =   "EXIT"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":557E
            Key             =   "TABLE"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56D8
            Key             =   "LISTBOX"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58B2
            Key             =   "LIST"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A0C
            Key             =   "FRAME2"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B66
            Key             =   "FRAME"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CC0
            Key             =   "PUSHPIN2"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E1A
            Key             =   "PUSHPIN"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":63B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":650E
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6668
            Key             =   "TEXTBOX"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6842
            Key             =   "PASSWORD"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":699C
            Key             =   "HIDDEN"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AF6
            Key             =   "CHECKBOX"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CD0
            Key             =   "RADIOBUTTON"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EAA
            Key             =   "COMBOBOX"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7084
            Key             =   "SUBMIT"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71DE
            Key             =   "RESET"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7338
            Key             =   "HYPERLINK"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7512
            Key             =   "SOUND"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":766C
            Key             =   "COLOR"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77C6
            Key             =   "MARQUEE"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7920
            Key             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D72
            Key             =   "BOLD"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7ECC
            Key             =   "ITALIC"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8026
            Key             =   "UNDERLINE"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85C0
            Key             =   "ABOUT"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":871A
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8CB4
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E0E
            Key             =   "CENTER"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F68
            Key             =   "RIGHT"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":90C2
            Key             =   "PARAGRAPH"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":921C
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9376
            Key             =   "WEB"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94D0
            Key             =   "SPELLCHECK"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":962A
            Key             =   "PRINTPREVIEW"
         EndProperty
      EndProperty
   End
   Begin cPopMenu.PopMenu ctlPopMenu 
      Left            =   720
      Top             =   480
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2625
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6174
            MinWidth        =   6174
            Key             =   "lsp"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1147
            MinWidth        =   1147
            Text            =   "ln X"
            TextSave        =   "ln X"
            Key             =   "row"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1147
            MinWidth        =   1147
            Text            =   "col X"
            TextSave        =   "col X"
            Key             =   "col"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "X"
            TextSave        =   "X"
            Key             =   "total"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   794
            MinWidth        =   794
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Clos&e All"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAsTemplate 
         Caption         =   "Save as &Template..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuMRUSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find Next"
      End
      Begin VB.Menu mnuSearchFindPrevious 
         Caption         =   "Find Previous"
      End
      Begin VB.Menu mnuEditSearchAgain 
         Caption         =   "Search Again"
      End
      Begin VB.Menu mnuSearchSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchReplace 
         Caption         =   "&Replace"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTestDefault 
         Caption         =   "Test with &Default Browser..."
      End
   End
   Begin VB.Menu mnuHTML 
      Caption         =   "&HTML"
      Begin VB.Menu mnuHTMLSource 
         Caption         =   "Source"
         Begin VB.Menu mnuHTMLSourceHierachal 
            Caption         =   "&Hierarchal"
         End
         Begin VB.Menu mnuHTMLSourceSimple 
            Caption         =   "&Simple"
         End
         Begin VB.Menu mnuHTMLSourceCompact 
            Caption         =   "&Compact"
         End
         Begin VB.Menu mnuHTMLSourceSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHTMLStripLinks 
            Caption         =   "Strip Links"
         End
         Begin VB.Menu mnuHTMLSourceStripText 
            Caption         =   "Strip Text"
         End
      End
      Begin VB.Menu mnuHTMLSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLineBreak 
         Caption         =   "Line Break"
      End
      Begin VB.Menu mnuHTMLParagraph 
         Caption         =   "Paragraph"
      End
      Begin VB.Menu mnuHorizonLine 
         Caption         =   "Horizon Line"
      End
      Begin VB.Menu mnuHardBreak 
         Caption         =   "Hard Break"
      End
      Begin VB.Menu mnuWordBreak 
         Caption         =   "Word Break"
      End
      Begin VB.Menu mnuNoBreak 
         Caption         =   "No Break"
      End
      Begin VB.Menu mnuHTMLSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTMLBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuHTMLItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuHTMLUnderline 
         Caption         =   "Underline"
      End
      Begin VB.Menu mnuHTMLSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHeading 
         Caption         =   "&Headings"
         Begin VB.Menu mnuH1 
            Caption         =   "H1"
         End
         Begin VB.Menu mnuH2 
            Caption         =   "H2"
         End
         Begin VB.Menu mnuH3 
            Caption         =   "H3"
         End
         Begin VB.Menu mnuH4 
            Caption         =   "H4"
         End
         Begin VB.Menu mnuH5 
            Caption         =   "H5"
         End
         Begin VB.Menu mnuH6 
            Caption         =   "H6"
         End
      End
      Begin VB.Menu mnuHTMLSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTMLText 
         Caption         =   "Text Formatting"
         Begin VB.Menu mnuPRE 
            Caption         =   "Preformatted..."
         End
         Begin VB.Menu mnuTT 
            Caption         =   "Teletype..."
         End
         Begin VB.Menu mnuADDRESS 
            Caption         =   "Address..."
         End
         Begin VB.Menu mnuIns 
            Caption         =   "Insertion..."
         End
         Begin VB.Menu mnuEM 
            Caption         =   "Emphasis..."
         End
         Begin VB.Menu mnuSAMPLE 
            Caption         =   "Sample..."
         End
         Begin VB.Menu mnuCODE 
            Caption         =   "Code..."
         End
         Begin VB.Menu mnuSTRONG 
            Caption         =   "Strong..."
         End
         Begin VB.Menu mnuVAR 
            Caption         =   "Var..."
         End
         Begin VB.Menu mnuBLOCKQUOTE 
            Caption         =   "Blockquote..."
         End
         Begin VB.Menu mnuSUB 
            Caption         =   "Subscript..."
         End
         Begin VB.Menu mnuSUP 
            Caption         =   "Superscript..."
         End
      End
      Begin VB.Menu mnuHTMLSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTMLCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuHTMLRight 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuHTMLLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuHTMLJustify 
         Caption         =   "Justify"
      End
      Begin VB.Menu mnuHTMLSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTMLFont1 
         Caption         =   "Font size + 1"
      End
      Begin VB.Menu mnuHTMLFont 
         Caption         =   "Font size - 1"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuElements 
         Caption         =   "&Elements"
         Begin VB.Menu mnuHyperlink 
            Caption         =   "&Hyperlink..."
         End
         Begin VB.Menu mnuBookmark 
            Caption         =   "&Bookmark..."
         End
         Begin VB.Menu mnuElementEMail 
            Caption         =   "&E-Mail..."
         End
         Begin VB.Menu mnuGraphic 
            Caption         =   "&Graphic..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSound 
            Caption         =   "&Sound..."
         End
         Begin VB.Menu mnuComment 
            Caption         =   "&Comment..."
         End
         Begin VB.Menu mnuFont 
            Caption         =   "&Font..."
         End
         Begin VB.Menu mnuBasecolor 
            Caption         =   "Ba&secolor..."
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuForms 
         Caption         =   "&Forms"
         Begin VB.Menu mnuScrollingTextBox 
            Caption         =   "&Scrolling Text Box..."
         End
         Begin VB.Menu mnuOneLineTextBox 
            Caption         =   "&One Line Text Box..."
         End
         Begin VB.Menu mnuPassword 
            Caption         =   "&Password..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuHidden 
            Caption         =   "&Hidden..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCheckbox 
            Caption         =   "&Checkbox..."
         End
         Begin VB.Menu mnuRadioButton 
            Caption         =   "&Radio Button..."
         End
         Begin VB.Menu mnuComboBox 
            Caption         =   "&Combo Box..."
         End
      End
      Begin VB.Menu mnuInsertSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertMarquee 
         Caption         =   "&Marquee"
      End
      Begin VB.Menu mnuInsertSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertLastUpdated 
         Caption         =   "&Last Updated"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuViewCodeClips 
         Caption         =   "Code Clips"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsColorPicker 
         Caption         =   "Color Picker"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSpellCheck 
         Caption         =   "&Spell Check"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Fav&orites"
      Begin VB.Menu mnuFavs 
         Caption         =   ""
         Index           =   1
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowNext 
         Caption         =   "&Switch to Next Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHTML 
         Caption         =   "HTML Help..."
      End
      Begin VB.Menu mnuHelpStyleSheet 
         Caption         =   "Style Sheet Help..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpToDo 
         Caption         =   "&To Do List..."
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About HTML2K..."
      End
   End
   Begin VB.Menu mnuTSBars 
      Caption         =   "TSBars"
      Visible         =   0   'False
      Begin VB.Menu mnuTB 
         Caption         =   "View Toolbar"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuTB 
         Caption         =   "View Statusbar"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuDummy 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDummyUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuDummyRedo 
         Caption         =   "&Redo"
      End
   End
   Begin VB.Menu mnuEditTag 
      Caption         =   "Edit Tag"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
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

Public MsgMode As String
Public Documents As New Collection
Public Filename As String
Public CurrentDoc As frmDoc
Public ExitProcess As Boolean

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long


Private Sub MDIForm_Initialize()
    On Error Resume Next
    frmMain.Left = GetSetting(ThisApp, "Settings", "Left")
    frmMain.Top = GetSetting(ThisApp, "Settings", "Top")
    frmMain.Width = GetSetting(ThisApp, "Settings", "Width")
    frmMain.Height = GetSetting(ThisApp, "Settings", "Height")
    
    LoadOptions
End Sub

Private Sub MDIForm_Load()
    Dim tmpName As String
    Dim tmpCaption As String
    Dim i As Integer
    
    On Error GoTo LoadError
    If Documents.Count = 0 Then
        NoDocs
    End If
    
    LoadIcons
    
    File1.Path = App.Path & "\Templates\"
    For i = 1 To File1.ListCount
        On Error Resume Next
        tmpName = File1.List(i - 1)
        Load mnuFavs(i)
        mnuFavs(i).Caption = tmpName  ' Left$(tmpName, Len(tmpName) - 4)
    Next i

    ' Open file if passed on command line
    If Len(Command$) > 0 Then
        OpenDoc Command, True
    End If

    FileFillCol TempFiles, AppPath & "tempdoc.his"
    ' Check for temp documents
    If TempFiles.Count > 0 Then
        For i = 1 To TempFiles.Count
            OpenDoc TempFiles(i), False
            With CurrentDoc
                .SetModified True
                .UnSaved = True
                .TempName = .Filename
                .Filename = ""
            End With
        Next i
        
        MsgBox CStr(TempFiles.Count) & " files recovered " & _
            "from a previous session.", vbInformation, "Documents Recovered"
    End If
    Exit Sub
    
LoadError:
    ErrHandler vbObjectError, "Error Loading HTML2K", "Load", , , True
End Sub

' ========================================================
' = Here's the code to add incons to the pull-down menus =
' ========================================================
Private Sub pSetIcon(ByVal sIconKey As String, ByVal sMenuKey As String)
    Dim lIconIndex As Long
    lIconIndex = plGetIconIndex(sIconKey)
    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
End Sub '

Private Function plGetIconIndex(ByVal sKey As String) As Long
    plGetIconIndex = ilsIcons.ListImages.item(sKey).Index - 1
End Function

Sub LoadIcons()
    Dim i As Integer
    Dim l As Long
    Dim lIndex As Long
    Dim lc As Long '

    On Error GoTo LoadIconError
    With ctlPopMenu
        ' Associate the image list:
        .ImageList = ilsIcons
        ' Parse through the VB designed menu and sub class the items:
        .SubClassMenu Me
        lIndex = .MenuIndex("mnuFile")
        ' Add the icons:
        ' File pull-down...
        pSetIcon "NEW", "mnuFileNew"
        pSetIcon "OPEN", "mnuFileOpen"
        pSetIcon "CLOSE", "mnuFileClose"
        pSetIcon "SAVE", "mnuFileSave"
        pSetIcon "SAVEAS", "mnuFileSaveAs"
        pSetIcon "PRINT", "mnuFilePrint"
        pSetIcon "PRINTPREVIEW", "mnuFilePrintPreview"
        pSetIcon "SEND", "mnuFileSend"
        pSetIcon "EXIT", "mnuFileExit"
        pSetIcon "PAGESETUP", "mnuFilePageSetup"

        pSetIcon "UNDO", "mnuEditUndo"
        pSetIcon "REDO", "mnuEditRedo"
        pSetIcon "CUT", "mnuEditCut"
        pSetIcon "COPY", "mnuEditCopy"
        pSetIcon "PASTE", "mnuEditPaste"

        pSetIcon "PUSHPIN2", "mnuViewToolbar"
        pSetIcon "PUSHPIN2", "mnuViewStatusBar"
        pSetIcon "SETTINGS", "mnuViewOptions"
        pSetIcon "WEB", "mnuTestDefault"

        pSetIcon "H1", "mnuH1"
        pSetIcon "H2", "mnuH2"
        pSetIcon "H3", "mnuH3"
        pSetIcon "H4", "mnuH4"
        pSetIcon "H5", "mnuH5"
        pSetIcon "H6", "mnuH6"
        
        pSetIcon "HYPERLINK", "mnuHyperlink"
        pSetIcon "SOUND", "mnuSound"
        pSetIcon "MARQUEE", "mnuInsertMarquee"
        pSetIcon "PASSWORD", "mnuPassword"
        pSetIcon "COMBOBOX", "mnuComboBox"
        pSetIcon "RADIOBUTTON", "mnuRadioButton"
        pSetIcon "CHECKBOX", "mnuCheckbox"
        pSetIcon "TEXTBOX", "mnuOneLineTextBox"
        pSetIcon "HIDDEN", "mnuHidden"
        
        pSetIcon "BOLD", "mnuHTMLBold"
        pSetIcon "ITALIC", "mnuHTMLItalic"
        pSetIcon "UNDERLINE", "mnuHTMLUnderline"
        pSetIcon "CENTER", "mnuHTMLCenter"
        pSetIcon "LEFT", "mnuHTMLLeft"
        pSetIcon "RIGHT", "mnuHTMLRight"
        
        pSetIcon "SPELLCHECK", "mnuToolsSpellCheck"
        
        pSetIcon "NEWWINDOW", "mnuWindowNewWindow"
        pSetIcon "CASCADE", "mnuWindowCascade"
        pSetIcon "TILEHOR", "mnuWindowTileHorizontal"
        pSetIcon "TILEVERT", "mnuWindowTileVertical"
        pSetIcon "ARRANGEICONS", "mnuWindowArrangeIcons"
    End With
    Exit Sub
LoadIconError:
    ErrHandler vbObjectError, "Error Loading Menu Icons", "LoadIcons", , , True
End Sub
'' ================================
' = End of menu icon source code =
' ================================

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbStatusBar.Panels("lsp").Text = ""  ' Clear when not over ToolBar
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    If Me.ScaleWidth < 8415 Then
        Me.Width = 8415
    End If
'    If Me.ScaleHeight < 3500 Then
'        Me.Height = 3500
'    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Integer
    Dim num As Integer
    
    For num = Forms.Count - 1 To 1 Step -1
        Unload Forms(num)
    Next
    For i = 1 To Documents.Count
        If Cancel = 1 Then Exit Sub
    Next i
    
    SaveSetting ThisApp, "Settings", "Left", frmMain.Left
    SaveSetting ThisApp, "Settings", "Top", frmMain.Top
    SaveSetting ThisApp, "Settings", "Width", frmMain.Width
    SaveSetting ThisApp, "Settings", "Height", frmMain.Height
    
    For i = 1 To MRUFiles.Count
        MRURegPoke "MRUFile" & CStr(i), MRUFiles(i)
    Next i
End Sub

Private Sub mnuADDRESS_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<ADDRESS>"), CaseTag("</ADDRESS>")
End Sub

Private Sub mnuBLOCKQUOTE_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<BLOCKQUOTE>"), CaseTag("</BLOCKQUOTE>")
End Sub

Private Sub mnuBookmark_Click()
    frmHTMLBookmark.Show vbModal, Me
End Sub

Private Sub mnuCheckbox_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P><input type=""") & CaseTag("checkbox") & """" & CaseTag(" name=") & """" & CaseTag("C1") & """" & ">", CaseTag("</P>")
End Sub

Private Sub mnuCODE_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<CODE>"), CaseTag("</CODE>")
End Sub

Private Sub mnuComboBox_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P><select name=""") & CaseTag("D1") & """" & CaseTag(" size=") & """" & "1" & """" & ">", CaseTag("</SELECT></P>")
End Sub

Private Sub mnuComment_Click()
    frmMain.CurrentDoc.InsertSurroundTag ("<!-- "), (" -->")
End Sub

Private Sub mnuEditCopy_Click()
    CopyText
End Sub

Private Sub mnuEditCut_Click()
    CutText
End Sub

Private Sub mnuEditDelete_Click()
    CurrentDoc.rtfHTML.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
'    PasteText
End Sub

Private Sub mnuEditRedo_Click()
    CurrentDoc.Redo
End Sub

Private Sub mnuEditSearchAgain_Click()
 Dim Result As Long
    ' Search from current position
    Result = CurrentDoc.rtfHTML.Find(LastSearch, CurrentDoc.rtfHTML.SelStart + 1)
    If Result = -1 Then
        ' Try whole text
        Result = CurrentDoc.rtfHTML.Find(LastSearch, 0, Len(CurrentDoc.rtfHTML.Text))
        
        If Result = -1 Then
            MsgBox "Search text not found:" & vbCrLf & LastSearch, _
                vbExclamation, "Not Found"
        End If
    End If
End Sub

Private Sub mnuEditSelectAll_Click()
    SelectAllText
End Sub

''''Private Sub mnuEditTag_Click()
''''On Error Resume Next
''''Dim TagRec As String
''''Dim FullTag As String
''''Dim TagRecNew
''''  Me.CurrentDoc.rtfHTML.Span "<", False, True        ' Select Full Tag
''''  Me.CurrentDoc.rtfHTML.Span ">", True, True         ' Select Full Tag
''''  TagRec = "<" & Me.CurrentDoc.rtfHTML.SelText & ">" ' Add <> to the tag
''''
''''  TagRec = Mid(Me.CurrentDoc.rtfHTML.SelText, 1, 2)
''''    If Mid(TagRec, 1, 1) = "/" Then
''''      MsgBox "Can't edit the end of TAG", vbInformation
''''      Exit Sub
''''    End If
'''''body
''''    If LCase(TagRec) = "bo" Then
''''''     TabNumber = 1
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''div
''''    If LCase(TagRec) = "di" Then
''''''     TabNumber = 3
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''anchor
''''    If LCase(TagRec) = "a " Then
''''''     TabNumber = 0
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''font
''''    If LCase(TagRec) = "fo" Then
''''''     TabNumber = 4
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''img
''''    If LCase(TagRec) = "im" Then
''''''     TabNumber = 6
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''select
''''    If LCase(TagRec) = "se" Then
'''''     TabNumber = 8
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''hr
''''    If LCase(TagRec) = "hr" Then
'''''     TabNumber = 5
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''textarea
''''    If LCase(TagRec) = "te" Then
'''''     TabNumber = 10
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''table
''''    If LCase(TagRec) = "ta" Then
'''''     TabNumber = 13
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''Td
''''    If LCase(TagRec) = "td" Then
'''''     TabNumber = 14
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
'''''tr
''''    If LCase(TagRec) = "tr" Then
'''''     TabNumber = 15
'''''     frmTagEdit.Show 1, fMainForm
''''    MsgBox "Test", vbInformation
''''     Exit Sub
''''    End If
''''
''''  If LCase(TagRec) = "in" Then
''''    TagRecNew = LCase(Mid(Me.CurrentDoc.rtfHTML.SelText, 13, 3))
''''
'''''Submit
''''     If TagRecNew = Chr(34) & "sub" Then
'''' '     TabNumber = 9
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''     If TagRecNew = "sub" Then
'''' '     TabNumber = 9
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
'''''Radio
''''     If TagRecNew = Chr(34) & "rad" Then
'''' '     TabNumber = 7
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''     If TagRecNew = "rad" Then
'''' '     TabNumber = 7
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
'''''Checkbox
''''     If TagRecNew = Chr(34) & "che" Then
'''' '     TabNumber = 2
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''     If TagRecNew = "che" Then
'''' '     TabNumber = 2
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
'''''text
''''     If TagRecNew = Chr(34) & "tex" Then
'''' '     TabNumber = 12
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''     If TagRecNew = "tex" Then
'''' '     TabNumber = 12
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
'''''hidden
''''     If TagRecNew = Chr(34) & "hid" Then
'''' '     TabNumber = 11
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''     If TagRecNew = "hid" Then
'''' '     TabNumber = 11
'''' '     frmTagEdit.Show 1, fMainForm
''''     MsgBox "Test", vbInformation
''''      Exit Sub
''''     End If
''''  End If
'''''If TabNumber = 0 Then MsgBox "Tag not supported by TagEditor", vbInformation, "Unsupported Tag"
''''' End If
''''End Sub

Private Sub mnuEditUndo_Click()
    CurrentDoc.Undo
End Sub

Private Sub mnuElementEMail_Click()
    frmHTMLEMail.Show vbModal, Me
End Sub

Private Sub mnuEM_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<EM>"), CaseTag("</EM>")
End Sub

Private Sub mnuFavs_Click(Index As Integer)
    OpenDoc App.Path & "\Templates\" & mnuFavs(Index).Caption, False
End Sub

Private Sub mnuFileClose_Click()
    CurrentDoc.CloseDoc
End Sub

Private Sub mnuFileCloseAll_Click()
    Dim d As frmDoc
    
    For Each d In Documents
        d.CloseDoc
    Next d
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    frmNew.Show vbModal, Me
'    LoadNewDoc
End Sub

Public Sub NoDocs()
    mnuFileSave.Enabled = False
    mnuFileSaveAs.Enabled = False
    mnuFileSaveAll.Enabled = False
    mnuFileSaveAsTemplate.Enabled = False
    mnuFileClose.Enabled = False
    mnuFileCloseAll.Enabled = False
    mnuFileProperties.Enabled = False
    mnuFilePrint.Enabled = False
    mnuFileSend.Enabled = False
    mnuEdit.Enabled = False
    mnuTestDefault.Enabled = False
    mnuSearch.Enabled = False
    mnuHTML.Enabled = False
    mnuInsert.Enabled = False
    mnuTools.Enabled = False
    mnuWindow.Enabled = False
    Dim i As Integer
    For i = 3 To 20
        Toolbar1.Buttons(i).Enabled = False
    Next i
    sbStatusBar.Panels("row").Text = ""
    sbStatusBar.Panels("col").Text = ""
    sbStatusBar.Panels("total").Text = ""
End Sub

Public Sub AllDocs()
    mnuFileSave.Enabled = True
    mnuFileSaveAs.Enabled = True
    mnuFileSaveAll.Enabled = True
    mnuFileSaveAsTemplate.Enabled = True
    mnuFileClose.Enabled = True
    mnuFileCloseAll.Enabled = True
    mnuFileProperties.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileSend.Enabled = True
    mnuEdit.Enabled = True
    mnuTestDefault.Enabled = True
    mnuSearch.Enabled = True
    mnuHTML.Enabled = True
    mnuInsert.Enabled = True
    mnuTools.Enabled = True
    mnuWindow.Enabled = True
    Dim i As Integer
    For i = 3 To 20
        Toolbar1.Buttons(i).Enabled = True
    Next i
End Sub

Private Sub mnuFileOpen_Click()
    OpenDoc "", True
End Sub

Private Sub mnuFilePageSetup_Click()
    frmPageSetup.Show vbModal, Me
End Sub

Private Sub mnuFilePrint_Click()
    PrintDoc
End Sub

Private Sub mnuFileProperties_Click()
    Dim R As Long
    Dim fName As String
    'get the path and filename of the current document
    fName = CurrentDoc.Filename
    'show the properties dialog, passing the filename and the owner of the dialog
    R = ShowProperties(fName, Me.hwnd)
    'Display an error message if things didn't go as planned
    If R <= 32 Then MsgBox "The current file has not been saved!", vbInformation + vbOKOnly, "Properties Not Available"
End Sub

Private Sub mnuFileSave_Click()
    SaveDoc
End Sub

Private Sub mnuFileSaveAll_Click()
 Dim d As frmDoc
    For Each d In Documents
        ' If this document has never been saved,
        ' we do a Save As, since the user must specify
        ' a filename.
        '
        If d.UnSaved Then
            mnuFileSaveAs_Click
        Else
            d.rtfHTML.SaveFile CurrentDoc.Filename
            d.SetModified False
            
            ' If the doc has been saved in a temp file,
            ' kill the temp file and remove it from the
            ' temp file list, since we don't need it any more
            If Len(d.TempName) > 0 Then
                TempFiles.Remove d.TempName
                FileDumpCol TempFiles, AppPath & "tempdoc.his"
                Kill d.TempName
                d.TempName = ""
            End If

        End If
    Next d
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveDocAs
End Sub

Private Sub mnuFileSaveAsTemplate_Click()
    If Len(Dir(AppPath & "Templates", vbDirectory)) = 0 Then
        MkDir AppPath & "Templates"
    End If
    
    On Error GoTo Canceled
    With dlgFiles
        .DialogTitle = "Save as Template"
        .InitDir = AppPath & "Templates"
        .Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
        .Filter = "HTML Document (*.html, *.htm)|*.htm;*.html|Cascading Style Sheet (*.css)|*.css|All files (*.*)|*.*"
        
        .ShowSave
        CurrentDoc.rtfHTML.SaveFile .Filename
        
        MsgBox "A copy of this document has been saved in the templates directory.", _
            vbInformation, "Template saved"
        
        Templates.Add .Filename
    End With
Canceled:

End Sub

Private Sub mnuFileSend_Click()
    MsgBox "This option is currently under construction...", vbInformation + vbOKOnly, "Option Not Available"
'    OpenURL "mailto:mydixiewrecked2@hotmail.com?subject=Test?body=" & CurrentDoc.rtfHTML.SelText
End Sub

Private Sub mnuFont_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<FONT SIZE=""" & "1" & """ COLOR=""" & "#000000""" & " FACE=""" & "TIMES NEW ROMAN" & """>"), CaseTag("</FONT>")
End Sub

Private Sub mnuH1_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H1>"), CaseTag("</H1>")
End Sub

Private Sub mnuH2_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H2>"), CaseTag("</H2>")
End Sub

Private Sub mnuH3_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H3>"), CaseTag("</H3>")
End Sub

Private Sub mnuH4_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H4>"), CaseTag("</H4>")
End Sub

Private Sub mnuH5_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H5>"), CaseTag("</H5>")
End Sub

Private Sub mnuH6_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<H6>"), CaseTag("</H6>")
End Sub

Private Sub mnuHardBreak_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CaseTag("&nbsp;")
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpHTML_Click()
    WinHelp Me.hwnd, App.Path & "\Help\HTML\index.hlp", cdlHelpContents, 0
End Sub

Private Sub mnuHelpStyleSheet_Click()
    WinHelp Me.hwnd, App.Path & "\Help\CSS\CSS.hlp", cdlHelpContents, 0
End Sub

Private Sub mnuHelpToDo_Click()
'    Display ReadMe file
    frmFileView.Show
    xFile = "xx"
End Sub

Private Sub mnuHorizonLine_Click()
    frmHTMLHorizonRule.Show vbModal, Me
End Sub

Private Sub mnuHTMLBold_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<B>"), CaseTag("</B>")
End Sub

Private Sub mnuHTMLCenter_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P ALIGN=""CENTER"">"), CaseTag("</P>")
End Sub

Private Sub mnuHTMLFont_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<FONT SIZE=""-1"">"), CaseTag("</FONT>")
End Sub

Private Sub mnuHTMLFont1_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<FONT SIZE=""+1"">"), CaseTag("</FONT>")
End Sub

Private Sub mnuHTMLItalic_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<I>"), CaseTag("</I>")
End Sub

Private Sub mnuHTMLJustify_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P ALIGN=""JUSTIFY"">"), CaseTag("</P>")
End Sub

Private Sub mnuHTMLLeft_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P ALIGN=""LEFT"">"), CaseTag("</P>")
End Sub

Private Sub mnuHTMLParagraph_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CaseTag("<P>")
End Sub

Private Sub mnuHTMLRight_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P ALIGN=""RIGHT"">"), CaseTag("</P>")
End Sub

Private Sub mnuHTMLSourceCompact_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CompactFormat(frmMain.CurrentDoc.rtfHTML.SelText)
End Sub

Private Sub mnuHTMLSourceHierachal_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = HierarchalFormat(frmMain.CurrentDoc.rtfHTML.SelText)
End Sub

Private Sub mnuHTMLSourceSimple_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = SimpleFormat(frmMain.CurrentDoc.rtfHTML.SelText)
End Sub

Private Sub mnuHTMLSourceStripText_Click()
    Dim returnString As String
    ' STRIP TEXT FROM HTML FILE
    '   This will strip all the text from an HTML file
        ' Call function
        returnString = stripText(frmMain.CurrentDoc.rtfHTML.Text)
        ' We must put the returned string in a new window
    '            Call addEditorWindow
        ' Put new text in the window
        frmMain.CurrentDoc.rtfHTML.Text = returnString
End Sub

Private Sub mnuHTMLStripLinks_Click()
    Dim returnString As String
    '' STRIP LINKS FROM HTML FILE
    ''   This will strip links from an HTML file
    '    ' Return string
        ' Call function
        returnString = stripLinks(frmMain.CurrentDoc.rtfHTML.Text)
        ' Put string in a new window
    '            Call addEditorWindow
        ' Put text in
        frmMain.CurrentDoc.rtfHTML.Text = returnString
End Sub

Private Sub mnuHTMLUnderline_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<U>"), CaseTag("</U>")
End Sub

Private Sub mnuHyperlink_Click()
    frmHTMLLink.Show vbModal, Me
End Sub

Private Sub mnuINS_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<INS>"), CaseTag("</INS>")
End Sub

Private Sub mnuInsertLastUpdated_Click()
    ' Last updated
    frmMain.CurrentDoc.rtfHTML.SelText = "<script language=""JavaScript""><!--" & vbCrLf & _
                                        "document.write('<p align=""center""><i><b>This page was last updated: ')" & vbCrLf & _
                                        "document.write (document.lastModified)" & vbCrLf & "document.write(' Copyright 2000</i></b></p>')" & vbCrLf & _
                                        "// --></script>" & vbCrLf
End Sub

Private Sub mnuInsertMarquee_Click()
    frmMarquee.Show vbModal, Me
End Sub

Private Sub mnuLineBreak_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CaseTag("<BR>")
End Sub

Private Sub mnuMRU_Click(Index As Integer)
    OpenDoc (MRUFiles.item(Index + 1)), False
End Sub

Private Sub mnuNoBreak_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CaseTag("NOBR;")
End Sub

Private Sub mnuOneLineTextBox_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P><input type=""") & CaseTag("text") & """" & CaseTag(" size=") & """" & "20" & """" & CaseTag(" name=") & """" & CaseTag("T1") & """" & ">", CaseTag("</P>")
End Sub

Private Sub mnuPRE_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<PRE>"), CaseTag("</PRE>")
End Sub

Private Sub mnuRadioButton_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P><input type=""") & CaseTag("radio") & """" & CaseTag(" checked name=") & """" & CaseTag("R1") & """" & CaseTag(" value=") & """" & CaseTag("V1") & """" & ">", CaseTag("</P>")
End Sub

Private Sub mnuSAMPLE_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<SAMPLE>"), CaseTag("</SAMPLE>")
End Sub

Private Sub mnuScrollingTextBox_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<P><TEXTAREA NAME=""") & CaseTag("S1") & """" & CaseTag(" ROWS=") & """" & "2" & """" & CaseTag(" COLS=") & """" & "20" & """" & ">", CaseTag("</TEXTAREA></P>")
End Sub

Private Sub mnuSearchFind_Click()
    frmFind.Replace = False
    frmFind.Show vbModeless, Me
End Sub

Private Sub mnuSearchReplace_Click()
    frmFind.Replace = True
    frmFind.Show vbModeless, Me
End Sub

Private Sub mnuSound_Click()
    frmSound.Show vbModal, Me
End Sub

Private Sub mnuSTRONG_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<STRONG>"), CaseTag("</STRONG>")
End Sub

Private Sub mnuSUB_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<SUB>"), CaseTag("</SUB>")
End Sub

Private Sub mnuSUP_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<SUP>"), CaseTag("</SUP>")
End Sub

Private Sub mnuTB_Click(Index As Integer)
    Select Case Index
        Case 0
            UpdateToolbarDisplay
        Case 1
            UpdateStatusBarDisplay
    End Select
End Sub

Private Sub mnuTestDefault_Click()
    Dim URL$
    
    If CurrentDoc.UnSaved And Len(CurrentDoc.Filename) = 0 Then
        CurrentDoc.SaveTemp
        URL$ = CurrentDoc.TempName
    Else
        URL$ = CurrentDoc.Filename
    End If
    
    ShellExecute 0&, vbNullString, URL$, vbNullString, GetFilePart(URL, kfpDrive), vbNormalFocus
End Sub

Private Sub mnuToolsColorPicker_Click()
    frmColors.Show vbModal, Me
End Sub

Private Sub mnuToolsSpellCheck_Click()
    SpellCheck
End Sub

Private Sub mnuTT_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<TT>"), CaseTag("</TT>")
End Sub

Private Sub mnuVAR_Click()
    CurrentDoc.InsertSurroundTag CaseTag("<VAR>"), CaseTag("</VAR>")
End Sub

Private Sub mnuViewCodeClips_Click()
    frmCodeClip.Show vbModeless, Me
End Sub

Private Sub mnuViewOptions_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    On Error Resume Next
    Me.CurrentDoc.rtfHTML.Colorize
    frmMain.CurrentDoc.Refresh
End Sub

Private Sub mnuViewStatusBar_Click()
    UpdateStatusBarDisplay
End Sub

Private Sub mnuViewToolbar_Click()
    UpdateToolbarDisplay
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Public Function stripText(ByVal inputString As String) As String
' STRIP TEXT FROM HTML
'   Strip the text from HTML files

    ' Variables
    Dim outputString As String
    Dim i As Long
    Dim inTag As Boolean
    
    ' Show hourglass
    Screen.MousePointer = 11
    
    ' Walk through the text and strip non-links
    outputString = ""
    For i = 1 To Len(inputString)
        Select Case Mid(inputString, i, 1)
            Case "<"
                inTag = True
            Case ">"
                inTag = False
            Case Else
                ' Copy contents
                If inTag = False Then
                    outputString = outputString & Mid(inputString, i, 1)
                End If
        End Select
    Next i
    
    ' Copy the return string to the function
    stripText = outputString
    
    ' Show normal pointer
    Screen.MousePointer = 0

End Function

Public Function stripLinks(inputString As String) As String
' STRIP LINKS FROM HTML FILE
'   This function will strip the links from an html function

    ' Variables
    Dim i As Long
    Dim outputString As String
    Dim startPos As Long, endPos As Long

    ' Show hourglass
    Screen.MousePointer = 11
    
    Debug.Print Len(inputString)
    ' Walk through the string
    For i = 1 To Len(inputString)
        ' Check, is this the start of a link tag?
        If Mid(LCase(inputString), i, 7) = "<a href" Then
            ' Record the start
            startPos = i
        ElseIf Mid(LCase(inputString), i, 4) = "</a>" Then
            ' Record the end
            endPos = i + 4
        
            ' Record the link
            outputString = outputString + Mid(inputString, startPos, (endPos - startPos)) + vbCrLf

            ' Reset start
            startPos = 0
        End If
        Debug.Print i & " / " & Len(inputString)
    Next i
    ' Set the function equal to the return string
    stripLinks = outputString
    ' Show normal mouse pointer
    Screen.MousePointer = 0
End Function

Private Sub mnuWindowNext_Click()
    Dim i%
    
    If Documents.Count < 2 Then Exit Sub
    For i = 1 To Documents.Count
        If Documents(i).DocKey = CurrentDoc.DocKey Then
            If i = Documents.Count Then
                Documents(1).SetFocus
            Else
                Documents(i + 1).SetFocus
            End If
            
            Exit For
        End If
    Next i
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWordBreak_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = CaseTag("WBR;")
End Sub

Public Sub OpenDoc(fName$, AddMRU As Boolean)
    Screen.MousePointer = 11
    Dim i As Integer
    If Len(fName) = 0 Then
        On Error GoTo OpenError
        With frmMain.dlgFiles
            .DialogTitle = "Open file..."
            .Filter = Filter
            .Filename = ""
            .ShowOpen
            fName = .Filename
            HTMLFile = True
            Filename = dlgFiles.Filename
        End With
    End If


    Dim f As New frmDoc
    Set f = New frmDoc
    If Len(fName) > 0 And Len(Dir(fName)) = 0 Then
        ErrHandler vbObjectError, "Cannot Open File:" & fName$, "OpenDoc", , , True
        Screen.MousePointer = 0
        Exit Sub
    Else
        f.Filename = fName
        f.Show
        f.Refresh
        f.Modified = False
    End If
    
    If AddMRU Then
        If InList(MRUFiles, fName) Then
            MRUFiles.Remove fName
        End If
        
        If MRUFiles.Count > 0 Then
            MRUFiles.Add item:=fName, Before:=1, key:=fName
        Else
            MRUFiles.Add fName, fName
        End If
        UpdateMRU
        Filename = fName
    End If
    Screen.MousePointer = 0
    Exit Sub
    
OpenError:
    HTMLFile = False
    Screen.MousePointer = 0
End Sub

Private Sub sbStatusBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuTSBars
    End If
End Sub

Private Sub tmrAutosave_Timer()
    Static Minutes As Integer
    Dim doc As frmDoc, Stat$
        
    If Minutes = AutosaveInterval Then
        For Each doc In Documents
            If doc.Modified Then
                doc.SaveTemp
            End If
        Next doc
        Minutes = 0
    Else
        Minutes = Minutes + 1
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "NEW"
            LoadNewDoc
            Me.CurrentDoc.rtfHTML.SetFocus
        Case "OPEN"
            OpenDoc "", True
        Case "SAVE"
            SaveDoc
        Case "CUT"
            CutText
        Case "COPY"
            CopyText
        Case "PASTE"
            PasteText
        Case "UNDO"
            CurrentDoc.Undo
        Case "DELETE"
            CurrentDoc.rtfHTML.SelText = ""
        Case "REDO"
            CurrentDoc.Redo
        Case "WEB"
            mnuTestDefault_Click
        Case "TABLE"
            frmTable.Show vbModal
        Case "LIST"
            frmList.Show vbModal
        Case "FRAME"
            frmFrames.Show vbModal
        Case "BOLD"
            CurrentDoc.InsertSurroundTag CaseTag("<B>"), CaseTag("</B>")
        Case "ITALIC"
            CurrentDoc.InsertSurroundTag CaseTag("<I>"), CaseTag("</I>")
        Case "UNDERLINE"
            CurrentDoc.InsertSurroundTag CaseTag("<U>"), CaseTag("</U>")
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Parent.key
        Case "OPEN"
            OpenDoc ButtonMenu.key, False
        Case "TABLE"
            If ButtonMenu.Text = "TH" Then
                CurrentDoc.InsertSurroundTag CaseTag("<TH>"), CaseTag("</TH>")
            ElseIf ButtonMenu.Text = "TR" Then
                CurrentDoc.InsertSurroundTag CaseTag("<TR>"), CaseTag("</TR>")
            ElseIf ButtonMenu.Text = "TD" Then
                CurrentDoc.InsertSurroundTag CaseTag("<TD>"), CaseTag("</TD>")
            ElseIf ButtonMenu.Text = "CAPTION" Then
                CurrentDoc.InsertSurroundTag CaseTag("<CAPTION>"), CaseTag("</CAPTION>")
            ElseIf ButtonMenu.Text = "TABLE" Then
                CurrentDoc.InsertSurroundTag CaseTag("<TABLE>"), CaseTag("</TABLE>")
            End If
        Case "LIST"
            If ButtonMenu.Text = "LI" Then
                CurrentDoc.InsertSurroundTag CaseTag("<LI>"), CaseTag("</LI>")
            ElseIf ButtonMenu.Text = "DT" Then
                CurrentDoc.InsertSurroundTag CaseTag("<DT>"), CaseTag("</DT>")
            ElseIf ButtonMenu.Text = "DD" Then
                CurrentDoc.InsertSurroundTag CaseTag("<DD>"), CaseTag("</DD>")
            ElseIf ButtonMenu.Text = "MENU" Then
                CurrentDoc.InsertSurroundTag CaseTag("<MENU>"), CaseTag("</MENU>")
            End If
        Case "FRAME"
            If ButtonMenu.Text = "FRAMESET" Then
                CurrentDoc.InsertSurroundTag CaseTag("<FRAMESET>"), CaseTag("</FRAMESET>")
            ElseIf ButtonMenu.Text = "FRAME" Then
                CurrentDoc.InsertSurroundTag CaseTag("<FRAME>"), CaseTag("</FRAME>")
            ElseIf ButtonMenu.Text = "NOFRAME" Then
                CurrentDoc.InsertSurroundTag CaseTag("<NOFRAME>"), CaseTag("</NOFRAME>")
            End If
    End Select
End Sub

Private Sub UpdateToolbarDisplay()
    ' Toggle Toolbar
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked = True Then
        pSetIcon "PUSHPIN2", "mnuViewToolbar"
        mnuTB(0).Checked = True
    Else
        pSetIcon "PUSHPIN", "mnuViewToolbar"
        mnuTB(0).Checked = False
    End If
    Toolbar1.Visible = mnuViewToolbar.Checked Or mnuTB(0).Checked
End Sub

Private Sub UpdateStatusBarDisplay()
    '  Toggle Statusbar
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    If mnuViewStatusBar.Checked = True Then
        pSetIcon "PUSHPIN2", "mnuViewStatusBar"
        mnuTB(1).Checked = True
    Else
        pSetIcon "PUSHPIN", "mnuViewStatusBar"
        mnuTB(1).Checked = False
    End If
    sbStatusBar.Visible = mnuViewStatusBar.Checked Or mnuTB(1).Checked
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuTSBars
    End If
End Sub

Sub SaveFileAs()
    Dim f$, OldFile$, Done As Boolean
    On Error GoTo Canceled

    ' Display dialog until user chooses a unique filename, decides to overwrite
    ' an existing one, or cancels the save.
    '
    Do
        With dlgFiles
            .DialogTitle = "Save Document...?"
            .Flags = cdlOFNHideReadOnly
            .Filter = Filter
            .Filename = ""
'            .InitDir = LastSaveDir
            .ShowSave
            f$ = .Filename
        End With

        If LCase$(f$) Like "*htm" Then
            f$ = f$ & "l"           ' we like the full html extension
        End If

        If Len(Dir(f$)) > 0 Then
            Dim Result As VbMsgBoxResult
            Result = MsgBox(f$ & vbCrLf & "This file already exists.  Do you want to replace it?", _
                vbQuestion Or vbYesNo, "Overwrite file?")
            Done = IIf(Result = vbYes, True, False)
        Else
            Done = True
        End If
    Loop Until Done
'    LastSaveDir = GetFilePart(f$, kfpDrivePath)
    With CurrentDoc
        If Right$(dlgFiles.Filename, 4) = ".rtf" Then
            OldFile = .Filename
            .rtfHTML.SaveFile f$
'            .rtfHtml.SaveFile f$, 0   ' save it as RTF
            .UnSaved = False        ' now an existing doc
            .SetModified False      ' unmodified since last save
            .Filename = f$          ' new filename...
            .Caption = f$           ' ...and a matching caption
    
            ' If the doc has been saved in a temp file,
            ' kill the temp file and remove it from the
            ' temp file list, since we don't need it any more
'            If Len(.Tempname) > 0 Then
'                TempFiles.Remove .Tempname
'                FileDumpCol TempFiles, AppPath & "tempdoc.his"
'                Kill .Tempname
'                .Tempname = ""
'            End If
'        ElseIf Right$(dlgFiles.FileName, 4) = ".lsp" And .AutoLisp = True Then
'            OldFile = .FileName
'            .txtLSP.SaveFile f$, rtfhtml
'            .UnSaved = False
'            .SetModified False
'            .FileName = f$
'            .Caption = f$
'            ' If the doc has been saved in a temp file,
'            ' kill the temp file and remove it from the
'            ' temp file list, since we don't need it any more
'            If Len(.Tempname) > 0 Then
'                TempFiles.Remove .Tempname
'                FileDumpCol TempFiles, AppPath & "tempdoc.his"
'                Kill .Tempname
'                .Tempname = ""
'            End If
        ElseIf OldFile = .Filename Then
            .rtfHTML.SaveFile f$
'            .rtfHtml.SaveFile f$, 1  ' save it as plain text
            .UnSaved = False        ' now an existing doc
            .SetModified False      ' unmodified since last save
            .Filename = f$          ' new filename...
            .Caption = f$           ' ...and a matching caption
    
            ' If the doc has been saved in a temp file,
            ' kill the temp file and remove it from the
            ' temp file list, since we don't need it any more
'            If Len(.Tempname) > 0 Then
'                TempFiles.Remove .Tempname
'                FileDumpCol TempFiles, AppPath & "tempdoc.his"
'                Kill .Tempname
'                .Tempname = ""
'            End If
        End If
    End With
    ' Add it to the Most Recently Used files list
    If Not InList(MRUFiles, f$) Then
        MRUFiles.Add item:=f$, Before:=1
        UpdateMRU
    End If
Canceled:
End Sub

' Update toolbar buttons, menu items for undo/redo
Public Sub UpdateUndoFunctions()
    With CurrentDoc
        mnuEditUndo.Enabled = .UndoStack.Count > 1
        mnuEditRedo.Enabled = .RedoStack.Count > 0
        Toolbar1.Buttons("UNDO").Enabled = .UndoStack.Count > 1
        Toolbar1.Buttons("REDO").Enabled = .RedoStack.Count > 0
    End With
End Sub

Public Sub PrintDoc()
    Dim T As String
    Dim H As Integer
    Dim l As Integer
    Dim P As Integer
    Dim ct As String
    Dim i As Integer
    Dim J As Long
    Dim tt As String
    Dim garniture As Boolean
    Dim page As Integer
'   On Error GoTo NotepadError
    TitreVar = "Printing"
    MsgVar = "Printer: " + Printer.DeviceName + LfVar
    MsgVar = MsgVar + " over " + Printer.Port + LfVar + LfVar
    MsgVar = MsgVar + "Would you like to print... ?" + LfVar
    MsgVar = MsgVar + frmMain.ActiveForm.Caption
    RepVar = MsgBox(MsgVar, 33, TitreVar)
    Select Case RepVar
    Case 1  ' Ok
    ' rien à faire
    Case 2  ' Annuler
    Exit Sub
    End Select
    ' retrouver les paramètres de l'imprimante
    MousePointer = 11
    Printer.ScaleMode = vbInches
    H = Printer.ScaleHeight
    l = Printer.ScaleWidth
    
    ' initialise le printer
    Printer.Print " ";
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    page = 1
    
    ' Si imprime la garniture
    If GetSetting(App.Title, "Settings", "PositionPrefere 0", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = 0.25
    Printer.Print frmMain.ActiveForm.Caption
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 1", 1) = 1 Then
    Printer.Line (0.5, 0.5)-(l - 1, 0.5)
    Printer.Line (0.05, H - 1)-(l - 1, H - 1)
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 2", 1) = 1 Then
    Printer.CurrentX = 0.5     ' was 1
    Printer.CurrentY = H - 0.75
    Printer.Print "Page " + str(page);
    garniture = True
    End If
    
    ' replace le printer au debut
    Printer.CurrentX = 0.5
    If garniture = True Then
    Printer.CurrentY = 0.65
    Else
    Printer.CurrentY = 0.25
    End If
    
    ' Impression du texte
    T = CurrentDoc.rtfHTML.Text
    J = Len(T)
    P = l - 2   ' largeur de la page -2 po.
    ' Sélectionne les caractères d'impression
    Printer.FontBold = GetSetting(App.Title, "Settings", "FontBold", 0)
    Printer.FontItalic = GetSetting(App.Title, "Settings", "Italic", 0)
    Printer.FontName = GetSetting(App.Title, "Settings", "FontName", "Arial")
    Printer.FontSize = GetSetting(App.Title, "Settings", "FontSize", 10)
    Do
    i = i + 1
    ct = ct + Mid(T, i, 1)  ' chaine temporaire
    ' si la page est pleine
    If Printer.CurrentY > H - 1.25 Then     ' was 1.25
    Printer.NewPage
    page = page + 1
    Printer.Print " ";
    
    ' Si imprime la garniture
    If GetSetting(App.Title, "Settings", "PositionPrefere 0", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = 0.25
    Printer.Print frmMain.ActiveForm.Caption
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 1", 1) = 1 Then
    Printer.Line (0.5, 0.5)-(l - 1, 0.5)
    Printer.Line (0.5, H - 1)-(l - 1, H - 1)
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 2", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = H - 0.75
    Printer.Print "Page " + str(page);
    garniture = True
    End If
    
    ' replace le printer au debut
    Printer.CurrentX = 0.5
    If garniture = True Then
    Printer.CurrentY = 0.65
    Else
    Printer.CurrentY = 0.25
    End If
    
    Printer.FontBold = GetSetting(App.Title, "Settings", "FontBold", 0)
    Printer.FontItalic = GetSetting(App.Title, "Settings", "Italic", 0)
    Printer.FontName = GetSetting(App.Title, "Settings", "FontName", "Arial")
    Printer.FontSize = GetSetting(App.Title, "Settings", "FontSize", 10)
    End If
    
    If i >= J Then ' si la fin du texte
    'Debug.Print ct
    Printer.CurrentX = 0.5
    ' vide le buffer
    Printer.Print ct;
    ' pi sort de là
    Exit Do
    End If
    If Printer.TextWidth(ct) >= P Then
    ' wrap
    ' couper à un espace
    tt = ct
    Do
    ' recule dans tt (text temporaire)
    ' jusqu'au premier espace
    If Right(tt, 1) = Chr(32) Then
    ' conserve dans ct le reste
    ct = Mid(ct, Len(tt) + 1)
    Exit Do
    End If
    tt = Left(tt, Len(tt) - 1)
    Loop
    'Debug.Print tt
    Printer.CurrentX = 0.5
    Printer.Print tt
    tt = ""
    End If
    If Mid(T, i, 1) = Chr(13) Or Mid(T, i, 1) = Chr(10) Then
    ' si je frappe un carriage return
    ct = Left(ct, Len(ct) - 1)
    'Debug.Print ct
    Printer.CurrentX = 0.5
    Printer.Print ct
    ct = ""
    i = i + 1   ' sauter le chr(10)
    End If
    Loop
    Printer.EndDoc
    MousePointer = 1
    On Error GoTo 0
End Sub

Private Sub SpellCheck()
'Purpose:   Demo how to use another application's services with in our
'           program to add some sort of functionality.  In this case spell checking
'           Very simple example. A lot of functionality was left out on purpose.
'REQUIRES:  Word97 or 2000 Reference
'Basic procedure
'           Open word
'           Copy text to a new document
'           spell check each word one at a time
'               if it fails then display a list of suggestions for the user to select from
'               if the user indicates a replacement then replace the word with the replacement
'           when all words are checked copy the text back to the text box
'           Close word

    Dim wApp As Word.Application    'Object for word application
    Dim doc As Word.Document        'Object for word document
    Dim wd As Word.Words            'object for a collection of words in the document
    Dim wSuggList As Word.SpellingSuggestions 'object for a collection of spelling suggestions (result of a method)
    Dim ss As Word.SpellingSuggestion 'object for one speeling suggestion in the above collection of spelling suggestions
    Dim bPassCheck As Boolean 'Spell Check Results
    Dim sMsg As String 'Holds text to be checked or changed
    Dim i As Integer 'counter
    
    Set wApp = New Word.Application 'Open word
    
    'Add a new document then
    'Copy the all or the selected text from the active textbox to the new document
    'I are using the InsertAfter method to add the text
    'I could have also used the InsertBefore method. In this case it does not matter
    Set doc = wApp.Documents.Add
    If frmMain.CurrentDoc.rtfHTML.Visible = True Then
        If frmMain.CurrentDoc.rtfHTML.SelLength = 0 Then
            doc.Range.InsertAfter frmMain.CurrentDoc.rtfHTML.Text
        Else
            doc.Range.InsertAfter frmMain.CurrentDoc.rtfHTML.SelText
        End If
    Else
        If frmMain.CurrentDoc.rtfHTML.SelLength = 0 Then
            doc.Range.InsertAfter frmMain.CurrentDoc.rtfHTML.Text
        Else
            doc.Range.InsertAfter frmMain.CurrentDoc.rtfHTML.SelText
        End If
    End If
    
    'Create a collection of all the words on the new document
    Set wd = doc.Words
    
    'loop through all words in the list
    'Performing a spell check on each word one at a time
    i = 0
    Do
        i = i + 1
        'Perform Spell Check and store results in bPassCheck
        bPassCheck = wApp.CheckSpelling(wd(i))
        If bPassCheck = False Then 'False the spell check failed
            Set wSuggList = wApp.GetSpellingSuggestions(wd(i))  'get a list of suggestions
            Load frmSpellCheck 'load the formn that displays the list of suggestions (ignored if already loaded)
            frmSpellCheck.txtWord.Text = wd(i)  'Add the bad word
            frmSpellCheck.lstWords.Clear        'clear any existing suggestions
            If wSuggList.Count <> 0 Then        'check to see if there are any suggestions
                For Each ss In wSuggList        'Add the new suggestions from the collection
                    frmSpellCheck.lstWords.AddItem ss.Name
                Next
                frmSpellCheck.lstWords.ListIndex = 0    'Select the first item in the list of suggestions
                frmSpellCheck.txtReplaceWith.Text = frmSpellCheck.lstWords.List(frmSpellCheck.lstWords.ListIndex) 'display the text also
            Else
                frmSpellCheck.txtReplaceWith.Text = "" 'No suggestions
            End If
            frmSpellCheck.Show vbModal 'display the spell check form
            'when the user selects ignore, replace, or cancel the form is hidden not unloaded
            'perform indicated action using the properties of the form
            If frmSpellCheck.bCancelCheck = True Then Exit Do
            If frmSpellCheck.bReplaceWord = True Then
                wd(i) = frmSpellCheck.txtReplaceWith.Text & " " 'Add a space as new suggestions don't have the space
            End If
        End If
    Loop Until i = wd.Count 'Loop until there are no more words in the collection
    
    'Copy the text back to the correct textbox on the activeform
    If frmMain.CurrentDoc.rtfHTML.Visible = True Then
        If frmMain.CurrentDoc.rtfHTML.SelLength = 0 Then
            sMsg = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf) 'Word only uses CR's for hard breaks.  VB needs CR and LF
            frmMain.CurrentDoc.rtfHTML.Text = sMsg
        Else
            frmMain.CurrentDoc.rtfHTML.SelText = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf)
            frmMain.CurrentDoc.rtfHTML.SelText = sMsg
        End If
    Else
        If frmMain.CurrentDoc.rtfHTML.SelLength = 0 Then
            sMsg = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf) 'Word only uses CR's for hard breaks.  VB needs CR and LF
            frmMain.CurrentDoc.rtfHTML.Text = sMsg
        Else
            frmMain.CurrentDoc.rtfHTML.SelText = doc.Range.Text
            sMsg = Replace(sMsg, Chr(13), vbCrLf)
            frmMain.CurrentDoc.rtfHTML.SelText = sMsg
        End If
    End If
    MsgBox "Spell Check has been completed" & vbCrLf & "Spell Checked " & CStr(wd.Count) & " words", vbInformation
    
    'Clean up
    doc.Close False  'close the document indicating that we should not save anything
    wApp.Quit False  'close word application
End Sub

Public Sub PlaceCursor(Text$, Cursor As Long)
    Dim T As Long
        
    T = CurrentDoc.rtfHTML.SelStart
    CurrentDoc.rtfHTML.SelStart = (T + Len(Tag)) - Cursor
End Sub
