VERSION 5.00
Begin VB.Form frmHTMLLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Link"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmHTMLLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Link"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox cboLink 
         Height          =   315
         ItemData        =   "frmHTMLLink.frx":01CA
         Left            =   1320
         List            =   "frmHTMLLink.frx":01CC
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboTarget 
         Height          =   315
         ItemData        =   "frmHTMLLink.frx":01CE
         Left            =   3120
         List            =   "frmHTMLLink.frx":01D0
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblHyperlink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hyperlink Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label lblTarget 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrption:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmHTMLLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = "<a href=""" & cboLink.Text & txtURL.Text & """ target=""" & cboTarget.Text & """ >" & txtDescription.Text & "</a>" & vbCr
    Unload Me
End Sub

Private Sub Form_Load()
    RemoveMenus Me
    With cboLink
        .AddItem "file://"
        .AddItem "ftp://"
        .AddItem "gopher://"
        .AddItem "http://"
        .AddItem "https://"
        .AddItem "mailto:"
        .AddItem "news:"
        .AddItem "telnet:"
        .AddItem "wais:"
    End With
    cboLink.ListIndex = 3
    
    With cboTarget
        .AddItem "_Default"
        .AddItem "_Self"
        .AddItem "Top"
    End With
    cboTarget.ListIndex = 0
End Sub
