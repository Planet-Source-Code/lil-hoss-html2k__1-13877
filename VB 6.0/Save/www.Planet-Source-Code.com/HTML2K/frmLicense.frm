VERSION 5.00
Begin VB.Form frmLicense 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "License Information"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Organization:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmLicense"
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


Private Sub Command1_Click()
    Dim T As String
    If Text1(0).Text = "" And Text1(1).Text = "" Then
       MsgBox "Please enter your name and organization.", 16, "Message"
       Text1(0).SetFocus
       Exit Sub
    End If
    SaveSetting ThisApp, "Settings", "User Name", Text1(0).Text
    SaveSetting ThisApp, "Settings", "Organization", Text1(1).Text
    SaveSetting ThisApp, "Settings", "License", "License"
    Unload frmLicense
    Load frmMain
    frmMain.Show
End Sub
