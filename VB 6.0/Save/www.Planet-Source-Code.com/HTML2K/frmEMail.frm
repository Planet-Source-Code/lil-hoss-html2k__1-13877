VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-mail"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtSubject 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   4095
      End
      Begin VB.CheckBox chkSubject 
         Caption         =   "Include Subject <NO>"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtEMail 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description..."
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail Address..."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmEMail"
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


Private Sub chkSubject_Click()
    If chkSubject.Value = 1 Then
        chkSubject.Caption = "Include Subject <YES>"
        txtSubject.Enabled = True
        txtSubject.BackColor = &H80000005
    Else
        chkSubject.Caption = "Include Subject <NO>"
        txtSubject.Enabled = False
        txtSubject.BackColor = &H8000000B
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim EMailStart As String
    Dim EMailEnd As String
    Dim EMailSubject As String
    Select Case Index
        Case 0
            If chkSubject.Value = 1 Then
                EMailSubject = "?subject=" & txtSubject.Text
            End If
            EMailStart = "<A HREF=""mailto:"
            EMailEnd = "</A>"
            frmMain.CurrentDoc.rtfHTML.SelText = EMailStart & txtEMail.Text & EMailSubject & """>" & txtDesc.Text & EMailEnd
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub
