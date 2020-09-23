VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comment"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   4575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "VBScript"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "JavaScript"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comment text..."
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   435
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmComment"
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

Dim CommentStyle As String


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMain.CurrentDoc.rtfHTML.SelText = CommentStyle & txtComment.Text & vbCrLf
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    CommentStyle = "'"
End Sub

Private Sub Option1_Click()
    CommentStyle = "//"
End Sub

Private Sub Option2_Click()
    CommentStyle = "'"
End Sub
