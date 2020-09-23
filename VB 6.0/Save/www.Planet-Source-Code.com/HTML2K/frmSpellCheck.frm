VERSION 5.00
Begin VB.Form frmSpellCheck 
   Caption         =   "Spell Check"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   Icon            =   "frmSpellCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstWords 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtReplaceWith 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblReplaceWith 
      Caption         =   "Replace With"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblWord 
      Caption         =   "Word to Replace"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSpellCheck"
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

Public bCancelCheck As Boolean
Public bReplaceWord As Boolean


Private Sub cmdCancel_Click()
    bCancelCheck = True
    bReplaceWord = False
    Me.Hide
End Sub


Private Sub cmdIgnore_Click()
    bCancelCheck = False
    bReplaceWord = False
    Me.Hide
End Sub

Private Sub cmdReplace_Click()
    bCancelCheck = False
    bReplaceWord = True
    Me.Hide
End Sub


Private Sub lstWords_Click()
    txtReplaceWith.Text = lstWords.List(lstWords.ListIndex)
End Sub


Private Sub lstWords_DblClick()
    txtReplaceWith.Text = lstWords.List(lstWords.ListIndex)
    cmdReplace_Click
End Sub


