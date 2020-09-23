VERSION 5.00
Begin VB.Form frmNewCodeClip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Code Clip Name"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmNewCodeClip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmNewCodeClip"
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

Public blnCanceled As Boolean
Public Text As String


Private Sub Combo1_Change()
    Static T$
    
    If Len(Combo1.Text) > Len(T$) Then      ' user did not delete text
        T$ = Combo1.Text
        AutoComplete Combo1
    Else
        T$ = Combo1.Text
    End If
    Text = Combo1.Text
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            blnCanceled = False
            Unload Me
        Case 1
            blnCanceled = True
            Unload Me
    End Select
End Sub
