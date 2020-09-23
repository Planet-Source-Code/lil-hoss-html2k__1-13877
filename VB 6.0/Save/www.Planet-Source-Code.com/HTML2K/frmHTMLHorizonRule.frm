VERSION 5.00
Begin VB.Form frmHTMLHorizonRule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horizon Rule"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmHTMLHorizonRule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Noshade"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmHTMLHorizonRule"
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

Dim OptText As String
Dim chkTxt As String


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        chkTxt = "NOSHADE"
    Else
        chkTxt = ""
    End If
    If Check1.Value = 0 Then
        chkTxt = ""
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMain.CurrentDoc.rtfHTML.SelText = "<HR SIZE=""" & Text1(0).Text & """ ALIGN= """ & OptText & """ WIDTH=""" & Text1(1).Text & "%" & """ " & chkTxt & ">"
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            OptText = "LEFT"
        Case 1
            OptText = "CENTER"
        Case 2
            OptText = "RIGHT"
    End Select
End Sub
