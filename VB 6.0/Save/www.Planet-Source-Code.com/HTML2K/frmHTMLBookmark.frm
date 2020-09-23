VERSION 5.00
Begin VB.Form frmHTMLBookmark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Bookmark"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmHTMLBookmark.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtBookmark 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblBookmark 
      AutoSize        =   -1  'True
      Caption         =   "Bookmark Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "frmHTMLBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = "<p><a name=""" & txtBookmark.Text & """></a></p>"
    Unload Me
End Sub

Private Sub Form_Load()
    RemoveMenus Me
End Sub

