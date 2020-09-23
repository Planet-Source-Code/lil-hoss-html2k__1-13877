VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHtmlHead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Head"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmHtmlHead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   1695
      TabIndex        =   10
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtHeadLevel"
      BuddyDispid     =   196609
      OrigLeft        =   1680
      OrigTop         =   360
      OrigRight       =   1920
      OrigBottom      =   615
      Max             =   6
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtHeadLevel 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alignment"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3495
      Begin VB.OptionButton optAlignment 
         Caption         =   "&Right"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optAlignment 
         Caption         =   "&Center"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optAlignment 
         Caption         =   "&Left"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtHead 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Head Level:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   360
   End
End
Attribute VB_Name = "frmHtmlHead"
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


Private Sub cmdCommand_Click(Index As Integer)
    Select Case Index
        Case 0
        
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    RemoveMenus Me
End Sub

Private Function EnableOK()
    cmdCommand(0).Enabled = True
End Function

Private Sub optAlignment_Click(Index As Integer)
    Select Case Index
        Case 0
            EnableOK
        Case 1
            EnableOK
        Case 2
            EnableOK
    End Select
End Sub

Private Sub txtHead_Change()
    EnableOK
End Sub

Private Sub txtHeadLevel_Change()
    EnableOK
End Sub
