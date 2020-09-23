VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Sound"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmSound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraSound 
      Caption         =   "Select Sound File"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5130
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   270
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Wav Files"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   7
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Midi Files"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   2400
         Pattern         =   "*.wav"
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "Text1"
         BuddyDispid     =   196617
         OrigLeft        =   5535
         OrigTop         =   480
         OrigRight       =   5775
         OrigBottom      =   735
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loop"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   615
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = "<bgsound src=""" & File1.Filename & """ loop=""" & Text1.Text & """>" & Chr(13) & Chr(10)
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveErrs
        Dir1.Path = Drive1.Drive
        Exit Sub
        
DriveErrs:
    Select Case Err
        Case 71
            MsgBox prompt:="Drive not ready. Please insert disk in drive.", _
            Buttons:=vbExclamation
            ' Reset path to previous drive.
            Drive1.Drive = Dir1.Path
            Exit Sub
        Case 68
            MsgBox prompt:="Drive not available.", Title:="Drive Error", Buttons:=vbExclamation
            Drive1.Drive = Dir1.Path
        Case Else
            MsgBox prompt:="Application error.", Buttons:=vbExclamation
    End Select
End Sub

Private Sub File1_Click()
    Dir1.Path = File1.Path
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
        File1.Pattern = "*.mid"
    Else
        File1.Pattern = "*.wav"
    End If
    
'    If Option1(1).Value = True Then
'        File1.Pattern = "*.wav"
'    Else
'        File1.Pattern = "*.mid"
'    End If
End Sub
