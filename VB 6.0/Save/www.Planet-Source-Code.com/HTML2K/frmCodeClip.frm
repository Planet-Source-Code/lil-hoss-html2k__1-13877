VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCodeClip 
   Caption         =   "Code Clips"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmCodeClip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New Code Entry"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EDIT"
            Object.ToolTipText     =   "Edit Code Entry"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste Code Entry"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DELETE"
            Object.ToolTipText     =   "Delete Code Entry"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      Begin HTML2K.SplitPanel ctlSplit 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
         _extentx        =   7858
         _extenty        =   4260
         Begin VB.ListBox List1 
            Height          =   1620
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   1005
            Left            =   1320
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   0
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmCodeClip"
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


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMain.CurrentDoc.rtfHTML.SelText = Text1.Text
            SaveCodeClips
        Case 1
            SaveCodeClips
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Set ctlSplit.Control1 = List1
    Set ctlSplit.Control2 = Text1
    LoadCodeClips
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "NEW"
            NewClip
        Case "EDIT"
            EditClip
        Case "PASTE"
            PasteClip
        Case "DELETE"
            DeleteClip
    End Select
End Sub

Private Sub NewClip()
    Dim Clip As New clsCodeClip
    With frmNewCodeClip
'        .Prompt = "Enter Clip Name:"
'        .SetList Nothing
'        .Top = Me.Top + (Me.Height / 2)
'        .Left = Me.Left + (Me.Width / 2)
        .Show vbModal, Me
        
        If .blnCanceled Then
            Exit Sub
        End If
    End With
        
    With Clip
        .Name = frmNewCodeClip.Text
        CodeClips.Add Clip, .Name
        List1.AddItem .Name
        List1.ListIndex = List1.ListCount - 1
    End With
    
    Text1.Locked = False
    SaveCodeClips
End Sub

Private Sub EditClip()
    MsgBox "Put some code here!", , "EditClip"
End Sub

Private Sub PasteClip()
    MsgBox "Put some code here!", , "PasteClip"
End Sub

Private Sub DeleteClip()
    MsgBox "Put some code here!", , "DeleteClip"
End Sub
