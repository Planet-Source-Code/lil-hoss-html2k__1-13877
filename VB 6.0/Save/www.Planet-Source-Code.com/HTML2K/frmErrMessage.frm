VERSION 5.00
Begin VB.Form ErrMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   3030
   ClientLeft      =   30
   ClientTop       =   240
   ClientWidth     =   5745
   Icon            =   "frmErrMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmErrMessage.frx":000C
      Top             =   1440
      Width           =   5535
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmErrMessage.frx":0012
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
   Begin VB.Image imgImage 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmErrMessage.frx":0022
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact for HTML2K Support"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1560
      MouseIcon       =   "frmErrMessage.frx":0BE4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "ErrMessage"
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

Public Result As ShowYesNoResult


Private Sub cmdDetails_Click()
    ErrMessage.Height = 3375
    cmdDetails.Enabled = False
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    Hide
    End '// ensure close
End Sub

Public Sub SetForm(bYesNo As Boolean)
    imgImage(0).Visible = Not bYesNo
    cmdOK.Visible = Not bYesNo
    cmdDetails.Visible = Not bYesNo
   cmdExit.Visible = Not bYesNo
    lblEmail.Visible = Not bYesNo
    cmdOK.Default = Not bYesNo
    If (bYesNo) Then
        Caption = "Question"
        Height = 1785
    Else
        Height = 1785
        Caption = "Alert"
        cmdDetails.Enabled = True
        cmdOK.TabIndex = 0
    End If
End Sub

Private Sub Form_Load()
    RemoveMenus Me
End Sub

Private Sub lblEmail_Click()
    OpenURL "mailto:mydixiewrecked2@hotmail.com"
End Sub

Public Function ChangeErrMsg(lErrorNum As Long, sErrorText As String, Optional sErrorSource) As String
    Select Case lErrorNum
    Case 429
        Select Case sErrorSource
        Case "Main.LoadMenus", "Main.LoadTools", "Main.SetupMain"
            ChangeErrMsg = "One of the following files is not registered, missing or corrupt: SSubTmr.dll, CNewMenu.dll. Please run Setup again."
        Case Else
            ChangeErrMsg = sErrorText
        End Select
    Case Else
        ChangeErrMsg = sErrorText
    End Select
End Function

Public Sub WritetoLog(sText As String)
    Dim filenum As Integer
    filenum = FreeFile
    Open App.Path & "\Errorlog.txt" For Append As filenum
    Print #filenum, sText ' & vbCrLf ' GetComponentVersions(False) & vbCrLf
    Close #filenum
End Sub
