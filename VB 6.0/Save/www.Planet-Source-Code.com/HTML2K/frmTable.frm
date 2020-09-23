VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table Wizard"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   2655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtBorder"
         BuddyDispid     =   196613
         OrigLeft        =   1920
         OrigTop         =   2160
         OrigRight       =   2160
         OrigBottom      =   2415
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtBorder 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   12
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtSpacing 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtPadding 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCols 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtRows 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Caption:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Border:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cell Spacing:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cell Padding:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Columns:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rows:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Label lblErr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1245
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmTable"
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

Dim TableCount As String
Dim TempText As String
Dim T$
Dim CastTag
Dim C As Integer
Dim R As Integer
Dim i As Integer
Dim J As Integer
Dim TableStart As String
Dim TableEnd As String
Dim Header As String
Dim Footer As String
Dim Numbers As Integer


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    On Error Resume Next
    R = CInt(txtRows.Text)
    C = CInt(txtCols.Text)

    Header = "<!-- This Table was created by HTML2K -->" & vbCrLf & _
             "<!-- Created on: " & Now & " -->" & vbCrLf
    Footer = vbCrLf & "<!-- End Table -->"

    TableStart = "<TABLE border=""" & txtBorder.Text & """ width=""" & txtWidth.Text & "%"" cellspacing=""" & txtSpacing.Text & """" & " cellpadding=""" & txtPadding.Text & """>" & vbCrLf & _
                 "<CAPTION>" & txtCaption.Text & "</CAPTION>"
    TableEnd = "</TABLE>"

    For i = 1 To R
        T$ = T$ & "  <!-- Row " & CStr(i) & "-->" & vbCrLf & _
            CaseTag("  <TR>") & vbCrLf
        For J = 1 To C
            T$ = T$ & CaseTag("    <TD></TD>") & vbCrLf
        Next J
        T$ = T$ & CaseTag("  </TR>") & vbCrLf
    Next i
    
    With frmMain.CurrentDoc.rtfHTML
        .SelText = Header
        .SelText = TableStart & vbCrLf
        .SelText = T$
        .SelText = TableEnd
        .SelText = Footer
    End With
    
    Header = ""
    TableStart = ""
    T$ = ""
    TableEnd = ""
    Footer = ""
    
    Unload Me
End Sub

Public Sub NumbersOnly()
    lblErr.Visible = True
    lblErr.Caption = "Only numbers are allowed."
End Sub

Private Sub txtBorder_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtCols_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtPadding_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtRows_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpacing_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    Numbers = KeyAscii
    lblErr.Visible = False
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        NumbersOnly
        KeyAscii = 0
    End If
End Sub
