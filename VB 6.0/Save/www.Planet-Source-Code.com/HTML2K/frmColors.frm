VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Color"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2175
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "&Custom..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HTML Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1480
      Width           =   900
   End
End
Attribute VB_Name = "frmColors"
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCustom_Click()
    On Error GoTo Canceled
    With dlgColors
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .Color = Picture1.BackColor
        .ShowColor
        
        ' Update the text box. This in turn will trigger
        ' an update of the color box.
        Dim R%, G%, B%
        ExtractRGB .Color, R, G, B
        Text1.Text = "#" & Format(Hex(R), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")
        Picture1.BackColor = .Color
    End With
Canceled:
End Sub

Private Sub cmdInsert_Click()
    frmMain.CurrentDoc.rtfHTML.SelText = Right$(Text1.Text, 6)
    Unload Me
End Sub

Private Sub Form_Load()
    RemoveMenus Me
'    btnFlat cmdCancel
'    btnFlat cmdInsert
'    btnFlat cmdCustom
End Sub
