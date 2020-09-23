VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About AutoLisp Colorizer"
   ClientHeight    =   4215
   ClientLeft      =   9240
   ClientTop       =   2985
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2909.268
   ScaleMode       =   0  'User
   ScaleWidth      =   5282.165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSpecialCopyright 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmAbout.frx":1272
      Top             =   2760
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label lblDotLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "................"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   255
      Width           =   1920
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      Height          =   675
      Index           =   0
      Left            =   4500
      Top             =   540
      Width           =   615
   End
   Begin VB.Label lblDotLine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   2160
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000 Matt Wunch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "frmAbout.frx":14C9
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Tag             =   "http://vbaccelerator.com/j-index.htm?url=cright.htm"
      Top             =   3540
      Width           =   3990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   675
      Index           =   3
      Left            =   1740
      TabIndex        =   6
      Top             =   1500
      Width           =   2100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   675
      Index           =   0
      Left            =   2220
      TabIndex        =   5
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTERED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   435
      Left            =   1320
      TabIndex        =   2
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   660
      Width           =   1575
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   16
      Height          =   855
      Index           =   1
      Left            =   180
      Top             =   180
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":17D3
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1014.176
      X2              =   5310.337
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080C0FF&
      Height          =   225
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   4860
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmAbout"
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


Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim txtCopyright As String
    txtCopyright = "The use of this package indicates your understanding and acceptance of the following terms and conditions. This license shall supersede any verbal, or prior verbal or written, statement or agreement to the contrary. If you do not understand or accept these terms, or your local regulations prohibit 'after sale' license agreements or limited disclaimers, you must cease and desist using this product immediately." & vbCrLf & vbCrLf
    txtCopyright = txtCopyright & "Liability disclaimer:" & vbCrLf & "This product and/or license is provided as is, without any representation or warranty of any kind, either express or implied, including without limitation any representations or endorsements regarding the use of, the results of, or performance of the product, its appropriateness, accuracy, reliability, or correctness. The user and/or licensee assume the entire risk as to the use of this product.  The author will not not assume liability for the use of this program beyond the original purchase price of the software. In no event will the author be liable for additional direct or indirect damages including any lost profits, lost savings, or other incidental or consequential damages arising from any defects, or the use or inability to use these programs, even if the author has been advised of the possibility of such damages." & vbCrLf & vbCrLf
    txtCopyright = txtCopyright & "Terms:" & vbCrLf & "This license is effective until terminated. You may terminate it by destroying the program, the documentation and copies thereof. This license will also terminate if you fail to comply with any terms or conditions of this agreement. You agree upon such termination to destroy all copies of the program and of the documentation, or return them to the author for disposal." & vbCrLf & vbCrLf
    txtCopyright = txtCopyright & "Other Rights And Restrictions:" & vbCrLf & "All other rights and restrictions not specifically granted in this license are reserved by the author." & vbCrLf & vbCrLf
    txtCopyright = txtCopyright & "Thank you for your interest in this product."
    lblInfo(2).Caption = App.Major & "." & App.Minor & "." & App.Revision
    
    txtSpecialCopyright = txtCopyright
End Sub

Private Sub lblCopyright_Click()
    OpenURL "mailto:mydixiewrecked2@hotmail.com?subject=HTML2K?body=In regards to your HTML2K program..."
End Sub
