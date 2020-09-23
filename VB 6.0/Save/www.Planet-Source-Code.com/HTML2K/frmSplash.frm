VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      TabIndex        =   5
      Top             =   300
      Width           =   2055
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
      Left            =   3540
      TabIndex        =   0
      Top             =   255
      Width           =   1920
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      Height          =   675
      Index           =   0
      Left            =   4560
      Top             =   540
      Width           =   615
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   4920
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   420
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   16
      Height          =   855
      Index           =   1
      Left            =   240
      Top             =   180
      Width           =   855
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
      TabIndex        =   6
      Top             =   660
      Width           =   1575
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   1440
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1500
      Width           =   2100
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
      Left            =   3060
      TabIndex        =   2
      Top             =   2040
      Width           =   225
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
      Left            =   120
      TabIndex        =   1
      Top             =   870
      Width           =   2160
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1140
      TabIndex        =   7
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080C0FF&
      Height          =   225
      Index           =   4
      Left            =   1140
      TabIndex        =   8
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmSplash"
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


Private Sub Form_Load()
    lblInfo(2).Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub
