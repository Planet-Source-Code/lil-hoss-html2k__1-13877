VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMarquee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Marquee"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmMarquee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   1080
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   6255
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VSpace"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   23
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "HSpace"
         Height          =   195
         Index           =   7
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Delay"
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Amount"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6255
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   5520
         Picture         =   "frmMarquee.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1280
         Width           =   375
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   4320
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Index           =   9
         Left            =   4320
         TabIndex        =   25
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Direction"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   19
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Behavior"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Align"
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scrolling Text"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6255
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmMarquee"
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
            If Text5.Text = "" Then
                MsgBox "You did not type any text to scroll!"
                Exit Sub
            Else
                frmMain.CurrentDoc.rtfHTML.SelText = "<MARQUEE ALIGN=""" & Combo2.Text & """" & " HEIGHT=""" & Combo1(0).Text & """" & " WIDTH=""" & Combo1(1).Text & """" & " SCROLLAMOUNT=""" & Text1.Text & """" & " BEHAVIOR=""" & Combo3.Text & """" & " SCROLLDELAY=""" & Text2.Text & """" & " DIRECTION=""" & Combo4.Text & """" & " BGCOLOR=""" & Combo5.Text & """" & " HSPACE=""" & Text3.Text & """" & " VSPACE=""" & Text4.Text & """" & " LOOP=INFINITE>" & Text5.Text & "</MARQUEE>" & vbCrLf
                Unload Me
            End If
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Command2_Click()
    dlgColor.CancelError = False
    dlgColor.ShowColor
    Combo5.AddItem dlgColor.Color
End Sub

Private Sub FlatCombo(ctl As ComboBox)
    Dim ctlx As Control

    For Each ctlx In Me.Controls
        If TypeOf ctl Is ComboBox Then
            m_iCount = m_iCount + 1
            ReDim Preserve m_cFlatten(1 To m_iCount) As cFlatControl
            Set m_cFlatten(m_iCount) = New cFlatControl
            m_cFlatten(m_iCount).Attach ctl
        End If
    Next ctlx
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To 10
        Combo1(0).AddItem i * 10 & "%"
        Combo1(1).AddItem i * 10 & "%"
    Next i
    
    Combo2.AddItem "BOTTOM"
    Combo2.AddItem "MIDDLE"
    Combo2.AddItem "TOP"
    
    Combo3.AddItem "SCROLL"
    Combo3.AddItem "SLIDE"
    Combo3.AddItem "ALTERNATE"
    
    Combo4.AddItem "LEFT"
    Combo4.AddItem "RIGHT"
    
    Combo5.AddItem "RED"
    Combo5.AddItem "BLUE"
    Combo5.AddItem "GREEN"
    Combo5.AddItem "YELLOW"
    Combo5.AddItem "ORANGE"
    Combo5.AddItem "WHITE"
    Combo5.AddItem "BLACK"
    Combo5.AddItem "PURPLE"
    
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Combo5.ListIndex = 0
End Sub
