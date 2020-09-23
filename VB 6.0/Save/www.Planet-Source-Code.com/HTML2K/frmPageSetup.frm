VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Setup..."
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Page Number"
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Line Demarcation"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   390
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Title"
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   60
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   525
      Left            =   3360
      TabIndex        =   1
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   2190
      X2              =   2190
      Y1              =   60
      Y2              =   2880
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   90
      X2              =   2160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   90
      X2              =   2190
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   90
      X2              =   90
      Y1              =   2820
      Y2              =   60
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   1200
      X2              =   1680
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   390
      X2              =   1680
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   390
      X2              =   1440
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   390
      X2              =   1680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   390
      X2              =   1680
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   390
      X2              =   1680
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   390
      X2              =   1350
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   390
      X2              =   1680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   390
      X2              =   1680
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   390
      X2              =   1560
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   390
      X2              =   1680
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   390
      X2              =   1680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   240
      Top             =   570
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Page Number"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   270
      X2              =   2040
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   1845
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   2010
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2775
      Left            =   120
      Top             =   90
      Width           =   2055
   End
End
Attribute VB_Name = "frmPageSetup"
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


Private Sub Check1_Click(Index As Integer)
   Select Case Index
      Case 0
         If Check1(0).Value = 1 Then
            Label1(0).Visible = True
         Else
            Label1(0).Visible = False
         End If
      Case 1
         If Check1(1).Value = 1 Then
            Line1(0).Visible = True
            Line1(1).Visible = True
         Else
            Line1(0).Visible = False
            Line1(1).Visible = False
         End If
      Case 2
         If Check1(2).Value = 1 Then
            Label1(1).Visible = True
         Else
            Label1(1).Visible = False
         End If
   End Select
   If Check1(0).Value = False And Check1(1).Value = False And Check1(2).Value = False Then
      Shape3.Top = 220
      Shape3.Height = 2500
   Else
      Shape3.Top = 570
      Shape3.Height = 1935
   End If
End Sub

Private Sub Command1_Click()
   Dim i As Integer
   For i = 0 To 2
      SaveSetting App.Title, "Settings", "PositionPrefere" + str(i), Check1(i).Value
   Next
   Unload Me
   Set frmPageSetup = Nothing
End Sub


Private Sub Form_Load()
   Dim i As Integer
   For i = 0 To 2
      Check1(i).Value = GetSetting(App.Title, "Settings", "PositionPrefere" + str(i), 1)
   Next
   Label1(0).Caption = "Title: " + Left(frmMain.Caption, 17) + "..."
   If Check1(0).Value = 0 Then Label1(0).Visible = False
   If Check1(2).Value = 0 Then Label1(1).Visible = False
   If Check1(1).Value = 0 Then
      Line1(0).Visible = False
      Line1(1).Visible = False
   End If
End Sub


