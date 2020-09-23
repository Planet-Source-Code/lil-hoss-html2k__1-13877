VERSION 5.00
Begin VB.Form frmFrames 
   Caption         =   "Frames Wizard"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   Icon            =   "frmFrames.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox FrameTwo 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   7275
      TabIndex        =   11
      Top             =   3840
      Width           =   7335
      Begin VB.Frame F4 
         Height          =   3375
         Left            =   4680
         TabIndex        =   56
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Resize4 
            Caption         =   "Resize"
            Height          =   255
            Left            =   1320
            TabIndex        =   63
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox Border4 
            Caption         =   "Borders"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox MW4 
            Height          =   285
            Left            =   1320
            TabIndex        =   61
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox MH4 
            Height          =   285
            Left            =   1320
            TabIndex        =   60
            Top             =   1185
            Width           =   1095
         End
         Begin VB.ComboBox CoS4 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2850
            Width           =   1575
         End
         Begin VB.TextBox Source4 
            Height          =   285
            Left            =   840
            TabIndex        =   58
            Top             =   705
            Width           =   1575
         End
         Begin VB.TextBox Name4 
            Height          =   285
            Left            =   840
            TabIndex        =   57
            Top             =   225
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Width:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Height:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrolling:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox MSapmle 
         BackColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3435
         ScaleWidth      =   4515
         TabIndex        =   51
         Top             =   0
         Width           =   4575
         Begin VB.PictureBox P2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   0
            ScaleHeight     =   825
            ScaleWidth      =   705
            TabIndex        =   55
            Top             =   0
            Width           =   735
         End
         Begin VB.PictureBox P3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1545
            ScaleWidth      =   705
            TabIndex        =   54
            Top             =   960
            Width           =   735
         End
         Begin VB.PictureBox P4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   840
            ScaleHeight     =   1545
            ScaleWidth      =   3705
            TabIndex        =   53
            Top             =   960
            Width           =   3735
         End
         Begin VB.PictureBox P5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   840
            ScaleHeight     =   825
            ScaleWidth      =   3705
            TabIndex        =   52
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.Frame F1 
         Height          =   1695
         Left            =   0
         TabIndex        =   38
         Top             =   -1440
         Width           =   4575
         Begin VB.TextBox Name1 
            Height          =   285
            Left            =   840
            TabIndex        =   45
            Top             =   225
            Width           =   1575
         End
         Begin VB.TextBox Source1 
            Height          =   285
            Left            =   840
            TabIndex        =   44
            Top             =   705
            Width           =   1575
         End
         Begin VB.ComboBox CoS1 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1170
            Width           =   1575
         End
         Begin VB.TextBox MH1 
            Height          =   285
            Left            =   3720
            TabIndex        =   42
            Top             =   225
            Width           =   615
         End
         Begin VB.TextBox MW1 
            Height          =   285
            Left            =   3720
            TabIndex        =   41
            Top             =   705
            Width           =   615
         End
         Begin VB.CheckBox Border1 
            Caption         =   "Borders"
            Height          =   255
            Left            =   2520
            TabIndex        =   40
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox Resize1 
            Caption         =   "Resize"
            Height          =   255
            Left            =   3480
            TabIndex        =   39
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrolling:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Height:"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Width:"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   46
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame F2 
         Height          =   1695
         Left            =   0
         TabIndex        =   25
         Top             =   -1320
         Width           =   4575
         Begin VB.CheckBox Resize2 
            Caption         =   "Resize"
            Height          =   255
            Left            =   3480
            TabIndex        =   32
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Border2 
            Caption         =   "Borders"
            Height          =   255
            Left            =   2520
            TabIndex        =   31
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox MW2 
            Height          =   285
            Left            =   3720
            TabIndex        =   30
            Top             =   705
            Width           =   615
         End
         Begin VB.TextBox MH2 
            Height          =   285
            Left            =   3720
            TabIndex        =   29
            Top             =   225
            Width           =   615
         End
         Begin VB.ComboBox CoS2 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1170
            Width           =   1575
         End
         Begin VB.TextBox Source2 
            Height          =   285
            Left            =   840
            TabIndex        =   27
            Top             =   705
            Width           =   1575
         End
         Begin VB.TextBox Name2 
            Height          =   285
            Left            =   840
            TabIndex        =   26
            Top             =   225
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Width:"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   37
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Height:"
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrolling:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame F3 
         Height          =   1695
         Left            =   0
         TabIndex        =   12
         Top             =   -1440
         Width           =   4575
         Begin VB.TextBox Name3 
            Height          =   285
            Left            =   840
            TabIndex        =   19
            Top             =   225
            Width           =   1575
         End
         Begin VB.TextBox Source3 
            Height          =   285
            Left            =   840
            TabIndex        =   18
            Top             =   705
            Width           =   1575
         End
         Begin VB.ComboBox CoS3 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1170
            Width           =   1575
         End
         Begin VB.TextBox MH3 
            Height          =   285
            Left            =   3720
            TabIndex        =   16
            Top             =   225
            Width           =   615
         End
         Begin VB.TextBox MW3 
            Height          =   285
            Left            =   3720
            TabIndex        =   15
            Top             =   705
            Width           =   615
         End
         Begin VB.CheckBox Border3 
            Caption         =   "Borders"
            Height          =   255
            Left            =   2520
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox Resize3 
            Caption         =   "Resize"
            Height          =   255
            Left            =   3480
            TabIndex        =   13
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrolling:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Height:"
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Margin Width:"
            Height          =   255
            Index           =   8
            Left            =   2520
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox FrameOne 
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox M 
         BackColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   30
         ScaleHeight     =   3435
         ScaleWidth      =   4515
         TabIndex        =   8
         Top             =   0
         Width           =   4575
         Begin VB.PictureBox H 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   50
            Left            =   0
            ScaleHeight     =   45
            ScaleWidth      =   6975
            TabIndex        =   10
            Top             =   360
            Width           =   6975
         End
         Begin VB.PictureBox V 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H000000FF&
            ForeColor       =   &H000000FF&
            Height          =   3495
            Left            =   360
            ScaleHeight     =   3495
            ScaleWidth      =   45
            TabIndex        =   9
            Top             =   0
            Width           =   50
         End
      End
      Begin VB.OptionButton optV 
         Caption         =   "Vertical Frame"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optH 
         Caption         =   "Horizontal Frame"
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label LBLv 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50 %"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label LBLh 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50 %"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GoV As Boolean, GoH As Boolean
Dim RowNum As Integer, ColNum As Integer

Private Sub cmdBack_Click()
If cmdBack.Caption = "&Back" Then
  FrameOne.Visible = True
  FrameTwo.Visible = False
  cmdBack.Caption = "&Cancel"
  cmdFinish.Enabled = False
  Command1.Enabled = True
Else
 Unload Me
End If
End Sub

Private Sub cmdFinish_Click()
 BuildCode
 Unload Me
End Sub

Private Sub Command1_Click()
 Command1.Enabled = False
 cmdBack.Caption = "&Back"
 cmdFinish.Enabled = True
 cmdBack.Enabled = True
 F4.Visible = False
 F3.Visible = False
 F2.Visible = False
 F1.Visible = True
 F1.Top = 3600
 F1.Left = 0

 FrameOne.Visible = False
 FrameTwo.Visible = True
 FrameTwo.Left = FrameOne.Left
 FrameTwo.Top = FrameOne.Top
 
 
 P5.Left = Int(V.Left)
 P4.Left = Int(V.Left)
 P2.Top = 0
 P2.Height = Int(H.Top)
 P5.Top = 0
 P5.Height = Int(H.Top)
 P3.Top = Int(H.Top)
 P4.Top = Int(H.Top)
 
 
 ColNum = Int(100 - V.Left) + 1
 RowNum = Int(100 - H.Top) + 1
End Sub


Private Sub Form_Load()
 M.ScaleHeight = 100
 M.ScaleWidth = 100
 MSapmle.ScaleHeight = 100
 MSapmle.ScaleWidth = 100
 GoV = False
 GoH = False
 V.Left = 10
 H.Top = 10
' H.Left = V.Left
'@#@
 P2.Width = 100
 P3.Width = 100
 P4.Width = 100
 P5.Width = 100
 P3.Height = 100
 P4.Height = 100
 
 CoS1.AddItem "Auto"
 CoS1.AddItem "Yes"
 CoS1.AddItem "No"
 CoS2.AddItem "Auto"
 CoS2.AddItem "Yes"
 CoS2.AddItem "No"
 CoS3.AddItem "Auto"
 CoS3.AddItem "Yes"
 CoS3.AddItem "No"
 CoS4.AddItem "Auto"
 CoS4.AddItem "Yes"
 CoS4.AddItem "No"
 '
 CoS1.ListIndex = 0
 CoS2.ListIndex = 0
 CoS3.ListIndex = 0
 CoS4.ListIndex = 0
 
V.Left = 50
H.Top = 50
End Sub


Private Sub M_Click()

If optV.Value = True Then
 If GoV = False Then
  GoV = True
 Else
  GoV = False
 End If
End If


If optH.Value = True Then
 If GoH = False Then
  GoH = True
 Else
  GoH = False
 End If
End If




End Sub

Private Sub M_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If optV.Value = True Then
 If GoV = True Then
  V.Left = X + 1
   Procentage = Int(V.Left - 1)
    If Procentage < 0 Then Procentage = 0
    If Procentage > 100 Then Procentage = 100
   LBLv.Caption = Procentage & " %"
  'H.Left = V.Left
 End If
  
End If

If optH.Value = True Then
 If GoH = True Then
  H.Top = Y + 1
  Procentage = Int(H.Top - 1)
   If Procentage < 0 Then Procentage = 0
   If Procentage > 100 Then Procentage = 100
  LBLh.Caption = Procentage & " %"
 End If
  
End If

End Sub

Private Sub optH_Click()
 GoH = True
 GoV = False
End Sub

Private Sub optV_Click()
 GoH = False
 GoV = True
End Sub



Private Sub V_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X + 50
End Sub

'Sampler
Private Sub P2_Click()
If P2.BackColor = &HFFFFFF Then
 P2.BackColor = vbRed
 P3.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF
  
  F4.Visible = False
  F3.Visible = False
  F2.Visible = False
  F1.Visible = True
  F1.Top = 3600
  F1.Left = 0

Else
 P2.BackColor = &HFFFFFF
 P3.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF
End If
End Sub


Private Sub P3_Click()
If P3.BackColor = &HFFFFFF Then
 P3.BackColor = vbRed
 P2.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF

 F4.Visible = False
 F3.Visible = False
 F2.Visible = True
 F1.Visible = False
 F2.Top = 3600
 F2.Left = 0

Else
 P3.BackColor = &HFFFFFF
 P2.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF
End If
End Sub

Private Sub P4_Click()
If P4.BackColor = &HFFFFFF Then
 P4.BackColor = vbRed
 P3.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P2.BackColor = &HFFFFFF

 F4.Visible = True
 F3.Visible = False
 F2.Visible = False
 F1.Visible = False
 F4.Top = 3600
 F4.Left = 0


Else
 P4.BackColor = &HFFFFFF
 P3.BackColor = &HFFFFFF
 P5.BackColor = &HFFFFFF
 P2.BackColor = &HFFFFFF
End If
End Sub

Private Sub P5_Click()
If P5.BackColor = &HFFFFFF Then
 P5.BackColor = vbRed
 P3.BackColor = &HFFFFFF
 P2.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF
 
 F4.Visible = False
 F3.Visible = True
 F2.Visible = False
 F1.Visible = False
 F3.Top = 3600
 F3.Left = 0
 
Else
 P5.BackColor = &HFFFFFF
 P3.BackColor = &HFFFFFF
 P2.BackColor = &HFFFFFF
 P4.BackColor = &HFFFFFF
End If
End Sub


Function BuildCode()
Dim Insert As String
Dim Border(4) As String, Resize(4) As String

If Border1.Value = 1 Then
 Border(1) = " frameborder=" & Chr(34) & "1" & Chr(34)
Else
 Border(1) = ""
End If
If Resize1.Value = 1 Then
 Resize(1) = " noresize"
Else
 Resize(1) = ""
End If

If Border2.Value = 1 Then
 Border(2) = " frameborder=" & Chr(34) & "1" & Chr(34)
Else
 Border(2) = ""
End If
If Resize2.Value = 1 Then
 Resize(2) = " noresize"
Else
 Resize(2) = ""
End If
If Border3.Value = 1 Then
 Border(3) = " frameborder=" & Chr(34) & "1" & Chr(34)
Else
 Border(3) = ""
End If
If Resize3.Value = 1 Then
 Resize(3) = " noresize"
Else
 Resize(3) = ""
End If
If Border4.Value = 1 Then
 Border(4) = " frameborder=" & Chr(34) & "1" & Chr(34)
Else
 Border(4) = ""
End If
If Resize4.Value = 1 Then
 Resize(4) = " noresize"
Else
 Resize(4) = ""
End If

 Insert = "<!--- Frames --->" & vbCrLf
 Insert = Insert & "<frameset  rows=" & Chr(34) & RowNum & "%,*" & Chr(34) & " cols=" & Chr(34) & ColNum & "%,*" & ">" & vbCrLf
 Insert = Insert & "    <frame name=" & Chr(34) & Name1.Text & Chr(34) & " src=" & Chr(34) & Source1.Text & Chr(34) & " marginwidth=" & Chr(34) & MW1.Text & Chr(34) & " marginheight=" & Chr(34) & MH1.Text & Chr(34) & " scrolling=" & Chr(34) & CoS1.Text & Chr(34) & Border(1) & Resize(1) & ">" & vbCrLf
 Insert = Insert & "    <frame name=" & Chr(34) & Name3.Text & Chr(34) & " src=" & Chr(34) & Source3.Text & Chr(34) & " marginwidth=" & Chr(34) & MW3.Text & Chr(34) & " marginheight=" & Chr(34) & MH3.Text & Chr(34) & " scrolling=" & Chr(34) & CoS3.Text & Chr(34) & Border(3) & Resize(3) & ">" & vbCrLf
 Insert = Insert & "    <frame name=" & Chr(34) & Name2.Text & Chr(34) & " src=" & Chr(34) & Source2.Text & Chr(34) & " marginwidth=" & Chr(34) & MW2.Text & Chr(34) & " marginheight=" & Chr(34) & MH2.Text & Chr(34) & " scrolling=" & Chr(34) & CoS2.Text & Chr(34) & Border(2) & Resize(2) & ">" & vbCrLf
 Insert = Insert & "    <frame name=" & Chr(34) & Name4.Text & Chr(34) & " src=" & Chr(34) & Source4.Text & Chr(34) & " marginwidth=" & Chr(34) & MW4.Text & Chr(34) & " marginheight=" & Chr(34) & MH4.Text & Chr(34) & " scrolling=" & Chr(34) & CoS4.Text & Chr(34) & Border(4) & Resize(4) & ">" & vbCrLf
 Insert = Insert & "</frameset>" & vbCrLf
 Insert = Insert & "<!--- Frames Ends. --->" & vbCrLf
  'MsgBox Insert
  frmMain.CurrentDoc.rtfHTML.SelText = Insert
End Function

