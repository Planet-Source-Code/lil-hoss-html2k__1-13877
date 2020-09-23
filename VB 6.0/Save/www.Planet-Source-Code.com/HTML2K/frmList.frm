VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Wizard"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtItems 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   550
         Width           =   1935
      End
      Begin VB.OptionButton optListStyle 
         Caption         =   "&Ordered List"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Style"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   550
         Width           =   1935
      End
      Begin VB.OptionButton optListStyle 
         Caption         =   "&Unordered List"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Style"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmList"
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

Dim ListStyle As String
Dim lstStyle As String
Dim StartTag As String
Dim EndTag As String


Private Sub Command1_Click(Index As Integer)
    Dim nRet
    Dim Tags$
    Dim ListTags
    Dim EndTag
    Dim J As Integer
    Dim i As Integer
    Dim item$
    Dim enditem$
    Dim Header As String
    Dim Footer As String
    
    Header = "<!-- This List was created by HTML2K -->" & vbCrLf & _
             "<!-- Created on: " & Now & " -->" & vbCrLf
    Footer = vbCrLf & "<!-- End List -->"
    
    Select Case Index
        Case 0
            ' CREATE
            item = CaseTag("<LI>")
            enditem$ = CaseTag("</LI>")
            
            If optListStyle(0).Value = True Then
                EndTag = "</OL>"
                If Combo1.ListIndex = 0 Then
                    StartTag = "<OL>" & vbCrLf
                ElseIf Combo1.ListIndex = 1 Then
                    StartTag = "<OL type=""1"">" & vbCrLf
                ElseIf Combo1.ListIndex = 2 Then
                    StartTag = "<OL type=""A"">" & vbCrLf
                ElseIf Combo1.ListIndex = 3 Then
                    StartTag = "<OL type=""a"">" & vbCrLf
                ElseIf Combo1.ListIndex = 4 Then
                    StartTag = "<OL type=""I"">" & vbCrLf
                ElseIf Combo1.ListIndex = 5 Then
                    StartTag = "<OL type=""i"">" & vbCrLf
                End If
            ElseIf optListStyle(1).Value = True Then
                EndTag = "</UL>"
                If Combo2.ListIndex = 0 Then
                    StartTag = "<UL>" & vbCrLf
                ElseIf Combo2.ListIndex = 1 Then
                    StartTag = "<UL type=""disc"">" & vbCrLf
                ElseIf Combo2.ListIndex = 2 Then
                    StartTag = "<UL type=""circle"">" & vbCrLf
                ElseIf Combo2.ListIndex = 3 Then
                    StartTag = "<UL type=""square"">" & vbCrLf
                End If
            End If

            ListTags = Tags & IIf(Len(Tags) > 0, vbCrLf, "")
            J = 1
            For i = 1 To Len(txtItems.Text)
                If Mid(txtItems.Text, i, 2) = vbCrLf Or i >= Len(txtItems.Text) Then
                    ListTags = ListTags & "    " & item & Mid(txtItems.Text, J, _
                        IIf(i >= Len(txtItems.Text), i - J + 1, i - J)) & enditem & vbCrLf
                    J = i + 2
                    i = i + 2
                End If
            Next i
            If Len(Tags) > 0 Then ListTags = ListTags & EndTag
            ' Delete selected text first to avoid screwing up the Undo stack
            With frmMain.CurrentDoc.rtfHTML
                .SelText = ""
                .SelText = Header
                .SelText = StartTag
                .SelText = ListTags
                .SelText = EndTag
                .SelText = Footer
            End With
            Unload Me
        Case 1
            ' CANCEL
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    With Combo1
        .AddItem "Default"
        .AddItem "1, 2, 3"
        .AddItem "A, B, C"
        .AddItem "a, b, c"
        .AddItem "I, II, III"
        .AddItem "i, ii, iii"
        .ListIndex = 0
    End With
    
    With Combo2
        .AddItem "Default"
        .AddItem "Disc"
        .AddItem "Circle"
        .AddItem "Square"
        .ListIndex = 0
    End With
    
    EndTag = "</OL>"
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

Private Sub optListStyle_Click(Index As Integer)
    Select Case Index
        Case 0
            If optListStyle(0).Value = True Then
                Combo1.Enabled = True
                Combo2.Enabled = False
                optListStyle(1).Value = False
                lstStyle = "OL"
                EndTag = "</OL>"
            End If
        Case 1
            If optListStyle(1).Value = True Then
                Combo2.Enabled = True
                Combo1.Enabled = False
                optListStyle(0).Value = False
                lstStyle = "UL"
                EndTag = "</UL>"
            End If
    End Select
End Sub
