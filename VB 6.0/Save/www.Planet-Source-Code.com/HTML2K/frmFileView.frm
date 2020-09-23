VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFileView 
   Caption         =   "X"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   ScaleHeight     =   4035
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":06F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":09A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":0B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":0C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":11F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1790
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":18EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2106
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2260
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":23BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEdit 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save To Do List"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print To Do List"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo Typing"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Del"
            Object.ToolTipText     =   "Delete Selection"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo Typing"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThru"
            Object.ToolTipText     =   "Strike Through"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Change text color"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullet"
            Object.ToolTipText     =   "Bullets"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtFile 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmFileView.frx":2514
   End
End
Attribute VB_Name = "frmFileView"
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

Public ToDoChanged As Boolean

Private trapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements

Private Sub Form_Load()
    Dim intFileNum As Integer
    
    On Error Resume Next
    frmFileView.Width = GetSetting(ThisApp, "To Do List", "Width")
    frmFileView.Height = GetSetting(ThisApp, "To Do List", "Height")
    frmFileView.Left = GetSetting(ThisApp, "To Do List", "Left")
    frmFileView.Top = GetSetting(ThisApp, "To Do List", "Top")

    Me.Caption = "To Do List"
    
    If Dir(App.Path & "\To_Do.rtf") <> Empty Then
        txtFile.LoadFile (App.Path & "\To_Do.rtf")
    Else
        '// create file
        intFileNum = FreeFile
        Open App.Path & "\To_Do.rtf" For Output As intFileNum
        Close #intFileNum
    End If
    
'    txtFile.LoadFile (App.Path & "\Miscellaneous\To_Do.rtf")

    trapUndo = True     'Enable Undo Trapping
    txtFile_Change      'Initialize First Undo
    txtFile_SelChange   'Initialize Menus
    Show
    DoEvents
    txtFile.SetFocus
    tbEdit.Buttons(1).Enabled = False
    ToDoChanged = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' msgbox "This is a test"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtFile.Width = Me.ScaleWidth
    txtFile.Height = Me.ScaleHeight - tbEdit.Height - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting ThisApp, "To Do List", "Width", frmFileView.Width
    SaveSetting ThisApp, "To Do List", "Height", frmFileView.Height
    SaveSetting ThisApp, "To Do List", "Left", frmFileView.Left
    SaveSetting ThisApp, "To Do List", "Top", frmFileView.Top
    
    ' Delete an entry
    Dim reply As Integer
    ' Confirm delete
    If ToDoChanged = True Then
        reply = MsgBox("Do you want to save your changes?", vbOKCancel + vbQuestion, "Confirm Save")
        ' Cancel delete
        If reply = vbCancel Then
            Unload frmFileView
            Exit Sub
        Else
            txtFile.SaveFile (App.Path & "\To_Do.rtf")
            ToDoChanged = False
        End If
    End If
    ToDoChanged = False
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "Save"
            txtFile.SaveFile (App.Path & "\To_Do.rtf")
            tbEdit.Buttons(1).Enabled = False
            ToDoChanged = False
        Case "Print"
            PrintToDoList
        Case "Cut"
            ' Copy the selected text onto the Clipboard.
            Clipboard.SetText txtFile.SelText
            ' Delete the selected text.
            txtFile.SelText = ""
        Case "Copy"
            Clipboard.SetText txtFile.SelText
        Case "Paste"
            ' Place the text from the Clipboard into the active control.
            txtFile.SelText = Clipboard.GetText()
        Case "Undo"
            Undo
        Case "Del"
            txtFile.SelText = ""
        Case "Redo"
            Redo
        Case "Bold"
            txtFile.SelBold = Not txtFile.SelBold
            Button.Value = IIf(txtFile.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            txtFile.SelItalic = Not txtFile.SelItalic
            Button.Value = IIf(txtFile.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            txtFile.SelUnderline = Not txtFile.SelUnderline
            Button.Value = IIf(txtFile.SelUnderline, tbrPressed, tbrUnpressed)
        Case "StrikeThru"
            txtFile.SelStrikeThru = Not txtFile.SelStrikeThru
            Button.Value = IIf(txtFile.SelStrikeThru, tbrPressed, tbrUnpressed)
        Case "Color"
            On Error GoTo ColorError
            dlgColor.CancelError = True
            dlgColor.FontName = Screen.ActiveForm.FontName
            dlgColor.Flags = cdlCCFullOpen ' used for the full color dialog box.
            dlgColor.Color = txtFile.SelColor
            dlgColor.ShowColor
            txtFile.SelColor = dlgColor.Color
            Exit Sub
ColorError:
            Exit Sub
        Case "Left"
            txtFile.SelAlignment = rtfLeft
        Case "Center"
            txtFile.SelAlignment = rtfCenter
        Case "Right"
            txtFile.SelAlignment = rtfRight
        Case "Bullet"
            txtFile.SelBullet = Not txtFile.SelBullet
            Button.Value = IIf(txtFile.SelBullet, tbrPressed, tbrUnpressed)
        Case "Exit"
            SaveSetting ThisApp, "To Do List", "Width", frmFileView.Width
            SaveSetting ThisApp, "To Do List", "Height", frmFileView.Height
            Unload Me
    End Select
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
    Dim tempParam$
    Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub Undo()
    Dim chg$, X&
    Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
    Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            X& = SendMessage(txtFile.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            txtFile.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            txtFile.SelLength = objElement.TextLen - objElement2.TextLen
            txtFile.SelText = ""
            X& = SendMessage(txtFile.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            txtFile.SelStart = objElement2.SelStart
            txtFile.SelLength = 0
            txtFile.SelText = chg$
            txtFile.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                txtFile.SelLength = Len(chg$)
            Else
                txtFile.SelStart = txtFile.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    EnableControls
    trapUndo = True
    txtFile.SetFocus
End Sub

Public Sub Redo()
    Dim chg$
    Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
    Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(txtFile.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            txtFile.SelStart = objElement.SelStart
            txtFile.SelLength = Len(txtFile.Text) - objElement.TextLen
            txtFile.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(txtFile.Text, objElement.Text, objElement.SelStart + 1)
            txtFile.SelStart = objElement.SelStart - Len(chg$)
            txtFile.SelLength = 0
            txtFile.SelText = chg$
            txtFile.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                txtFile.SelLength = Len(chg$)
            Else
                txtFile.SelStart = txtFile.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    EnableControls
    trapUndo = True
    txtFile.SetFocus
End Sub

Private Sub txtFile_Change()
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim C%, l&

    'remove all redo items because of the change
    For C% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next C%

    'set the values of the new element
    newElement.SelStart = txtFile.SelStart
    newElement.TextLen = Len(txtFile.Text)
    newElement.Text = txtFile.Text

    'add it to the undo stack
    UndoStack.Add item:=newElement
    'enable controls accordingly
    EnableControls
    
'    UpdateDisplay
    
    ToDoChanged = True
    tbEdit.Buttons(1).Enabled = True
    tbEdit.Buttons(2).Enabled = True
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If GetSetting(ThisApp, "Options", "AutoIndent") = "1" Then
        ' Do auto-indent
        If KeyCode = 13 Then
            ' Do not do auto-indent unless the user wants it
                If GetSetting(ThisApp, "Options", "AutoIndent") <> 0 Then
                ' Implement auto-indent
                ToDoIndent
                ' Null the keycode
                KeyCode = 0
            Else
                KeyCode = 0
            End If
        End If
    End If
End Sub

Private Sub txtFile_SelChange()
    Dim Ln As Long
    Ln = txtFile.SelLength
    
'    With frmMain
'        .Toolbar1.Buttons.item("CUT").Enabled = Ln
'        .Toolbar1.Buttons.item("COPY").Enabled = Ln
'        .mnuEdit1(3).Enabled = Ln
'        .mnuEdit1(4).Enabled = Ln
'        .mnuFormat(10).Enabled = Ln
'    End With
    
    tbEdit.Buttons("Cut").Enabled = Ln
    tbEdit.Buttons("Copy").Enabled = Ln
    tbEdit.Buttons("Bold").Value = IIf(txtFile.SelBold, tbrPressed, tbrUnpressed)
    tbEdit.Buttons("Italic").Value = IIf(txtFile.SelItalic, tbrPressed, tbrUnpressed)
    tbEdit.Buttons("Underline").Value = IIf(txtFile.SelUnderline, tbrPressed, tbrUnpressed)
    tbEdit.Buttons("Left").Value = IIf(txtFile.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    tbEdit.Buttons("Center").Value = IIf(txtFile.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    tbEdit.Buttons("Right").Value = IIf(txtFile.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub

Private Sub EnableControls()
    tbEdit.Buttons(8).Enabled = UndoStack.Count > 1
    tbEdit.Buttons(10).Enabled = RedoStack.Count > 0
    txtFile_SelChange
End Sub

Private Sub PrintToDoList()
    Dim T As String
    Dim H As Integer
    Dim l As Integer
    Dim P As Integer
    Dim ct As String
    Dim i As Integer
    Dim J As Long
    Dim tt As String
    Dim garniture As Boolean
    Dim page As Integer
'   On Error GoTo NotepadError
    TitreVar = "Printing"
    MsgVar = "Printer: " + Printer.DeviceName + LfVar
    MsgVar = MsgVar + " over " + Printer.Port + LfVar + LfVar
    MsgVar = MsgVar + "Would you like to print... ?" + LfVar
    MsgVar = MsgVar + Me.Caption
    RepVar = MsgBox(MsgVar, 33, TitreVar)
    Select Case RepVar
    Case 1  ' Ok
    ' rien à faire
    Case 2  ' Annuler
    Exit Sub
    End Select
    ' retrouver les paramètres de l'imprimante
    MousePointer = 11
    Printer.ScaleMode = vbInches
    H = Printer.ScaleHeight
    l = Printer.ScaleWidth
    
    ' initialise le printer
    Printer.Print " ";
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    page = 1
    
    ' Si imprime la garniture
    If GetSetting(App.Title, "Settings", "PositionPrefere 0", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = 0.25
    Printer.Print Me.Caption
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 1", 1) = 1 Then
    Printer.Line (0.5, 0.5)-(l - 1, 0.5)
    Printer.Line (0.05, H - 1)-(l - 1, H - 1)
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 2", 1) = 1 Then
    Printer.CurrentX = 0.5     ' was 1
    Printer.CurrentY = H - 0.75
    Printer.Print "Page " + str(page);
    garniture = True
    End If
    
    ' replace le printer au debut
    Printer.CurrentX = 0.5
    If garniture = True Then
    Printer.CurrentY = 0.65
    Else
    Printer.CurrentY = 0.25
    End If
    
    ' Impression du texte
    T = txtFile.Text
    J = Len(T)
    P = l - 2   ' largeur de la page -2 po.
    ' Sélectionne les caractères d'impression
    Printer.FontBold = GetSetting(App.Title, "Settings", "FontBold", 0)
    Printer.FontItalic = GetSetting(App.Title, "Settings", "Italic", 0)
    Printer.FontName = GetSetting(App.Title, "Settings", "FontName", "Arial")
    Printer.FontSize = GetSetting(App.Title, "Settings", "FontSize", 10)
    Do
    i = i + 1
    ct = ct + Mid(T, i, 1)  ' chaine temporaire
    ' si la page est pleine
    If Printer.CurrentY > H - 1.25 Then     ' was 1.25
    Printer.NewPage
    page = page + 1
    Printer.Print " ";
    
    ' Si imprime la garniture
    If GetSetting(App.Title, "Settings", "PositionPrefere 0", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = 0.25
    Printer.Print Me.Caption
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 1", 1) = 1 Then
    Printer.Line (0.5, 0.5)-(l - 1, 0.5)
    Printer.Line (0.5, H - 1)-(l - 1, H - 1)
    garniture = True
    End If
    If GetSetting(App.Title, "Settings", "PositionPrefere 2", 1) = 1 Then
    Printer.CurrentX = 0.5
    Printer.CurrentY = H - 0.75
    Printer.Print "Page " + str(page);
    garniture = True
    End If
    
    ' replace le printer au debut
    Printer.CurrentX = 0.5
    If garniture = True Then
    Printer.CurrentY = 0.65
    Else
    Printer.CurrentY = 0.25
    End If
    
    Printer.FontBold = GetSetting(App.Title, "Settings", "FontBold", 0)
    Printer.FontItalic = GetSetting(App.Title, "Settings", "Italic", 0)
    Printer.FontName = GetSetting(App.Title, "Settings", "FontName", "Arial")
    Printer.FontSize = GetSetting(App.Title, "Settings", "FontSize", 10)
    End If
    
    If i >= J Then ' si la fin du texte
    'Debug.Print ct
    Printer.CurrentX = 0.5
    ' vide le buffer
    Printer.Print ct;
    ' pi sort de là
    Exit Do
    End If
    If Printer.TextWidth(ct) >= P Then
    ' wrap
    ' couper à un espace
    tt = ct
    Do
    ' recule dans tt (text temporaire)
    ' jusqu'au premier espace
    If Right(tt, 1) = Chr(32) Then
    ' conserve dans ct le reste
    ct = Mid(ct, Len(tt) + 1)
    Exit Do
    End If
    tt = Left(tt, Len(tt) - 1)
    Loop
    'Debug.Print tt
    Printer.CurrentX = 0.5
    Printer.Print tt
    tt = ""
    End If
    If Mid(T, i, 1) = Chr(13) Or Mid(T, i, 1) = Chr(10) Then
    ' si je frappe un carriage return
    ct = Left(ct, Len(ct) - 1)
    'Debug.Print ct
    Printer.CurrentX = 0.5
    Printer.Print ct
    ct = ""
    i = i + 1   ' sauter le chr(10)
    End If
    Loop
    Printer.EndDoc
    MousePointer = 1
    On Error GoTo 0
End Sub
