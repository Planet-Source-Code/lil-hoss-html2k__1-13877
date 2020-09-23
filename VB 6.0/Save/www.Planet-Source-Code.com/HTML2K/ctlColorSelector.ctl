VERSION 5.00
Begin VB.UserControl ColorSelector 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   KeyPreview      =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   2535
   ToolboxBitmap   =   "ctlColorSelector.ctx":0000
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3240
      Left            =   45
      ScaleHeight     =   3210
      ScaleWidth      =   1965
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdPopup 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1815
      TabIndex        =   1
      Top             =   30
      Width           =   225
   End
   Begin VB.PictureBox picSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   15
      Width           =   1755
   End
End
Attribute VB_Name = "ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80

Dim m_SelectedColor As Integer
'Event Declarations:
Event Click() 'MappingInfo=picPopup,picPopup,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Public Sub AboutBox()
    About
End Sub

Private Sub cmdPopup_Click()
    picSelection.SetFocus 'so we dont see the Focus Rectangle
    'Show or hide the popup window
    If picPopup.Visible = False Then
        ShowPopUp
    Else
        HidePopUp
    End If
End Sub

Private Sub picPopup_Paint()
    Dim a As Integer
    Dim nRowHeight As Long
    
    'paint the color bands
    nRowHeight = Int(picPopup.ScaleHeight / 16)
    For a = 0 To 15
        picPopup.Line (Screen.TwipsPerPixelX, (a * nRowHeight) + Screen.TwipsPerPixelY)-(picPopup.ScaleWidth - (2 * Screen.TwipsPerPixelX), ((a + 1) * nRowHeight) - Screen.TwipsPerPixelY), QBColor(0), B
        picPopup.Line (2 * Screen.TwipsPerPixelX, (a * nRowHeight) + (2 * Screen.TwipsPerPixelY))-(picPopup.ScaleWidth - (3 * Screen.TwipsPerPixelX), ((a + 1) * nRowHeight) - (2 * Screen.TwipsPerPixelY)), QBColor(a), BF
    Next a
    
End Sub

Private Sub picSelection_Click()
    'Fire the click event
    cmdPopup_Click
End Sub

Private Sub picSelection_GotFocus()
    picSelection_Paint
End Sub

Private Sub picSelection_LostFocus()
    picSelection_Paint
End Sub

Private Sub picSelection_Paint()
    'Draw a focus rectangle
    Dim rct As RECT
    
    If GetFocus = picSelection.hwnd And picPopup.Visible = False Then
        GetClientRect picSelection.hwnd, rct
        With rct
            .Left = .Left + 1
            .Right = .Right - 1
            .Top = .Top + 1
            .Bottom = .Bottom - 1
        End With
        DrawFocusRect picSelection.hdc, rct
    Else
        picSelection.Cls
    End If
    'Paint the interior with the selected color
    picSelection.Line (2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY)-(picSelection.ScaleWidth - (3 * Screen.TwipsPerPixelX), picSelection.ScaleHeight - (3 * Screen.TwipsPerPixelY)), QBColor(m_SelectedColor), BF
    
End Sub

Private Sub UserControl_ExitFocus()
    'Although in most circumstances the popup window will have already been
    'hidden before this, we check here just in case.
    If picPopup.Visible Then HidePopUp
End Sub

Private Sub UserControl_Initialize()
    'Set the parent and window style for the popup picturebox
    'set style to Toolwindow so after we've set parent to the Desktop
    'the popup doesn't show in the Taskbar
    SetWindowLong picPopup.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent picPopup.hwnd, 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the popup to show/hide
    'Alt and either the up or down arrow toggle the show state of the popup
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And (Shift = 4) Then
        cmdPopup_Click
    ElseIf KeyCode = vbKeyDown And m_SelectedColor < 15 Then
        m_SelectedColor = m_SelectedColor + 1
        picSelection_Paint
        RaiseEvent Click
    ElseIf KeyCode = vbKeyUp And m_SelectedColor > 0 Then
        m_SelectedColor = m_SelectedColor - 1
        picSelection_Paint
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'Position the constituent controls
    cmdPopup.Move UserControl.ScaleWidth - cmdPopup.Width, 0, cmdPopup.Width, UserControl.ScaleHeight
    picSelection.Move 0, 0, UserControl.ScaleWidth - (cmdPopup.Width + Screen.TwipsPerPixelX), UserControl.ScaleHeight
    picPopup.Width = UserControl.Extender.Width
End Sub

Private Sub HidePopUp()
    'This procedure is called whenever the popup window needs to be hidden.
    If GetCapture = picPopup.hwnd Then
        ReleaseCapture
    End If
    picPopup.Visible = False
    DoEvents
    picSelection_Paint
End Sub

Private Sub ShowPopUp()

    'This procedure is called whenever the popup needs to be shown.
    
    Dim ileft As Long
    Dim iTop As Long
    Dim ctlRect As RECT
    
    'Determine position for pop up window
    'We want to show the popup below the control, but if we can't we'll show it above
    GetWindowRect UserControl.hwnd, ctlRect 'screen rectange of the control
    If ctlRect.Bottom + (picPopup.Height / Screen.TwipsPerPixelX) > Screen.Height / Screen.TwipsPerPixelY Then
        'put it above
        iTop = (ctlRect.Top - (picPopup.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
    Else
        'put it below
        iTop = ctlRect.Bottom * Screen.TwipsPerPixelY
    End If
    'If the popup window is as wide as, or wider than the control, we want to align
    'it to the left edge of the control.  Otherwise, we align it to the right.  If
    'we're too far to the right, we push it back left.
    If (ctlRect.Right - ctlRect.Left) > picPopup.Width / Screen.TwipsPerPixelX Then
        'try to align to the right of the control
        If ctlRect.Right > Screen.Width / Screen.TwipsPerPixelX Then
            ileft = Screen.Width - picPopup.Width
        Else
            ileft = ctlRect.Right * Screen.TwipsPerPixelX - picPopup.Width
        End If
        'Check we haven't gone outside the left edge of the screen
        If ileft < 0 Then ileft = 0
    Else
        'try to align to the left
        If ctlRect.Left < 0 Then
            ileft = 0
        Else
            ileft = ctlRect.Left * Screen.TwipsPerPixelX
        End If
        'Check we haven't gone outside the left edge of the screen
        If ileft + picPopup.Width > Screen.Width Then ileft = Screen.Width - picPopup.Width
    End If
    
    With picPopup
        .Top = iTop
        .Left = ileft
        .Visible = True
        .ZOrder
    End With
    picPopup_Paint
    DoEvents
    picSelection_Paint
    'Capture the mouse so we get all subsequent mouse clicks
    SetCapture picPopup.hwnd
    
End Sub

Private Sub picPopUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'We've set capture to the popup window, so here we check for mouse presses
    'and if the user clicks outside of the popup, we call the HidePopUp routine
    'to validate and dismiss the popup window.
    If X < 0 Or X > picPopup.Width Or Y < 0 Or Y > picPopup.Height Then
        'user has clicked outside the popup so hide it
        HidePopUp
    ElseIf Button = vbLeftButton Then
        'Calculate the row
        m_SelectedColor = Int(Y / (picPopup.ScaleHeight / 16))
        'update the display
        picSelection_Paint
        HidePopUp
        RaiseEvent Click
    Else
        'nothing to do
    End If

End Sub

Private Sub UserControl_Show()
    'Get the tooltip
    picSelection.ToolTipText = UserControl.Extender.ToolTipText
End Sub

Public Property Let SelectedColor(New_SelectedColor As Integer)
    If New_SelectedColor >= 0 And New_SelectedColor < 16 Then
        m_SelectedColor = New_SelectedColor
        picSelection_Paint
    End If
End Property

Public Property Get SelectedColor() As Integer
    SelectedColor = m_SelectedColor
End Property

Public Property Get Enabled() As Boolean
    Enabled = cmdPopup.Enabled
    Enabled = picPopup.Enabled
    Enabled = picSelection.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmdPopup.Enabled = New_Enabled
    picPopup.Enabled = New_Enabled
    picSelection.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
