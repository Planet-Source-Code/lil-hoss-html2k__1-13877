VERSION 5.00
Begin VB.UserControl SplitPanel 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   ControlContainer=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1170
   Begin VB.PictureBox Splitter 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   855
      ScaleHeight     =   3540
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   30
      Width           =   165
   End
End
Attribute VB_Name = "SplitPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Const SPLITWIDTH As Single = 80     ' width of splitterbar

'********************************
' Variables for properties
'********************************
Private mHorizontalSplit As Boolean
Private mControl1 As Object
Private mControl2 As Object
Private mSplitPercent As Single

'********************************
' Read-Write Properties
'********************************
Public Property Get HorizontalSplit() As Boolean
    HorizontalSplit = mHorizontalSplit
End Property
Public Property Let HorizontalSplit(val As Boolean)
    mHorizontalSplit = val
    If mHorizontalSplit Then
        Splitter.MousePointer = 7
    Else
        Splitter.MousePointer = 9
    End If
    PropertyChanged "HorizontalSplit"
    UserControl_Resize
End Property

Public Property Get Control1() As Object
    Set Control1 = mControl1
End Property
Public Property Set Control1(ctl As Object)
    Set mControl1 = ctl
    PropertyChanged "Control1"
    UserControl_Resize
End Property

Public Property Get Control2() As Object
    Set Control2 = mControl2
End Property
Public Property Set Control2(ctl As Object)
    Set mControl2 = ctl
    PropertyChanged "Control2"
    UserControl_Resize
End Property

Public Property Get SplitPercent() As Byte
    SplitPercent = mSplitPercent * 100
End Property
Public Property Let SplitPercent(val As Byte)
    mSplitPercent = val / 100
    PropertyChanged "SplitPercent"
    UserControl_Resize
End Property

'********************************
' Set up the defaults
'********************************
Private Sub UserControl_InitProperties()
    HorizontalSplit = False
    SplitPercent = 50
End Sub

'********************************
' Reload design-time settings
'********************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    HorizontalSplit = PropBag.ReadProperty("HorizontalSplit", False)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
End Sub

'********************************
' Save design-time settings
'********************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HorizontalSplit", HorizontalSplit, False
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
End Sub

'********************************
' These next three subs handle the actual
' dragging of the splitterbar.  The panes
' are updated when the mouse button is
' released.
'********************************
Private Sub splitter_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Splitter.BackColor = &H80000008     ' Make the splitter visible
    Splitter.ZOrder
End Sub

Private Sub splitter_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mHorizontalSplit Then        ' horizontal figures
            Y = Splitter.Top - (SPLITWIDTH - Y)
            mSplitPercent = Y / UserControl.Height
            Splitter.Move 0, Y
        Else                                    ' vertical
            X = Splitter.Left - (SPLITWIDTH - X)
            mSplitPercent = X / UserControl.Width
            Splitter.Move X
        End If
        If mSplitPercent < 0.1 Then mSplitPercent = 0.1     ' Check if in range
        If mSplitPercent > 0.9 Then mSplitPercent = 0.9
    End If
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Splitter.BackColor = &H8000000F     ' change the color back to normal
    UserControl_Resize                          ' update the panes
End Sub

'********************************
' The resize event is where it get's ugly
' Here we must figure out the sizes and
' positions of everything based on the splitter
' position, and the controls properties, then
' set everything
'********************************
Private Sub UserControl_Resize()
    On Error Resume Next
    
    If UserControl.Ambient.UserMode Then    ' get rid of border in run mode
        UserControl.BorderStyle = 0
    End If
    
    Dim pane1 As Single
    Dim pane2 As Single
    Dim totwidth As Single
    Dim totheight As Single
    totwidth = UserControl.Width
    totheight = UserControl.Height
    If mHorizontalSplit Then
        pane1 = (totheight - SPLITWIDTH) * mSplitPercent
        pane2 = (totheight - SPLITWIDTH) * (1 - mSplitPercent)
        mControl1.Move 0, 0, totwidth, pane1
        mControl2.Move 0, pane1 + SPLITWIDTH, totwidth, pane2
        Splitter.Move 0, pane1, totwidth, SPLITWIDTH
    Else
        pane1 = (totwidth - SPLITWIDTH) * mSplitPercent
        pane2 = (totwidth - SPLITWIDTH) * (1 - mSplitPercent)
        mControl1.Move 0, 0, pane1, totheight
        mControl2.Move pane1 + SPLITWIDTH, 0, pane2, totheight
        Splitter.Move pane1, 0, SPLITWIDTH, totheight
    End If
End Sub

