Attribute VB_Name = "modMenuStat"
Option Explicit

' Menustat sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/menustat.htm"
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Const GWL_WNDPROC = (-4)
Public Const WM_MENUSELECT = &H11F
Public Const MF_SYSMENU = &H2000&
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

Public origWndProc As Long

Public Sub SetHook(hwnd, bSet As Boolean)
    If bSet Then
        origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
    ElseIf origWndProc Then
        Dim lRet As Long
        lRet = SetWindowLong(hwnd, GWL_WNDPROC, origWndProc)
        origWndProc = 0
    End If
End Sub

Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim iHi As Integer, iLo As Integer
    Select Case Msg
        Case WM_MENUSELECT
            frmMain.sbStatusBar.Panels("lsp").Text = ""
            CopyMemory iLo, wParam, 2
            CopyMemory iHi, ByVal VarPtr(wParam) + 2, 2
            If (iHi And MF_SYSMENU) = 0 Then
                Dim m As MENUITEMINFO, aCap As String
                m.dwTypeData = Space$(64)
                m.cbSize = Len(m)
                m.cch = 64
                m.fMask = MIIM_DATA Or MIIM_TYPE
                If GetMenuItemInfo(lParam, CLng(iLo), False, m) Then
                    aCap = m.dwTypeData & Chr$(0)
                    aCap = Left$(aCap, InStr(aCap, Chr$(0)) - 1)
                    If GetMenu(hwnd) <> lParam Then
                        Select Case aCap
                            Case "&New" & vbTab & "Ctrl+N": frmMain.Caption = "New file..."
'                            Case "&New": frmMain.sbStatusBar.Panels("1").Text = "NEW file..."
                            Case "&Open": frmMain.sbStatusBar.Panels("1").Text = "Open a file"
                            Case "&Save": frmMain.sbStatusBar.Panels("lsp").Text = "Save a file"
                            Case "E&xit": frmMain.sbStatusBar.Panels("lsp").Text = "Exit the program"
                            Case "C&ut": frmMain.sbStatusBar.Panels("lsp").Text = "Cut text to the clipboard"
                            Case "&Copy": frmMain.sbStatusBar.Panels("lsp").Text = "Copy text to the clipboard"
                            Case "&Paste": frmMain.sbStatusBar.Panels("lsp").Text = "Paste text from the clipboard"
                            Case Else: frmMain.sbStatusBar.Panels("lsp").Text = aCap
                        End Select
                    Else
                        frmMain.sbStatusBar.Panels("lsp").Text = "Top level menu item - " & aCap
                    End If
                End If
            End If
    End Select
    AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
End Function
