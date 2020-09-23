Attribute VB_Name = "modStartUp"
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

Public fMainForm As frmMain


Sub Main()
    On Error Resume Next
    ShowSplash = RegPeek("ShowSplash", True)
    If GetSetting(ThisApp, "Settings", "Organization") <> "" Then
        If ShowSplash Then
            frmSplash.Show
            frmSplash.Refresh
            Load frmMain
            frmMain.Show
            Unload frmSplash
        Else
            Load frmMain
            frmMain.Show
        End If
    Else
        frmLicense.Show
    End If
End Sub
