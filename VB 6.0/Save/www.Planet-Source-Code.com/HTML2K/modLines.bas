Attribute VB_Name = "modLines"
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Const DT_CALCRECT = &H400
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20


Public Sub DrawLines(picTo As PictureBox, txtThis As htmSyntaxBox)
Dim lLine As Long
Dim lCount As Long
Dim lCurrent As Long
Dim hBr As Long
Dim lEnd As Long
Dim lhDC As Long
Dim bComplete As Boolean
Dim tR As RECT, tTR As RECT
Dim oCol As OLE_COLOR
Dim lStart As Long
Dim lEndLine As Long
Dim tPO As POINTAPI
Dim lLineHeight As Long
Dim hPen As Long
Dim hPenOld As Long

   'Debug.Print "DrawLines"
   lhDC = picTo.hdc
   DrawText lhDC, "Hy", 2, tTR, DT_CALCRECT
   lLineHeight = tTR.Bottom - tTR.Top
   
'   lCount = txtThis.LineCount
'   lCurrent = txtThis.CurrentLine
   lStart = txtThis.SelStart
   lEnd = txtThis.SelStart + txtThis.SelLength - 1
   If (lEnd > lStart) Then
      lEndLine = txtThis.LineForCharacterIndex(lEnd)
   Else
      lEndLine = lCurrent
   End If
   lLine = txtThis.FirstVisibleLine
   GetClientRect picTo.hwnd, tR
   lEnd = tR.Bottom - tR.Top
      
   hBr = CreateSolidBrush(TranslateColor(picTo.BackColor))
   FillRect lhDC, tR, hBr
   DeleteObject hBr
   tR.Left = 2
   tR.Right = tR.Right - 2
   tR.Top = 0
   tR.Bottom = tR.Top + lLineHeight
   
   SetTextColor lhDC, TranslateColor(vbButtonShadow)
   
   Do
      ' Ensure correct colour:
      If (lLine = lCurrent) Then
         SetTextColor lhDC, TranslateColor(vbWindowText)
      ElseIf (lLine = lEndLine + 1) Then
         SetTextColor lhDC, TranslateColor(vbButtonShadow)
      End If
      ' Draw the line number:
      DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
      
      ' Increment the line:
      lLine = lLine + 1
      ' Increment the position:
      OffsetRect tR, 0, lLineHeight
      If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
         bComplete = True
      End If
   Loop While Not bComplete
   
   ' Draw a line...
   MoveToEx lhDC, tR.Right + 1, 0, tPO
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
   hPenOld = SelectObject(lhDC, hPen)
   LineTo lhDC, tR.Right + 1, lEnd
   SelectObject lhDC, hPenOld
   DeleteObject hPen
   If picTo.AutoRedraw Then
      picTo.Refresh
   End If
   
End Sub

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function
