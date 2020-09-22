Attribute VB_Name = "SizeListView"
Option Explicit
' Known possible bugs:
' If the ListView control is set to report mode, AND an indent value is specified
' for any item, AND the computer has v4.70 or v4.72 of COMMCTL32.DLL installed,
' this will not work.
' Fix: Upgrade COMMCTL32.DLL
'   See: http://support.microsoft.com/default.aspx?scid=kb;EN-US;q246364
' If only some of the items in the ListView are bold, but the actual font for
'   the control is not set to bold, the column will resize to the correct width
'   for the items as if they were not bold.
' Fix: Set the font of the entire ListView control to bold, THEN call
'   LVM_SETCOLUMNWIDTH, then unset the bold font of the entire ListView control.
'   If you know of a better way, please contact me via my website in the about
'   dialog or post a reply on Planet Source Code.

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const LVM_FIRST As Long = &H1000
Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE As Long = -1
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Sub SizeColumn(ColumnIndex As Integer, lvw As ListView)
  'Lock update of ListView. Prevents ghostly text from appearing. I have seen it
  'happen in other projects, but not this one. Always a good idea to use nonetheless.
  LockWindowUpdate lvw.hWnd
  SendMessage lvw.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, LVSCW_AUTOSIZE_USEHEADER
  LockWindowUpdate 0
End Sub
