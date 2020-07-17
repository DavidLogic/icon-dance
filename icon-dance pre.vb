Private Const LVM_FIRST As Long = &H1000

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const LVM_SETITEMPOSITION As Long = (LVM_FIRST + 15)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim curPoint As POINTAPI

Dim listViewhwnd As Long

Public Function MAKELPARAM(ByVal l As Integer, ByVal h As Integer) As Long
Dim ll As String
Dim lh As String
Dim r As String
ll = Format(Hex(l), "@@@@")
lh = Format(Hex(h), "@@@@")
Dim result As Long

result = CLng("&h" & Replace(lh & ll, " ", "0"))
 MAKELPARAM = result
 
End Function


Private Function getDesktopHwnd() As Long
Dim hwndWorkerW As Long, hwndShelldll As Long, hwndDesktop As Long

Do While (hwndDesktop = 0)
    hwndWorkerW = FindWindowEx(0, hwndWorkerW, "WorkerW", vbNullString)
        If (hwndWorkerW <> 0) Then
            hwndShelldll = FindWindowEx(hwndWorkerW, 0, "SHELLDLL_DefView", vbNullString)
   
            hwndDesktop = FindWindowEx(hwndShelldll, 0, "SysListView32", vbNullString)
        End If
Loop
getDesktopHwnd = hwndDesktop
End Function
Private Sub Form_Load()
listViewhwnd = getDesktopHwnd()

End Sub

Private Sub Timer1_Timer()
curPoint.x = curPoint.x + 10

PostMessage listViewhwnd, LVM_SETITEMPOSITION, 10, MAKELPARAM(curPoint.x, 110)
End Sub
