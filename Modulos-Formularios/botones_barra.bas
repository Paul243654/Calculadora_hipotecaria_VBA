Attribute VB_Name = "botones_barra"
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16)
Private Const WS_SYSMENU As Long = &H80000

Private Sub UserForm_Initialize()
 Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    If Application.Version < 9 Then
        lngMyHandle = FindWindow("THUNDERXFRAME", Me.Caption)
    Else
        lngMyHandle = FindWindow("THUNDERDFRAME", Me.Caption)
    End If
    lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
    lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle
End Sub
