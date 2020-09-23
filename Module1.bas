Attribute VB_Name = "Module1"
'// does it really get any easier... special thanks to allapi.com
'// and its easy to mod to intercept any WM mouse event
'// just remember, dont exit the form by stopping it in the ide... knuckle head
'// yeah... and meatwad says hi
Option Explicit
Private Const TME_CANCEL = &H80000000
Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const TME_NONCLIENT = &H10&
Private Const TME_QUERY = &H40000000
Private Const WM_MOUSELEAVE = &H2A3&
Private Const WM_LBUTTONDBLCLK As Integer = &H203
Private Const WM_LBUTTONDOWN As Integer = &H201
Private Const WM_LBUTTONUP  As Integer = &H202
Private Const WM_MBUTTONDBLCLK  As Integer = &H209
Private Const WM_MBUTTONDOWN  As Integer = &H207
Private Const WM_MBUTTONUP  As Integer = &H208
Private Const WM_MOUSEACTIVATE  As Integer = &H21
Private Const WM_MOUSEFIRST  As Integer = &H200
Private Const WM_MOUSELAST  As Integer = &H209
Private Const WM_MOUSEMOVE  As Integer = &H200
Private Const WM_RBUTTONDBLCLK  As Integer = &H206
Private Const WM_RBUTTONDOWN  As Integer = &H204
Private Const WM_RBUTTONUP  As Integer = &H205

Private Type TMET
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Declare Function TrackMouseEvent2 Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TMET) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC As Long = (-4)
Private PrevProc As Long
Private ET As TMET

Public Sub Hook(F As Variant)

    PrevProc = SetWindowLong(F.hWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub UnHook(F As Variant)

    SetWindowLong F.hWnd, GWL_WNDPROC, PrevProc

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    With Form1
        If uMsg = WM_MOUSELEAVE Then

            If hWnd = .Command1.hWnd Then
                .Command1.Caption = "mouseleave"
                .Command1.BackColor = vbRed
              ElseIf hWnd = .Command2.hWnd Then 'NOT HWND...
                .Command2.Caption = "mouseleave"
                .Command2.BackColor = vbRed
              ElseIf hWnd = .Command3.hWnd Then 'NOT HWND...
                .Command3.Caption = "mouseleave"
                .Command3.BackColor = vbRed
              ElseIf hWnd = .Command4.hWnd Then 'NOT HWND...
                .Command4.Caption = "mouseleave"
                .Command4.BackColor = vbRed
              ElseIf hWnd = .hWnd Then 'NOT HWND...
                .Caption = "don't hit stop!"
                .BackColor = vbRed
            End If
        End If
    End With 'FORM1
    WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)

End Function

Public Sub mouseMoveHook(Object As Variant)

    ET.cbSize = Len(ET)
    ET.hwndTrack = Object.hWnd
    ET.dwFlags = TME_LEAVE
    TrackMouseEvent2 ET

End Sub

