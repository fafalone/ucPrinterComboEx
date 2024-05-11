Attribute VB_Name = "mUCPrinterComboExHelper"
Option Explicit
Public gUCPrinterHookInst As ucPrinterComboEx
Public gUCPrinterHookWindow As LongPtr
Public gUCPrinterHookHandle As LongPtr
#If TWINBASIC = 0 Then
Private Type MOUSEHOOKSTRUCT
    pt As Point
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_XBUTTONDBLCLK = &H20D
Private Const WM_XBUTTONDOWN = &H20B
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
#End If
Public Function ucPrinterComboWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterComboWndProc = dwRefData.zzCBWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterUserControlWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterUserControlWndProc = dwRefData.zzUCWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterLVWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterLVWndProc = dwRefData.zzLVWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterMouseHookProc(ByVal ncode As Long, ByVal wParam As LongPtr, lParam As MOUSEHOOKSTRUCT) As LongPtr
    If (lParam.hWnd <> gUCPrinterHookWindow) And (GetParent(lParam.hWnd) <> gUCPrinterHookWindow) Then
        Select Case wParam
            Case WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, WM_RBUTTONDBLCLK, WM_RBUTTONDOWN, WM_MBUTTONDBLCLK, WM_MBUTTONDOWN, WM_XBUTTONDBLCLK, WM_XBUTTONDOWN, _
                    WM_NCLBUTTONDBLCLK, WM_NCLBUTTONDOWN, WM_NCRBUTTONDBLCLK, WM_NCRBUTTONDOWN, WM_NCMBUTTONDBLCLK, WM_NCMBUTTONDOWN, WM_NCXBUTTONDBLCLK, WM_NCXBUTTONDOWN
                gUCPrinterHookInst.CloseDropdown lParam.hWnd
        End Select
    End If
    ucPrinterMouseHookProc = CallNextHookEx(gUCPrinterHookHandle, ncode, wParam, lParam)
End Function

