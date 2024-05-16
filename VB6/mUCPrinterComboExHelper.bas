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
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As VirtualKeyCodes) As Integer

#End If
Private VTableIPAO(0 To 9) As LongPtr, VTableIPAOData As VTableIPAODataStruct
Public Enum VTableInterfaceConstants
VTableInterfaceInPlaceActiveObject = 1
VTableInterfaceControl = 2
VTableInterfacePerPropertyBrowsing = 3
End Enum
Private Type VTableIPAODataStruct
VTable As LongPtr
RefCount As Long
OriginalIOleIPAO As OLEGuids.IOleInPlaceActiveObject
IOleIPAO As OLEGuids.IOleInPlaceActiveObjectVB
End Type
Public Function ucPrinterComboWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterComboWndProc = dwRefData.zzCBWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterComboCWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterComboCWndProc = dwRefData.zzCBCWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterComboLBWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterComboLBWndProc = dwRefData.zzCBLBWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
Public Function ucPrinterComboEditWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As ucPrinterComboEx) As LongPtr
    ucPrinterComboEditWndProc = dwRefData.zzCBEditWndProc(hWnd, uMsg, wParam, lParam, uIdSubclass)
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



Public Function ucPrinterComboSetVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableHandlingSupported(This, VTableInterfaceInPlaceActiveObject) = True Then
            VTableIPAOData.RefCount = VTableIPAOData.RefCount + 1
            ucPrinterComboSetVTableHandling = True
        End If
    ' Case VTableInterfaceControl
    '     If VTableHandlingSupported(This, VTableInterfaceControl) = True Then
    '         Call ReplaceIOleControl(This)
    '         SetVTableHandling = True
    '     End If
    ' Case VTableInterfacePerPropertyBrowsing
    '     If VTableHandlingSupported(This, VTableInterfacePerPropertyBrowsing) = True Then
    '         Call ReplaceIPPB(This)
    '         SetVTableHandling = True
    '     End If
End Select
End Function

Public Function ucPrinterComboRemoveVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableHandlingSupported(This, VTableInterfaceInPlaceActiveObject) = True Then
            VTableIPAOData.RefCount = VTableIPAOData.RefCount - 1
            ucPrinterComboRemoveVTableHandling = True
        End If
    ' Case VTableInterfaceControl
    '     If VTableHandlingSupported(This, VTableInterfaceControl) = True Then
    '         Call RestoreIOleControl(This)
    '         RemoveVTableHandling = True
    '     End If
    ' Case VTableInterfacePerPropertyBrowsing
    '     If VTableHandlingSupported(This, VTableInterfacePerPropertyBrowsing) = True Then
    '         Call RestoreIPPB(This)
    '         RemoveVTableHandling = True
    '     End If
End Select
End Function


Private Function VTableHandlingSupported(ByRef This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
On Error GoTo CATCH_EXCEPTION
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        Dim ShadowIOleIPAO As IOleInPlaceActiveObject
        Dim ShadowIOleInPlaceActiveObjectVB As IOleInPlaceActiveObjectVB
        Set ShadowIOleIPAO = This
        Set ShadowIOleInPlaceActiveObjectVB = This
        VTableHandlingSupported = Not CBool(ShadowIOleIPAO Is Nothing Or ShadowIOleInPlaceActiveObjectVB Is Nothing)
    ' Case VTableInterfaceControl
    '     Dim ShadowIOleControl As IOleControl
    '     Dim ShadowIOleControlVB As IOleControlVB
    '     Set ShadowIOleControl = This
    '     Set ShadowIOleControlVB = This
    '     VTableHandlingSupported = Not CBool(ShadowIOleControl Is Nothing Or ShadowIOleControlVB Is Nothing)
    ' Case VTableInterfacePerPropertyBrowsing
    '     Dim ShadowIPPB As IPerPropertyBrowsing
    '     Dim ShadowIPerPropertyBrowsingVB As IPerPropertyBrowsingVB
    '     Set ShadowIPPB = This
    '     Set ShadowIPerPropertyBrowsingVB = This
    '     VTableHandlingSupported = Not CBool(ShadowIPPB Is Nothing Or ShadowIPerPropertyBrowsingVB Is Nothing)
End Select
CATCH_EXCEPTION:
End Function
Public Sub ucPrinterComboActivateIPAO(ByVal This As Object)
On Error GoTo CATCH_EXCEPTION
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PropOleInPlaceActiveObject As OLEGuids.IOleInPlaceActiveObject
Dim PosRect As OLERECT
Dim ClipRect As OLERECT
Dim FrameInfo As OLEINPLACEFRAMEINFO
Set PropOleObject = This
If VTableIPAOData.RefCount > 0 Then
    With VTableIPAOData
    .VTable = GetVTableIPAO()
    Set .OriginalIOleIPAO = This
    Set .IOleIPAO = This
    End With
    CopyMemory ByVal VarPtr(PropOleInPlaceActiveObject), VarPtr(VTableIPAOData), LenB(VTableIPAOData.VTable)
    PropOleInPlaceActiveObject.AddRef
Else
    Set PropOleInPlaceActiveObject = This
End If
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject PropOleInPlaceActiveObject, 0
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject PropOleInPlaceActiveObject, 0
Exit Sub
CATCH_EXCEPTION:
Debug.Print "ucPrinterComboActivateIPAO->Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub ucPrinterComboDeActivateIPAO()
On Error GoTo CATCH_EXCEPTION
If VTableIPAOData.OriginalIOleIPAO Is Nothing Then Exit Sub
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PosRect As OLERECT
Dim ClipRect As OLERECT
Dim FrameInfo As OLEINPLACEFRAMEINFO
Set PropOleObject = VTableIPAOData.OriginalIOleIPAO
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject Nothing, 0
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject Nothing, 0
CATCH_EXCEPTION:
Set VTableIPAOData.OriginalIOleIPAO = Nothing
Set VTableIPAOData.IOleIPAO = Nothing
End Sub

Private Function ProcPtr(ByVal Address As LongPtr) As LongPtr
    ProcPtr = Address
End Function

Private Function GetVTableIPAO() As LongPtr
If VTableIPAO(0) = 0 Then
    VTableIPAO(0) = ProcPtr(AddressOf IOleIPAO_QueryInterface)
    VTableIPAO(1) = ProcPtr(AddressOf IOleIPAO_AddRef)
    VTableIPAO(2) = ProcPtr(AddressOf IOleIPAO_Release)
    VTableIPAO(3) = ProcPtr(AddressOf IOleIPAO_GetWindow)
    VTableIPAO(4) = ProcPtr(AddressOf IOleIPAO_ContextSensitiveHelp)
    VTableIPAO(5) = ProcPtr(AddressOf IOleIPAO_TranslateAccelerator)
    VTableIPAO(6) = ProcPtr(AddressOf IOleIPAO_OnFrameWindowActivate)
    VTableIPAO(7) = ProcPtr(AddressOf IOleIPAO_OnDocWindowActivate)
    VTableIPAO(8) = ProcPtr(AddressOf IOleIPAO_ResizeBorder)
    VTableIPAO(9) = ProcPtr(AddressOf IOleIPAO_EnableModeless)
End If
GetVTableIPAO = VarPtr(VTableIPAO(0))
End Function

Private Function IOleIPAO_QueryInterface(ByRef This As VTableIPAODataStruct, ByRef IID As OLECLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = 0 Then
    IOleIPAO_QueryInterface = E_POINTER
    Exit Function
End If
' IID_IOleInPlaceActiveObject = {00000117-0000-0000-C000-000000000046}
If IID.Data1 = &H117 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
    If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
    And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
        pvObj = VarPtr(This)
        IOleIPAO_AddRef This
        IOleIPAO_QueryInterface = S_OK
    Else
        IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
    End If
Else
    IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
End If
End Function

Private Function IOleIPAO_AddRef(ByRef This As VTableIPAODataStruct) As Long
IOleIPAO_AddRef = This.OriginalIOleIPAO.AddRef
End Function

Private Function IOleIPAO_Release(ByRef This As VTableIPAODataStruct) As Long
IOleIPAO_Release = This.OriginalIOleIPAO.Release
End Function

Private Function IOleIPAO_GetWindow(ByRef This As VTableIPAODataStruct, ByRef hWnd As LongPtr) As Long
IOleIPAO_GetWindow = This.OriginalIOleIPAO.GetWindow(hWnd)
End Function

Private Function IOleIPAO_ContextSensitiveHelp(ByRef This As VTableIPAODataStruct, ByVal EnterMode As Long) As Long
IOleIPAO_ContextSensitiveHelp = This.OriginalIOleIPAO.ContextSensitiveHelp(EnterMode)
End Function

Private Function IOleIPAO_TranslateAccelerator(ByRef This As VTableIPAODataStruct, ByRef Msg As Msg) As Long
Debug.Print "IOleIPAO_TranslateAccelerator"
If VarPtr(Msg) = 0 Then
    IOleIPAO_TranslateAccelerator = E_INVALIDARG
    Exit Function
End If
On Error GoTo CATCH_EXCEPTION
Dim Handled As Boolean
IOleIPAO_TranslateAccelerator = S_OK
This.IOleIPAO.TranslateAccelerator Handled, IOleIPAO_TranslateAccelerator, Msg.hWnd, Msg.message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
If Handled = False Then IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
Exit Function
CATCH_EXCEPTION:
Debug.Print "IOleIPAO_TranslateAccelerator->Error " & Err.Number & ": " & Err.Description
IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
End Function

Private Function IOleIPAO_OnFrameWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Long) As Long
IOleIPAO_OnFrameWindowActivate = This.OriginalIOleIPAO.OnFrameWindowActivate(Activate)
End Function

Private Function IOleIPAO_OnDocWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Long) As Long
IOleIPAO_OnDocWindowActivate = This.OriginalIOleIPAO.OnDocWindowActivate(Activate)
End Function

Private Function IOleIPAO_ResizeBorder(ByRef This As VTableIPAODataStruct, ByRef rc As OLERECT, ByVal UIWindow As IOleInPlaceUIWindow, ByVal FrameWindow As Long) As Long
IOleIPAO_ResizeBorder = This.OriginalIOleIPAO.ResizeBorder(VarPtr(rc), UIWindow, FrameWindow)
End Function

Private Function IOleIPAO_EnableModeless(ByRef This As VTableIPAODataStruct, ByVal Enable As Long) As Long
IOleIPAO_EnableModeless = This.OriginalIOleIPAO.EnableModeless(Enable)
End Function

Private Function GetShiftStateFromMsg() As ShiftConstants
If GetKeyState(vbKeyShift) < 0 Then GetShiftStateFromMsg = vbShiftMask
If GetKeyState(vbKeyControl) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbCtrlMask
If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbAltMask
End Function

