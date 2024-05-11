VERSION 5.00
Begin VB.UserControl ucPrinterComboEx 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   202
   ToolboxBitmap   =   "ucPrinterComboEx.ctx":0000
End
Attribute VB_Name = "ucPrinterComboEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
' ucPrinterCombo v1.0 ***BETA***
' by Jon Johnson (fafalone)
' https://github.com/fafalone/ucPrinterComboEx
'
' This control is designed to display a list of printers as both a
' normal ComboBoxEx (or ImageCombo), with a single line and small
' icons, and a large-icon double-line view found in many other printer
' selection dialogs today.
'
' The large icon view is done by substituting a ListView in tile view
' for the normal ListBox, and accordingly is controlled by the option
' UseListView, default True.
'
' Requirements:
'    .ctl/.tbcontrol form:
'      twinBASIC: Windows Development Library for twinBASIC v7.6+
'                 ucPrinterComboEx.tbcontrol/.twin
'      VB6: oleexp v6.0+
'           mIID.bas
'           ucPrinterComboEx.ctl/.ctx
'           mUCPrinterComboExHelper.bas
'
'   .ocx form: None.
'********************************************************************
Private Const dbg_PrintToImmediate As Boolean = True 'This control has very extensive debug information, you may not want
                                                      'to see that in your IDE.
Private Const dbg_IncludeDate As Boolean = False 'Prefix all Debug output with the date and time, [yyyy-mm-dd Hh:Mm:Ss]
Private Const dbg_IncludeName As Boolean = True 'Include Ambient.Name
Private Const dbg_dtFormat As String = "yyyy-mm-dd Hh:nn:Ss"
Private Const dbg_VerbosityLevel As Long = 3   'Only log to immediate/file messages <= this level






Public Event PrinterChanged(ByVal sNewPrinterName As String, ByVal sParsingPath As String, ByVal sModelName As String, ByVal sNetworkLocation As String, ByVal sLastStatusMessage As String, ByVal bIsDefaultPrinter As Boolean)
Attribute PrinterChanged.VB_MemberFlags = "200"
Public Event DropdownOpen()
Public Event DropdownClose()

    
Private mInit As Boolean
 
Private hCombo As LongPtr
Private hComboEd As LongPtr
Private hComboCB As LongPtr
Private hLVW As LongPtr, bLVVis As Boolean
Private hFont As LongPtr, hFontBold As LongPtr
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private mIFMain As IFont
Private hParOrig As LongPtr
Private pImlSmall As IImageList, himlSmall As LongPtr
Private pImlLarge As IImageList, himlLarge As LongPtr
Private himlMain As LongPtr
Private Type SysImgCacheEntry
    sysimlid As Long
    limlidx As Long
End Type
Private SysImgCache() As SysImgCacheEntry
Private nSysImgCache As Long
Private bOvrAdded(16) As Boolean
Private hTheme As LongPtr
Private mLastHT As Long

Private mMouseDown As Boolean 'Tracked by LV drop only
Private bFlagSuppressReopen As Boolean

Private mDPI As Single
Private mActualZoom As Single 'Get actual DPI even if virtualized
Private IsComCtl6 As Boolean
Private smCXEdge As Long, smCYEdge As Long


Private Const sCol0 = "Name"
Private Const sCol1 = "Status"

Private Type tPrinter
    sName As String
    sParsingPath As String
    sInfoTip As String
    sModel As String
    sLocation As String
    sLastStatus As String
    nIcon As Long
    nIconLV As Long
    nOvr As Long
    lvi As Long
    cbi As Long
    bDefault As Boolean
End Type
Private mPrinters() As tPrinter
Private nPr As Long
Private mPrintersOld() As tPrinter
Private nPrOld As Long
Private mIdxDef As Long
Private mIdxSel As Long
Private mIdxSelPrev As Long
Private mLabelSel As String
Private mLabelSelPrev As String

Private mRaiseOnLoad As Boolean
Private Const mDefRaiseOnLoad As Boolean = True

Private cyList As Long
Private Const mDefCY As Long = 800

Private cxList As Long
Private Const mDefCX As Long = 0 'scaLed by DPI in UC_Init

Private cxyIcon As Long
Private Const mDefIcon As Long = 32

Private mNotify As Boolean
Private Const mDefNotify As Boolean = True

Private mEnabled As Boolean
Private Const mDefEnabled As Boolean = True

Private mListView As Boolean
Private Const mDefListView As Boolean = True

Private mTrack As Boolean
Private Const mDefTrack As Boolean = True

Private mLimitCX As Boolean
Private Const mDefLimitCX As Boolean = False

Private mNoRf As Boolean
Private Const mDefNoRf As Boolean = False

Private mBk As OLE_COLOR
Private Const mDefBk As Long = &H8000000F

#If TWINBASIC Then
[EnumId("55209AC8-57EA-4644-AA85-4974AA31E101")]
#End If
Public Enum UCPCType
    UCPC_DropdownList
    UCPC_Combo
End Enum
Private mStyle As UCPCType
Private Const mDefStyle As Long = 0

'We don't need the .Flags argument so we can skip the whole
'issue with packing alignment padding bytes, but we don't
'want to read past the end of the struct in VB6
#If Win64 Then
Private Const cbnmlvkd = &H1E
#Else
Private Const cbnmlvkd = &H12
#End If



'Only a VB6 set of APIs is given here because they're covered by WinDevLib,
'the dependency for the interfaces, in twinBASIC.
'Also covered are SDK macros and helpers provided by WDL where the 64bit
'version is different.
#If TWINBASIC = 0 Then
    
    

     Private Const vbNullPtr = 0
     
     Private Const CTRUE = 1
     Private Const CFALSE = 0
    
     Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As WindowStylesEx, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As WindowStyles, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
     Private Declare Function SHLoadNonloadedIconOverlayIdentifiers Lib "shell32" () As Long
     Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As BOOL
     Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As BOOL
     Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As SWP_Flags) As Long
     Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As BOOL) As BOOL
     Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwTime As Long, ByVal dwFlags As AnimateWindowFlags) As Long
     Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As ShowWindow) As Long
     Private Declare Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
     Private Declare Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
     Private Declare Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, Optional ByVal hWndNewParent As LongPtr) As LongPtr
     Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uiAction As SPI, ByVal uiParam As Long, ByRef pvParam As Any, ByVal fWinIni As SPIF) As BOOL
     Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
     Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
     Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
     Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
     Private Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal pszSubAppName As LongPtr, ByVal pszSubIdList As LongPtr) As Long
     Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal pszClassList As LongPtr) As LongPtr
     Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As LongPtr) As Long
     Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
     Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As BOOL) As BOOL
     Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
     Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
     Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
     Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As LongPtr
     Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
     Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
     Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
     Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal Flags As RDW_Flags) As Long
     Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal idHook As WindowsHookCodes, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
     Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
     Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
     Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hDC As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, lpSize As Size) As BOOL
     Private Declare Function GetWindowsDirectoryW Lib "kernel32" (ByVal lpBuffer As LongPtr, ByVal nSize As Long) As Long
     Private Declare Function SHGetFileInfoW Lib "shell32" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFOW, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As LongPtr
     Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As LongPtr, ByVal nFolder As CSIDLs, ByVal hToken As LongPtr, ByVal dwReserved As Long, ppidl As LongPtr) As Long
     Private Declare Function PSGetPropertyDescription Lib "propsys" (PropKey As PROPERTYKEY, riid As UUID, ppv As Any) As Long
     Private Declare Function PropVariantToVariant Lib "propsys" (ByRef propvar As Any, ByRef var As Variant) As Long
     Private Declare Function PSFormatPropertyValue Lib "propsys" (ByVal pps As LongPtr, ByVal ppd As LongPtr, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As LongPtr) As Long
     Private Declare Function GetDefaultPrinterW Lib "winspool.drv" (ByVal pszBuffer As LongPtr, pcchBuffer As Long) As BOOL
     Private Declare Function ILCombine Lib "shell32" (ByVal pidl1 As LongPtr, ByVal pidl2 As LongPtr) As LongPtr
     Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
     Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
     Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
     Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
     Private Declare Function CompareMemory Lib "ntdll" Alias "RtlCompareMemory" (Source1 As Any, Source2 As Any, ByVal Length As LongPtr) As LongPtr
     Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
     Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
     Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
     Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
     Private Declare Function ImageList_Add Lib "comctl32" (ByVal himl As LongPtr, ByVal hbmImage As LongPtr, ByVal hBMMask As LongPtr) As Long
     Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal himl As LongPtr, ByVal hbmImage As LongPtr, ByVal crMask As Long) As Long
     Private Declare Function ImageList_Create Lib "comctl32" (ByVal CX As Long, ByVal cy As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As LongPtr
     Private Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal himl As LongPtr, ByVal i As Long, ByVal hIcon As LongPtr) As Long
     Private Declare Function ImageList_SetBkColor Lib "comctl32" (ByVal himl As LongPtr, ByVal clrBk As Long) As Long
     Private Declare Function ImageList_SetOverlayImage Lib "comctl32" (ByVal himl As LongPtr, ByVal iImage As Long, ByVal iOverlay As Long) As Long
     Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal himl As LongPtr) As Long
     Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal clr As stdole.OLE_COLOR, ByVal hpal As LongPtr, lpcolorref As Long) As Long
     Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
     Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As DefaultMonitorValues) As LongPtr
     Private Declare Function GetMonitorInfoW Lib "user32" (ByVal hMonitor As LongPtr, lpmi As Any) As BOOL
     Private Declare Function EnumDisplaySettingsW Lib "user32" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As EnumDispMode, lpDevMode As DEVMODEW) As BOOL
     
     
    Private Const EM_SETREADONLY = &HCF
    Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
    Private Const NFR_UNICODE = 2
    Private Const DESKTOPHORZRES As Long = 118
    Private Const DESKTOPVERTRES As Long = 117
    Private Const LOGPIXELSX = 88   ' Logical pixels/inch in X                 */
    Private Const LOGPIXELSY = 90   ' Logical pixels/inch in Y
    Private Const CCHDEVICENAME = 32
    Private Const USER_DEFAULT_SCREEN_DPI = 96
    
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    Private Enum EnumDispMode
        ENUM_CURRENT_SETTINGS = (-1)
        ENUM_REGISTRY_SETTINGS = (-2)
    End Enum
    
    Private Enum DefaultMonitorValues
        MONITOR_DEFAULTTONULL = &H0
        MONITOR_DEFAULTTOPRIMARY = &H1
        MONITOR_DEFAULTTONEAREST = &H2
    End Enum

    Private Enum MonitorInfoFlags
        MONITORINFOF_PRIMARY = &H1
    End Enum

    Private Type MONITORINFO
        cbSize As Long
        rcMonitor As RECT
        rcWork As RECT
        dwFlags As MonitorInfoFlags
    End Type
    Private Type MONITORINFOEXW
        info As MONITORINFO
        szDevice(0 To (CCHDEVICENAME - 1)) As Integer
    End Type
    
    Private Type DEVMODEW
        dmDeviceName(CCHDEVICENAME - 1) As Integer
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        PrinterOrDisplayFields(0 To 15) As Byte 'union {
                                                '   ' printer only fields */
                                                '   struct {
                                                '     short dmOrientation;
                                                '     short dmPaperSize;
                                                '     short dmPaperLength;
                                                '     short dmPaperWidth;
                                                '     short dmScale;
                                                '     short dmCopies;
                                                '     short dmDefaultSource;
                                                '     short dmPrintQuality;
                                                '   } DUMMYSTRUCTNAME;
                                                '   ' display only fields */
                                                '   struct {
                                                '     POINTL dmPosition;
                                                '     DWORD  dmDisplayOrientation;
                                                '     DWORD  dmDisplayFixedOutput;
                                                '   } DUMMYSTRUCTNAME2;
                                                ' } DUMMYUNIONNAME;
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName(32 - 1) As Integer
        dmLogPixels As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        DisplayFlagsOrNup As Long   ' union {
                                    '     DWORD  dmDisplayFlags;
                                    '     DWORD  dmNup;
                                    ' } DUMMYUNIONNAME2;
        dmDisplayFrequency As Long
        '#if(WINVER >= 0x0400)
        dmICMMethod As Long
        dmICMIntent As Long
        dmMediaType As Long
        dmDitherType As Long
        dmReserved1 As Long
        dmReserved2 As Long
        '#if (WINVER >= 0x0500) || (_WIN32_WINNT >= _WIN32_WINNT_NT4)
        dmPanningWidth As Long
        dmPanningHeight As Long
        '#endif
        '#endif /* WINVER >= 0x0400 */
    End Type
    
   Private Enum IL_CreateFlags
     ILC_MASK = &H1
     ILC_COLOR = &H0
     ILC_COLORDDB = &HFE
     ILC_COLOR4 = &H4
     ILC_COLOR8 = &H8
     ILC_COLOR16 = &H10
     ILC_COLOR24 = &H18
     ILC_COLOR32 = &H20
     ILC_PALETTE = &H800                  ' (no longer supported...never worked anyway)
     '5.0
     ILC_MIRROR = &H2000
     ILC_PERITEMMIRROR = &H8000&
     '6.0
     ILC_ORIGINALSIZE = &H10000
     ILC_HIGHQUALITYSCALE = &H20000
   End Enum
   
   Private Enum SHGFI_flags
     SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
     SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
     SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
     SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
     SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
     ' Indicates that the function should not attempt to access the file specified by pszPath.
     ' Rather, it should act as if the file specified by pszPath exists with the file attributes
     ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
     ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
     SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
     SHGFI_ADDOVERLAYS = &H20
     SHGFI_OVERLAYINDEX = &H40 'Return overlay index in upper 8 bits of iIcon.
     SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
     SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
     SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
     SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
     SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                           ' containing the icon, rtns BOOL
     SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
     SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
     SHGFI_LINKOVERLAY = &H8000&    ' add shortcut overlay to sfi.hIcon
     SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
     SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
   End Enum
   Private Type SHFILEINFOW   ' shfi
     hIcon As LongPtr
     iIcon As Long
     dwAttributes As Long
     szDisplayName(MAX_PATH - 1) As Integer
     szTypeName(79) As Integer
   End Type
   
   Private Const CDRF_DODEFAULT As Long = &H0
   Private Const CDRF_NEWFONT As Long = &H2
   Private Const CDRF_SKIPDEFAULT As Long = &H4
   Private Const CDRF_DOERASE As Long = &H8
   Private Const CDRF_NOTIFYPOSTPAINT As Long = &H10
   Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
   Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20
   Private Const CDRF_NOTIFYPOSTERASE As Long = &H40

   Private Const CDDS_PREPAINT = &H1&
   Private Const CDDS_POSTPAINT = &H2&
   Private Const CDDS_PREERASE = &H3&
   Private Const CDDS_POSTERASE = &H4&
   Private Const CDDS_ITEM As Long = &H10000
   Private Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
   Private Const CDDS_ITEMPOSTPAINT = CDDS_ITEM Or CDDS_POSTPAINT
   Private Const CDDS_ITEMPREERASE As Long = CDDS_ITEM Or CDDS_PREERASE
   Private Const CDDS_ITEMPOSTERASE = CDDS_ITEM Or CDDS_POSTERASE
   Private Const CDDS_SUBITEM = &H20000

 
   Private Enum CustomDrawItemStates
       CDIS_SELECTED = &H1
       CDIS_GRAYED = &H2
       CDIS_DISABLED = &H4
       CDIS_CHECKED = &H8
       CDIS_FOCUS = &H10
       CDIS_DEFAULT = &H20
       CDIS_HOT = &H40
       CDIS_MARKED = &H80
       CDIS_INDETERMINATE = &H100
       CDIS_SHOWKEYBOARDCUES = &H200
       CDIS_NEARHOT = &H400
       CDIS_OTHERSIDEHOT = &H800
       CDIS_DROPHILITED = &H1000
   End Enum
   
    Private Const WC_COMBOBOXW = "ComboBox"
    Private Const WC_COMBOBOX = WC_COMBOBOXW

    Private Enum SWP_Flags
        SWP_NOSIZE = &H1
        SWP_NOMOVE = &H2
        SWP_NOZORDER = &H4
        SWP_NOREDRAW = &H8
        SWP_NOACTIVATE = &H10
        SWP_FRAMECHANGED = &H20
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        SWP_SHOWWINDOW = &H40
        SWP_HIDEWINDOW = &H80
        SWP_NOCOPYBITS = &H100
        SWP_NOOWNERZORDER = &H200
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSENDCHANGING = &H400
    
        SWP_DEFERERASE = &H2000
        SWP_ASYNCWINDOWPOS = &H4000
    End Enum
    Private Enum WindowZOrderDefaults
        HWND_DESKTOP = 0&
        HWND_TOP = 0&
        HWND_BOTTOM = 1&
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
    End Enum
    Private Enum AnimateWindowFlags
        AW_HOR_POSITIVE = &H1
        AW_HOR_NEGATIVE = &H2
        AW_VER_POSITIVE = &H4
        AW_VER_NEGATIVE = &H8
        AW_CENTER = &H10
        AW_HIDE = &H10000
        AW_ACTIVATE = &H20000
        AW_SLIDE = &H40000
        AW_BLEND = &H80000
    End Enum
    Private Enum WindowsHookCodes
        WH_MIN = (-1)
        WH_MSGFILTER = (-1)
        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14
        WH_MAX = 14
    End Enum
    Private Enum HookCodes
        HC_ACTION = 0
        HC_GETNEXT = 1
        HC_SKIP = 2
        HC_NOREMOVE = 3
        HC_NOREM = HC_NOREMOVE
        HC_SYSMODALON = 4
        HC_SYSMODALOFF = 5
    End Enum
    Private Enum RDW_Flags
        RDW_INVALIDATE = &H1
        RDW_INTERNALPAINT = &H2
        RDW_ERASE = &H4
        RDW_VALIDATE = &H8
        RDW_NOINTERNALPAINT = &H10
        RDW_NOERASE = &H20
        RDW_NOCHILDREN = &H40
        RDW_ALLCHILDREN = &H80
        RDW_UPDATENOW = &H100
        RDW_ERASENOW = &H200
        RDW_FRAME = &H400
        RDW_NOFRAME = &H800
    End Enum
    
    
    Private Const LF_FACESIZE = 32
    Private Enum FontWeight
        FW_DONTCARE = 0
        FW_THIN = 100
        FW_EXTRALIGHT = 200
        FW_LIGHT = 300
        FW_NORMAL = 400
        FW_MEDIUM = 500
        FW_SEMIBOLD = 600
        FW_BOLD = 700
        FW_EXTRABOLD = 800
        FW_HEAVY = 900
        FW_ULTRALIGHT = FW_EXTRALIGHT
        FW_REGULAR = FW_NORMAL
        FW_DEMIBOLD = FW_SEMIBOLD
        FW_ULTRABOLD = FW_EXTRABOLD
        FW_BLACK = FW_HEAVY
    End Enum
    Private Type LOGFONT
        LFHeight As Long
        LFWidth As Long
        LFEscapement As Long
        LFOrientation As Long
        LFWeight As FontWeight
        LFItalic As Byte
        LFUnderline As Byte
        LFStrikeOut As Byte
        LFCharset As Byte
        LFOutPrecision As Byte
        LFClipPrecision As Byte
        LFQuality As Byte
        LFPitchAndFamily As Byte
        LFFaceName(LF_FACESIZE - 1) As Integer
    End Type
    
    
    Private Const NM_FIRST As Long = 0
    Private Const NM_OUTOFMEMORY As Long = (NM_FIRST - 1)
    Private Const NM_CLICK As Long = (NM_FIRST - 2) 'uses NMCLICK struct
    Private Const NM_DBLCLK As Long = (NM_FIRST - 3)
    Private Const NM_RETURN As Long = (NM_FIRST - 4)
    Private Const NM_RCLICK As Long = (NM_FIRST - 5) 'uses NMCLICK struct
    Private Const NM_RDBLCLK As Long = (NM_FIRST - 6)
    Private Const NM_SETFOCUS As Long = (NM_FIRST - 7)
    Private Const NM_KILLFOCUS As Long = (NM_FIRST - 8)
    Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
    
    Private Type NMLISTVIEW   ' was NM_LISTVIEW
      hdr As NMHDR
      iItem As Long
      iSubItem As Long
      uNewState As LVITEM_state
      uOldState As LVITEM_state
      uChanged As LVITEM_mask
      PTAction As Point
      lParam As LongPtr
    End Type
    Private Type NMLVKEYDOWN   ' was LV_KEYDOWN
       hdr As NMHDR
       wVKey As Integer   ' can't be KeyCodeConstants, enums are Longs!
       Flags As Long   ' Always zero.
    End Type
    
    Private Const WC_LISTVIEWW = "SysListView32"
    Private Const WC_LISTVIEW = WC_LISTVIEWW
 
    Private Enum LVStyles
      LVS_ICON = &H0
      LVS_REPORT = &H1
      LVS_SMALLICON = &H2
      LVS_LIST = &H3
      LVS_TYPEMASK = &H3
      LVS_SINGLESEL = &H4
      LVS_SHOWSELALWAYS = &H8
      LVS_SORTASCENDING = &H10
      LVS_SORTDESCENDING = &H20
      LVS_SHAREIMAGELISTS = &H40
      LVS_NOLABELWRAP = &H80
      LVS_AUTOARRANGE = &H100
      LVS_EDITLABELS = &H200
      LVS_OWNERDRAWFIXED = &H400
      LVS_ALIGNLEFT = &H800
      LVS_OWNERDATA = &H1000
      LVS_NOSCROLL = &H2000
      LVS_NOCOLUMNHEADER = &H4000
      LVS_NOSORTHEADER = &H8000&
      LVS_TYPESTYLEMASK = &HFC00&
      LVS_ALIGNTOP = &H0
      LVS_ALIGNRIGHT = &HC00 'UNDOCUMENTED
      LVS_ALIGNMASK = &HC00
    End Enum   ' LVStyles

    Private Enum LVStylesEx
      LVS_EX_GRIDLINES = &H1
      LVS_EX_SUBITEMIMAGES = &H2
      LVS_EX_CHECKBOXES = &H4
      LVS_EX_TRACKSELECT = &H8
      LVS_EX_HEADERDRAGDROP = &H10
      LVS_EX_FULLROWSELECT = &H20         ' // applies to report mode only
      LVS_EX_ONECLICKACTIVATE = &H40
      LVS_EX_TWOCLICKACTIVATE = &H80
      LVS_EX_FLATSB = &H100
      LVS_EX_REGIONAL = &H200             'Not supported on 6.0+ (Vista+)
      LVS_EX_INFOTIP = &H400              ' listview does InfoTips for you
      LVS_EX_UNDERLINEHOT = &H800
      LVS_EX_UNDERLINECOLD = &H1000
      LVS_EX_MULTIWORKAREAS = &H2000
      LVS_EX_LABELTIP = &H4000
      LVS_EX_BORDERSELECT = &H8000&
      LVS_EX_DOUBLEBUFFER = &H10000
      LVS_EX_HIDELABELS = &H20000
      LVS_EX_SINGLEROW = &H40000
      LVS_EX_SNAPTOGRID = &H80000 '// Icons automatically snap to grid.
      LVS_EX_SIMPLESELECT = &H100000        '// Also changes overlay rendering to top right for icon mode.
      LVS_EX_JUSTIFYCOLUMNS = &H200000      '// Icons are lined up in columns that use up the whole view area.
      LVS_EX_TRANSPARENTBKGND = &H400000    '// Background is painted by the parent via WM_PRINTCLIENT
      LVS_EX_TRANSPARENTSHADOWTEXT = &H800000    '// Enable shadow text on transparent backgrounds only (useful with bitmaps)
      LVS_EX_AUTOAUTOARRANGE = &H1000000    '// Icons automatically arrange if no icon positions have been set
      LVS_EX_HEADERINALLVIEWS = &H2000000   '// Display column header in all view modes
      LVS_EX_DRAWIMAGEASYNC = &H4000000     'UNDOCUMENTED. LVN_ASYNCDRAW, NMLVASYNCDRAW
      LVS_EX_AUTOCHECKSELECT = &H8000000
      LVS_EX_AUTOSIZECOLUMNS = &H10000000
      LVS_EX_COLUMNSNAPPOINTS = &H40000000
      LVS_EX_COLUMNOVERFLOW = &H80000000
    End Enum

    ' value returned by many listview messages indicating
    ' the index of no listview item (user defined)
    Private Const LVI_NOITEM = &HFFFFFFFF

    ' messages
    Private Const LVM_FIRST = &H1000
    Private Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
    Private Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
    Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
    Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
    Private Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
    Private Const LVM_GETITEMRECT = (LVM_FIRST + 14)
    Private Const LVM_GETSTRINGWIDTH = (LVM_FIRST + 17)
    Private Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
    Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
    Private Const LVM_SETHOVERTIME = (LVM_FIRST + 71)
    Private Const LVM_GETITEMW = (LVM_FIRST + 75)
    Private Const LVM_SETITEMW = (LVM_FIRST + 76)
    Private Const LVM_INSERTITEMW = (LVM_FIRST + 77)
    Private Const LVM_GETSTRINGWIDTHW = (LVM_FIRST + 87)
    Private Const LVM_GETCOLUMNW = (LVM_FIRST + 95)
    Private Const LVM_SETCOLUMNW = (LVM_FIRST + 96)
    Private Const LVM_INSERTCOLUMNW = (LVM_FIRST + 97)
    Private Const LVM_SETVIEW = (LVM_FIRST + 142)
    Private Const LVM_SETTILEVIEWINFO = (LVM_FIRST + 162)
    Private Const LVM_SETTILEINFO = (LVM_FIRST + 164)
    Private Const LVM_QUERYINTERFACE = (LVM_FIRST + 189)      'UNDOCUMENTED
    
 
    Private Const LVM_GETCOLUMN = LVM_GETCOLUMNW
    Private Const LVM_SETCOLUMN = LVM_SETCOLUMNW
    Private Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNW
    Private Const LVM_GETITEM = LVM_GETITEMW
    Private Const LVM_SETITEM = LVM_SETITEMW
    Private Const LVM_INSERTITEM = LVM_INSERTITEMW

    Private Enum LV_VIEW
        LV_VIEW_ICON = &H0
        LV_VIEW_DETAILS = &H1
        LV_VIEW_SMALLICON = &H2
        LV_VIEW_LIST = &H3
        LV_VIEW_TILE = &H4&
        LV_VIEW_CONTENTS = &H7
    'Below are not part of API, but are implemented by this project.
        LV_VIEW_THUMBNAIL = &H6&
        LV_VIEW_XLICON = &H8
        LV_VIEW_MDICON = &H9
        LV_VIEW_CUSTOM = &H100
    End Enum

    Private Enum LVNI_Flags
        LVNI_ALL = &H0
        LVNI_FOCUSED = &H1
        LVNI_SELECTED = &H2
        LVNI_CUT = &H4
        LVNI_DROPHILITED = &H8
    
        LVNI_ABOVE = &H100
        LVNI_BELOW = &H200
        LVNI_TOLEFT = &H400
        LVNI_TORIGHT = &H800
    '#If (WIN32_IE >= &H600) Then
        LVNI_STATEMASK = (LVNI_FOCUSED Or LVNI_SELECTED Or LVNI_CUT Or LVNI_DROPHILITED)
        LVNI_DIRECTIONMASK = (LVNI_ABOVE Or LVNI_BELOW Or LVNI_TOLEFT Or LVNI_TORIGHT)

        LVNI_PREVIOUS = &H20
        LVNI_VISIBLEORDER = &H10
        LVNI_VISIBLEONLY = &H40
        LVNI_SAMEGROUPONLY = &H80
    '#End If
    End Enum
    
    ' LVM_GETITEMRECT rc.Left (lParam)
    Private Enum LVIR_Flags
        LVIR_BOUNDS = 0
        LVIR_ICON = 1
        LVIR_LABEL = 2
        LVIR_SELECTBOUNDS = 3
    End Enum
    
    Private Enum LVTVI_Flags
        LVTVIF_AUTOSIZE = &H0
        LVTVIF_FIXEDWIDTH = &H1
        LVTVIF_FIXEDHEIGHT = &H2
        LVTVIF_FIXEDSIZE = &H3
        '6.0
        LVTVIF_EXTENDED = &H4
    End Enum
    Private Enum LVTVI_Mask
        LVTVIM_TILESIZE = &H1
        LVTVIM_COLUMNS = &H2
        LVTVIM_LABELMARGIN = &H4
    End Enum
    Private Type LVTILEVIEWINFO
        cbSize As Long
        dwMask As LVTVI_Mask ';     //LVTVIM_*
        dwFlags As LVTVI_Flags ';    //LVTVIF_*
        SizeTile As Size ' ;
        cLines As Long
        RCLabelMargin As RECT
    End Type
    
    Private Type LVTILEINFO
        cbSize As Long
        iItem As Long
        cColumns As Long
        puColumns As LongPtr
    '#if (_WIN32_WINNT >= 0x0600)
        piColFmt As LongPtr
    '#End If
    End Type
    
    Private Type NMCUSTOMDRAW
        hdr As NMHDR
        dwDrawStage As Long
        hDC As LongPtr
        rc As RECT
        dwItemSpec As LongPtr
        uItemState As Long
        lItemlParam As LongPtr
    End Type
    Private Enum LVCD_ItemType
        LVCDI_ITEM = &H0
        LVCDI_GROUP = &H1
        LVCDI_ITEMSLIST = &H2
    End Enum
    Private Type NMLVCUSTOMDRAW
      NMCD As NMCUSTOMDRAW
      ClrText As Long
      ClrTextBk As Long
      ' if IE >= 4.0 this member of the struct can be used
      iSubItem As Integer
      '>=5.01
      dwItemType As LVCD_ItemType
      clrFace As Long
      iIconEffect As Integer
      iIconPhase As Integer
      iPartId As Integer
      iStateId As Integer
      rcText As RECT
      uAlign As Long
    End Type
    
    
    
    ' ============================================
    ' Notifications

    Private Enum LVNotifications
      LVN_FIRST = -100&   ' &HFFFFFF9C   ' (0U-100U)
      LVN_LAST = -199&   ' &HFFFFFF39   ' (0U-199U)
                                                                              ' lParam points to:
      LVN_ITEMCHANGING = (LVN_FIRST - 0)            ' NMLISTVIEW, ?, rtn T/F
      LVN_ITEMCHANGED = (LVN_FIRST - 1)             ' NMLISTVIEW, ?
      LVN_INSERTITEM = (LVN_FIRST - 2)                  ' NMLISTVIEW, iItem
      LVN_DELETEITEM = (LVN_FIRST - 3)                 ' NMLISTVIEW, iItem
      LVN_DELETEALLITEMS = (LVN_FIRST - 4)         ' NMLISTVIEW, iItem = -1, rtn T/F

      LVN_COLUMNCLICK = (LVN_FIRST - 8)              ' NMLISTVIEW, iItem = -1, iSubItem = column
      LVN_BEGINDRAG = (LVN_FIRST - 9)                  ' NMLISTVIEW, iItem
      LVN_BEGINRDRAG = (LVN_FIRST - 11)              ' NMLISTVIEW, iItem

      LVN_ODCACHEHINT = (LVN_FIRST - 13)           ' NMLVCACHEHINT
      LVN_ITEMACTIVATE = (LVN_FIRST - 14)           ' v4.70 = NMHDR, v4.71 = NMITEMACTIVATE
      LVN_ODSTATECHANGED = (LVN_FIRST - 15)  ' NMLVODSTATECHANGE, rtn T/F
      LVN_HOTTRACK = (LVN_FIRST - 21)                 ' NMLISTVIEW, see docs, rtn T/F
      LVN_BEGINLABELEDITA = (LVN_FIRST - 5)        ' NMLVDISPINFO, iItem, rtn T/F
      LVN_ENDLABELEDITA = (LVN_FIRST - 6)           ' NMLVDISPINFO, see docs
 
      LVN_GETDISPINFOA = (LVN_FIRST - 50)            ' NMLVDISPINFO, see docs
      LVN_SETDISPINFOA = (LVN_FIRST - 51)            ' NMLVDISPINFO, see docs
      LVN_ODFINDITEMA = (LVN_FIRST - 52)             ' NMLVFINDITEM
 
      LVN_KEYDOWN = (LVN_FIRST - 55)                 ' NMLVKEYDOWN
      LVN_MARQUEEBEGIN = (LVN_FIRST - 56)       ' NMLISTVIEW, rtn T/F
      LVN_GETINFOTIPA = (LVN_FIRST - 57)             ' NMLVGETINFOTIP
      LVN_GETINFOTIPW = (LVN_FIRST - 58)              ' NMLVGETINFOTIP
      LVN_INCREMENTALSEARCHA = (LVN_FIRST - 62)
      LVN_INCREMENTALSEARCHW = (LVN_FIRST - 63)
    '#If (WIN32_IE >= &H600) Then
      LVN_COLUMNDROPDOWN = (LVN_FIRST - 64)
      LVN_COLUMNOVERFLOWCLICK = (LVN_FIRST - 66)
    '#End If
      LVN_BEGINLABELEDITW = (LVN_FIRST - 75)
      LVN_ENDLABELEDITW = (LVN_FIRST - 76)
      LVN_GETDISPINFOW = (LVN_FIRST - 77)
      LVN_SETDISPINFOW = (LVN_FIRST - 78)
      LVN_ODFINDITEMW = (LVN_FIRST - 79)             ' NMLVFINDITEM
      LVN_BEGINSCROLL = (LVN_FIRST - 80)
      LVN_ENDSCROLL = (LVN_FIRST - 81)
      LVN_LINKCLICK = (LVN_FIRST - 84)
      LVN_ASYNCDRAW = (LVN_FIRST - 86) 'Undocumented; NMLVASYNCDRAW
      LVN_GETEMPTYMARKUP = (LVN_FIRST - 87)
      LVN_GROUPCHANGED = (LVN_FIRST - 88)   ' Undocumented; NMLVGROUP
    'We're going to default to Unicode, but allow targeting ANSI
    #If ANSI = 1 Then
      LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITA
      LVN_ENDLABELEDIT = LVN_ENDLABELEDITA
      LVN_GETDISPINFO = LVN_GETDISPINFOA
      LVN_SETDISPINFO = LVN_SETDISPINFOA
      LVN_ODFINDITEM = LVN_ODFINDITEMA         ' NMLVFINDITEM
      LVN_GETINFOTIP = LVN_GETINFOTIPA              ' NMLVGETINFOTIP
      LVN_INCREMENTALSEARCH = LVN_INCREMENTALSEARCHA
    #Else
      LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITW
      LVN_ENDLABELEDIT = LVN_ENDLABELEDITW
      LVN_GETDISPINFO = LVN_GETDISPINFOW
      LVN_SETDISPINFO = LVN_SETDISPINFOW
      LVN_ODFINDITEM = LVN_ODFINDITEMW         ' NMLVFINDITEM
      LVN_GETINFOTIP = LVN_GETINFOTIPW              ' NMLVGETINFOTIP
      LVN_INCREMENTALSEARCH = LVN_INCREMENTALSEARCHW
    #End If
    End Enum   ' LVNotifications


    ' LVM_GET/SETIMAGELIST wParam

    Private Enum LV_ImageList
        LVSIL_NORMAL = 0
        LVSIL_SMALL = 1
        LVSIL_STATE = 2
        LVSIL_GROUPHEADER = 3
        LVSIL_FOOTER = 4 'UNDOCUMENTED: For footer items... see IListViewFooter
    End Enum

    Private Enum LVITEM_mask
      LVIF_TEXT = &H1
      LVIF_IMAGE = &H2
      LVIF_PARAM = &H4
      LVIF_STATE = &H8
      LVIF_INDENT = &H10
      LVIF_GROUPID = &H100
      LVIF_COLUMNS = &H200
      LVIF_NORECOMPUTE = &H800
      LVIF_DI_SETITEM = &H1000   ' NMLVDISPINFO notification
      '6.0
      LVIF_COLFMT = &H10000
    End Enum

    ' LVITEM state, stateMask, LVM_SETCALLBACKMASK wParam
    Private Enum LVITEM_state
      LVIS_FOCUSED = &H1
      LVIS_SELECTED = &H2
      LVIS_CUT = &H4
      LVIS_DROPHILITED = &H8
      LVIS_GLOW = &H10
      LVIS_ACTIVATING = &H20
      LVIS_LINK = &H40 'UNDOCUMENTED
      LVIS_OVERLAYMASK = &HF00
      LVIS_STATEIMAGEMASK = &HF000&
    End Enum
    ' LVM_GET/SETITEM lParam
    Private Type LVITEM 'LVITEMW
      Mask As LVITEM_mask
      iItem As Long
      iSubItem As Long
      State As LVITEM_state
      StateMask As LVITEM_state
      pszText As LongPtr
      cchTextMax As Long
      iImage As Long
      lParam As LongPtr
    '#If (WIN32_IE >= &H300) Then
      iIndent As Long
    '#End If
    '#If (WIN32_IE >= &H501) Then
      iGroupId As Long
      cColumns As Long
      puColumns As LongPtr
    '#End If
    '#If (WIN32_IE >= &H600) Then
      piColFmt As LongPtr 'array of certain LVCFMT_ for each subitem
      iGroup As Long 'for single item in multiple groups in virtual listview
    '#End If
    End Type
    Private Enum LVCOLUMN_mask
      LVCF_FMT = &H1
      LVCF_WIDTH = &H2
      LVCF_TEXT = &H4
      LVCF_SUBITEM = &H8
    '#If (WIN32_IE >= &H300) Then
      LVCF_IMAGE = &H10
      LVCF_ORDER = &H20
    '#End If
    '#If (WIN32_IE >= &H600) Then
      LVCF_MINWIDTH = &H40
      LVCF_DEFAULTWIDTH = &H80
      LVCF_IDEALWIDTH = &H100
    '#End If
    End Enum
    Private Enum LVCOLUMN_fmt
      LVCFMT_LEFT = &H0
      LVCFMT_RIGHT = &H1
      LVCFMT_CENTER = &H2
      LVCFMT_JUSTIFYMASK = &H3
    '#If (WIN32_IE >= &H300) Then
      LVCFMT_IMAGE = &H800
      LVCFMT_BITMAP_ON_RIGHT = &H1000
      LVCFMT_COL_HAS_IMAGES = &H8000&
    '#End If
    '#If (WIN32_IE >= &H600) Then
      LVCFMT_FIXED_WIDTH = &H100
      LVCFMT_NO_DPI_SCALE = &H40000
      LVCFMT_FIXED_RATIO = &H80000
      LVCFMT_LINE_BREAK = &H100000
      LVCFMT_FILL = &H200000
      LVCFMT_WRAP = &H400000
      LVCFMT_NO_TITLE = &H800000
      LVCFMT_TILE_PLACEMENTMASK = (LVCFMT_LINE_BREAK Or LVCFMT_FILL)
      LVCFMT_SPLITBUTTON = &H1000000
    '#End If
    End Enum
    Private Type LVCOLUMNW   ' was LV_COLUMN
      Mask As LVCOLUMN_mask
      fmt As LVCOLUMN_fmt
      CX As Long
      pszText As LongPtr  ' if String, must be pre-allocated
      cchTextMax As Long
      iSubItem As Long
    '#If (WIN32_IE >= &H300) Then
      iImage As Long
      iOrder As Long
    '#End If
    '#if (WIN32_IE >= &H600)
      cxMin As Long
      cxDefault As Long
      cxIdeal As Long
    '#End If
    End Type
    
    
    Private Const WC_COMBOBOXEXW = "ComboBoxEx32"
    Private Const WC_COMBOBOXEX = WC_COMBOBOXEXW
    
    Private Const H_MAX As Long = (&HFFFF + 1)

    Private Const CB_ADDSTRING = &H143
    Private Const CB_DELETESTRING = &H144
    Private Const CB_DIR = &H145
    Private Const CB_FINDSTRING = &H14C
    Private Const CB_FINDSTRINGEXACT = &H158
    Private Const CB_GETCOMBOBOXINFO = &H164
    Private Const CB_GETCOUNT = &H146
    Private Const CB_GETCURSEL = &H147
    Private Const CB_GETDROPPEDCONTROLRECT = &H152
    Private Const CB_GETDROPPEDSTATE = &H157
    Private Const CB_GETDROPPEDWIDTH = &H15F
    Private Const CB_GETEDITSEL = &H140
    Private Const CB_GETEXTENDEDUI = &H156
    Private Const CB_GETHORIZONTALEXTENT = &H15D
    Private Const CB_GETITEMDATA = &H150
    Private Const CB_GETITEMHEIGHT = &H154
    Private Const CB_GETLBTEXT = &H148
    Private Const CB_GETLBTEXTLEN = &H149
    Private Const CB_GETLOCALE = &H15A
    Private Const CB_GETTOPINDEX = &H15B
    Private Const CB_INITSTORAGE = &H161
    Private Const CB_INSERTSTRING = &H14A
    Private Const CB_LIMITTEXT = &H141
    Private Const CB_MSGMAX = &H15B
    Private Const CB_MULTIPLEADDSTRING = &H163
    Private Const CB_RESETCONTENT = &H14B
    Private Const CB_SELECTSTRING = &H14D
    Private Const CB_SETCURSEL = &H14E
    Private Const CB_SETDROPPEDWIDTH = &H160
    Private Const CB_SETEDITSEL = &H142
    Private Const CB_SETEXTENDEDUI = &H155
    Private Const CB_SETHORIZONTALEXTENT = &H15E
    Private Const CB_SETITEMDATA = &H151
    Private Const CB_SETITEMHEIGHT = &H153
    Private Const CB_SETLOCALE = &H159
    Private Const CB_SETTOPINDEX = &H15C
    Private Const CB_SHOWDROPDOWN = &H14F
    Private Const CBEC_SETCOMBOFOCUS = (&H165 + 1)   ' ;internal_nt
    Private Const CBEC_KILLCOMBOFOCUS = (&H165 + 2) ';internal_nt
    Private Const CBM_FIRST As Long = &H1700&
    Private Const CB_SETMINVISIBLE = (CBM_FIRST + 1)
    Private Const CB_GETMINVISIBLE = (CBM_FIRST + 2)
    Private Const CB_SETCUEBANNER = (CBM_FIRST + 3)
    Private Const CB_GETCUEBANNER = (CBM_FIRST + 4)
    Private Const CBEM_INSERTITEMA = (WM_USER + 1)
    Private Const CBEM_SETIMAGELIST = (WM_USER + 2)
    Private Const CBEM_GETIMAGELIST = (WM_USER + 3)
    Private Const CBEM_GETITEMA = (WM_USER + 4)
    Private Const CBEM_SETITEMA = (WM_USER + 5)
    Private Const CBEM_DELETEITEM = CB_DELETESTRING
    Private Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
    Private Const CBEM_GETEDITCONTROL = (WM_USER + 7)
    Private Const CBEM_SETEXTENDEDSTYLE = (WM_USER + 8)
    Private Const CBEM_GETEXTENDEDSTYLE = (WM_USER + 9)
    Private Const CBEM_HASEDITCHANGED = (WM_USER + 10)
    Private Const CBEM_INSERTITEMW = (WM_USER + 11)
    Private Const CBEM_SETITEMW = (WM_USER + 12)
    Private Const CBEM_GETITEMW = (WM_USER + 13)
    Private Const CBEM_INSERTITEM = CBEM_INSERTITEMW
    Private Const CBEM_SETITEM = CBEM_SETITEMW
    Private Const CBEM_GETITEM = CBEM_GETITEMW
 
    Private Enum ComboBox_Styles
        CBS_SIMPLE = &H1&
        CBS_DROPDOWN = &H2&
        CBS_DROPDOWNLIST = &H3&
        CBS_OWNERDRAWFIXED = &H10&
        CBS_OWNERDRAWVARIABLE = &H20&
        CBS_AUTOHSCROLL = &H40
        CBS_OEMCONVERT = &H80
        CBS_SORT = &H100&
        CBS_HASSTRINGS = &H200&
        CBS_NOINTEGRALHEIGHT = &H400&
        CBS_DISABLENOSCROLL = &H800&
        CBS_UPPERCASE = &H2000
        CBS_LOWERCASE = &H4000
    End Enum

    '// Notification messages
    Private Const CBN_ERRSPACE = (-1)
    Private Const CBN_SELCHANGE = 1
    Private Const CBN_DBLCLK = 2
    Private Const CBN_SETFOCUS = 3
    Private Const CBN_KILLFOCUS = 4
    Private Const CBN_EDITCHANGE = 5
    Private Const CBN_EDITUPDATE = 6
    Private Const CBN_DROPDOWN = 7
    Private Const CBN_CLOSEUP = 8
    Private Const CBN_SELENDOK = 9
    Private Const CBN_SELENDCANCEL = 10
    Private Const CBEN_FIRST = (H_MAX - 800&)
    Private Const CBEN_LAST = (H_MAX - 830&)
    Private Const CBEN_GETDISPINFOA = (CBEN_FIRST - 0)
    Private Const CBEN_GETDISPINFOW = (CBEN_FIRST - 7)
    Private Const CBEN_GETDISPINFO = CBEN_GETDISPINFOW
    Private Const CBEN_INSERTITEM = (CBEN_FIRST - 1)
    Private Const CBEN_DELETEITEM = (CBEN_FIRST - 2)
    Private Const CBEN_BEGINEDIT = (CBEN_FIRST - 4)
    Private Const CBEN_ENDEDITA = (CBEN_FIRST - 5)
    Private Const CBEN_ENDEDITW = (CBEN_FIRST - 6)
    Private Const CBEN_ENDEDIT = CBEN_ENDEDITW
    Private Const CBEN_DRAGBEGINA = (CBEN_FIRST - 8)
    Private Const CBEN_DRAGBEGINW = (CBEN_FIRST - 9)
    Private Const CBEN_DRAGBEGIN = CBEN_DRAGBEGINW
    '// lParam specifies why the endedit is happening
    Private Const CBENF_KILLFOCUS = 1
    Private Const CBENF_RETURN = 2
    Private Const CBENF_ESCAPE = 3
    Private Const CBENF_DROPDOWN = 4

    Private Enum CBEX_ExStyles
        CBES_EX_NOEDITIMAGE = &H1
        CBES_EX_NOEDITIMAGEINDENT = &H2
        CBES_EX_PATHWORDBREAKPROC = &H4
        CBES_EX_NOSIZELIMIT = &H8
        CBES_EX_CASESENSITIVE = &H10
        '6.0
        CBES_EX_TEXTENDELLIPSIS = &H20
    End Enum
    Private Type COMBOBOXEXITEM
        Mask As COMBOBOXEXITEM_Mask
        iItem As LongPtr
        pszText As String
        cchTextMax As Long
        iImage As Long
        iSelectedImage As Long
        iOverlay As Long
        iIndent As Long
        lParam As LongPtr
    End Type
    Private Type COMBOBOXEXITEMW
        Mask As COMBOBOXEXITEM_Mask
        iItem As LongPtr
        pszText As LongPtr      '// LPCSTR
        cchTextMax As Long
        iImage As Long
        iSelectedImage As Long
        iOverlay As Long
        iIndent As Long
        lParam As LongPtr
    End Type
    Private Enum COMBOBOXEXITEM_Mask
        CBEIF_TEXT = &H1
        CBEIF_IMAGE = &H2
        CBEIF_SELECTEDIMAGE = &H4
        CBEIF_OVERLAY = &H8
        CBEIF_INDENT = &H10
        CBEIF_LPARAM = &H20
        CBEIF_DI_SETITEM = &H10000000
    End Enum
    
    
    Private Function HIWORD(ByVal value As Long) As Integer
    HIWORD = (value And &HFFFF0000) \ &H10000
    End Function
    Private Function SUCCEEDED(hr As Long) As Boolean
        SUCCEEDED = (hr >= 0)
    End Function
    Private Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String
    SysReAllocStringW VarPtr(LPWSTRtoStr), lPtr
    If fFree Then
        Call CoTaskMemFree(lPtr)
    End If
    End Function
    Private Function ImageList_AddIcon(himl As LongPtr, hIcon As LongPtr) As Long
      ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
    End Function
    Private Function PKEY_PropList_InfoTip() As PROPERTYKEY
    Static pkk As PROPERTYKEY
     If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 4)
    PKEY_PropList_InfoTip = pkk
    End Function
    Private Function PKEY_Null() As PROPERTYKEY
    Static pkk As PROPERTYKEY
     If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, 0)
    PKEY_Null = pkk
    End Function
    Private Function ListView_SetSelectedItem(hwndLV As LongPtr, i As Long) As Boolean
      ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                                                                         LVIS_FOCUSED Or LVIS_SELECTED)
    End Function
    Private Function ListView_SetItemState(hwndLV As LongPtr, i As Long, State As LVITEM_state, Mask As LVITEM_state) As Boolean
      Dim lvi As LVITEM
      lvi.State = State
      lvi.StateMask = Mask
      ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
    End Function
    Private Function ListView_EnsureVisible(hwndLV As LongPtr, i As Long, fPartialOK As Long) As Boolean
      ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal fPartialOK)   ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
    End Function
    Private Function ListView_GetItemCount(hWnd As LongPtr) As LongPtr
      ListView_GetItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0)
    End Function
    Private Function ListView_GetItemRect(hWnd As LongPtr, i As Long, prc As RECT, code As LVIR_Flags) As Boolean
      prc.Left = code
      ListView_GetItemRect = SendMessage(hWnd, LVM_GETITEMRECT, ByVal i, prc)
    End Function
    Private Function ListView_GetSelectedItem(hwndLV As LongPtr) As LongPtr
      ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
    End Function
    Private Function ListView_GetNextItem(hWnd As LongPtr, i As Long, Flags As LVNI_Flags) As LongPtr
      ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal Flags)    ' ByVal MAKELPARAM(flags, 0))
    End Function
    
    
#End If



Private Sub DebugAppend(ByVal sMsg As String, Optional ilvl As Long = 0)
If ilvl > dbg_VerbosityLevel Then Exit Sub
Dim sOut As String
If dbg_IncludeDate Then sOut = "[" & Format$(Now, dbg_dtFormat) & "] "
If dbg_IncludeName Then sOut = sOut & Ambient.DisplayName & ": "
sOut = sOut & sMsg
If dbg_PrintToImmediate Then Debug.Print sOut
' If dbg_RaiseEvent = True Then RaiseEvent DebugMessage(sOut, CInt(ilvl))
' If dbg_PrintToFile Then
'     If log_hFile Then
'         WriteLog sOut
'     End If
' End If
End Sub

Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
    Dim hDC As LongPtr
    hDC = GetDC(0)
    mDPI = GetDeviceCaps(hDC, LOGPIXELSX) / USER_DEFAULT_SCREEN_DPI
    'Get the actual scale factor even if virtualized.
    
    Dim tDC As LongPtr, lRez As Long, lDPI As Long
    tDC = GetDC(0)
    lRez = GetDeviceCaps(tDC, DESKTOPHORZRES)
    lDPI = 96! * lRez / (Screen.Width / Screen.TwipsPerPixelX) * 15 / (1440 / GetDeviceCaps(tDC, LOGPIXELSX))
    ReleaseDC 0, tDC
    mActualZoom = CSng(lDPI) / 96!
    
    
    ' Dim hMonitor As LongPtr
    ' hMonitor = MonitorFromWindow(UserControl.hWnd, MONITOR_DEFAULTTONEAREST)

    ' ' Get the logical width And height of the monitor.
    ' Dim miex As MONITORINFOEXW
    ' miex.info.cbSize = LenB(miex)
    ' GetMonitorInfoW hMonitor, miex
    ' Dim cxLogical As Long
    ' cxLogical = (miex.info.rcMonitor.Right - miex.info.rcMonitor.Left)
    ' Dim cyLogical As Long
    ' cyLogical = (miex.info.rcMonitor.Bottom - miex.info.rcMonitor.Top)

    ' 'Get the physical width And height of the monitor.
    ' Dim dm As DEVMODEW
    ' dm.dmSize = LenB(dm)
    ' dm.dmDriverExtra = 0
    ' If EnumDisplaySettingsW(VarPtr(miex.szDevice(0)), ENUM_CURRENT_SETTINGS, dm) Then
    '     Debug.Print "EnumDisplaySettingsW success"
    ' Else
    '     Debug.Print "EnumDisplaySettingsW error: " & Err.LastDllError & " " & GetSystemErrorString(Err.LastDllError)
    '     Debug.Print GetLastError()
    ' End If
    ' Dim cxPhysical As Long
    ' cxPhysical = dm.dmPelsWidth
    ' Dim cyPhysical As Long
    ' cyPhysical = dm.dmPelsHeight

    ' ' Calculate the scaling factor.
    ' mActualZoom = CSng(cxPhysical) / CSng(cxLogical)
    Debug.Print "mActualZoom=" & mActualZoom & ", mDPI=" & mDPI
    ReDim SysImgCache(0)
    smCXEdge = GetSystemMetrics(SM_CXFIXEDFRAME)
    smCYEdge = GetSystemMetrics(SM_CXFIXEDFRAME)
    If smCXEdge = 0 Then smCXEdge = 1
    If smCYEdge = 0 Then smCYEdge = 1
    mIdxSel = -1
    IsComCtl6 = (ComCtlVersion >= 6)
    
End Sub
Private Sub InitImageLists()
    'himlTV = ImageList_Create(mIconSize, mIconSize, ILC_COLOR32 Or ILC_MASK Or ILC_HIGHQUALITYSCALE, 1, 1)
    If IsComCtl6 = False Then
        himlMain = ImageList_Create(cxyIcon * mDPI, cxyIcon * mDPI, ILC_COLOR32 Or ILC_MASK, 1, 1)
        Dim clbk As Long
        OleTranslateColor mBk, 0&, clbk
        ImageList_SetBkColor himlMain, clbk
    Else
        himlMain = ImageList_Create(cxyIcon * mDPI, cxyIcon * mDPI, ILC_COLOR32 Or ILC_MASK Or ILC_HIGHQUALITYSCALE, 1, 1)
    End If
    DebugAppend "InitImageLists->IsComCtl=" & IsComCtl6 & ",himlMain=" & himlMain
    Call SHGetImageList(SHIL_JUMBO, IID_IImageList, pImlLarge)
    Call SHGetImageList(SHIL_SMALL, IID_IImageList, pImlSmall)
    himlLarge = ObjPtr(pImlLarge)
    himlSmall = ObjPtr(pImlSmall)
End Sub
Private Function TranslateIcon(nIcon As Long, si As IShellItem, dwAttr As Long, pidlFQCur As LongPtr, pidlFQ As LongPtr, CX As Long, cy As Long, Optional pidlRel As LongPtr = 0, Optional bFlag1 As Boolean = False) As Long
    DebugAppend "TranslateIcon::Entry", 2
'Takes a system image list index and returns the local TreeView index.
'If not added already, adds it.
Dim lIdx As Long
lIdx = SysImlCacheLookup(nIcon)
DebugAppend "TranslateIcon::CacheLookup=" & lIdx

If lIdx > -1 Then
    DebugAppend "TranslateIcon " & nIcon & "|" & lIdx & " (Cached)", 2
    TranslateIcon = lIdx
    Exit Function
End If
DebugAppend "TranslateIcon::PreLoadUncached", 2
lIdx = AddToHIMLNoDLL(himlMain, si, dwAttr, pidlFQCur, pidlFQ, CX, cy, pidlRel)
DebugAppend "TranslateIcon::PostLoadUncache, lIdx=" & lIdx
ReDim Preserve SysImgCache(nSysImgCache)
SysImgCache(nSysImgCache).sysimlid = nIcon
SysImgCache(nSysImgCache).limlidx = lIdx
nSysImgCache = nSysImgCache + 1
DebugAppend "TranslateIcon " & nIcon & "|" & lIdx & " (added)", 2
TranslateIcon = lIdx
End Function
Private Function EnsureOverlay(nIdx As Long) As Long
    If nIdx = -1 Then Exit Function
    Debug.Print "added = " & bOvrAdded(nIdx)
    
    If bOvrAdded(nIdx) Then
        EnsureOverlay = 1
        Exit Function
    End If
    
    Dim nOvr As Long
    Dim hIcon As LongPtr
    Dim nPos As Long
    pImlLarge.GetOverlayImage nIdx, nOvr
    If nOvr >= 0 Then
        pImlLarge.GetIcon nOvr, ILD_TRANSPARENT, hIcon
        nPos = ImageList_AddIcon(himlMain, hIcon)
        Call DestroyIcon(hIcon)
        ImageList_SetOverlayImage himlMain, nPos, nIdx
        bOvrAdded(nIdx) = True
    End If
End Function
Private Function SysImlCacheLookup(nFI As Long) As Long
Dim i As Long
For i = 0 To UBound(SysImgCache)
    If SysImgCache(i).sysimlid = nFI Then
        SysImlCacheLookup = SysImgCache(i).limlidx
        Exit Function
    End If
Next i
SysImlCacheLookup = -1
End Function
Private Function AddToHIMLNoDLL(himl As LongPtr, si As IShellItem, dwAttr As Long, pidlFQCur As LongPtr, pidlFQ As LongPtr, CX As Long, cy As Long, Optional pidlRel As LongPtr = 0) As Long
Dim isiif As IShellItemImageFactory
Dim hr As Long
Dim pidlcr As LongPtr
Dim hBmp As LongPtr
If (si Is Nothing) Then
    If (pidlFQ = 0&) And (pidlRel <> 0&) Then
        'Virtual object; try to recreate pidl
        pidlcr = ILCombine(pidlFQCur, pidlRel)
    Else
        pidlcr = pidlFQ
    End If
    hr = SHCreateItemFromIDList(pidlcr, IID_IShellItemImageFactory, isiif)
Else
    Set isiif = si
End If
If isiif Is Nothing Then
    DebugAppend "AddToHIMLNoDLL->Couldn't get image factory."
    AddToHIMLNoDLL = -1
    Exit Function
End If
#If TWINBASIC Then
Dim tsz As Size
#Else
Dim tsz As oleexp.Size
#End If
'BUGFIX: Some Windows versions, for entirely unknown reasons, for the standard
'        printer icon, load the 32x32 version if you ask for 48x48; request
'        49x49 to actually get 48x48.
If CX = 48 Then
    tsz.CX = CX + 1: tsz.cy = cy + 1
Else
    tsz.CX = CX: tsz.cy = cy
End If
Dim lFlags As SIIGBF
lFlags = SIIGBF_BIGGERSIZEOK
#If TWINBASIC Then
Dim ull As LongLong
CopyMemory ull, tsz, 8
hr = isiif.GetImage(ull, lFlags, hBmp)
#Else
hr = isiif.GetImage(tsz.CX, tsz.cy, lFlags, hBmp)
#End If
If hr = S_OK Then
'    If ThumbShouldFrame(hBmp) Then
'        hr = E_FAIL 'This manual checking should only be needed for IL_AddMasked
                    'But it can't hurt to verify anyway; when a fail is returned
                    'from this function it goes to the GDIP scaler/framer.
'    End If
Else
    lFlags = SIIGBF_ICONONLY
    #If TWINBASIC Then
    hr = isiif.GetImage(ull, lFlags, hBmp)
    #Else
    hr = isiif.GetImage(tsz.CX, tsz.cy, lFlags, hBmp)
    #End If
End If
        
If hr = S_OK Then
    Dim clrMsk As Long
    If IsComCtl6 = False Then
        OleTranslateColor UserControl.ForeColor, 0&, clrMsk
        AddToHIMLNoDLL = ImageList_AddMasked(himl, hBmp, clrMsk)
    Else
        AddToHIMLNoDLL = ImageList_Add(himl, hBmp, 0&)
        ' If (AddToHIMLNoDLL = -1) And (hBmp <> 0) And (hBmp <> -1) Then
        '     AddToHIMLNoDLL = AddToImageListEx(himl, hBmp, cx, cy)
        ' End If
                
    End If
    DeleteObject hBmp
Else
    AddToHIMLNoDLL = -1
End If
    
Set isiif = Nothing
    Debug.Print "AddToHIMLNoDLL return=" & AddToHIMLNoDLL
End Function
Private Function ComCtlVersion() As Long
Dim tVI As DLLVERSIONINFO
On Error Resume Next
tVI.cbSize = LenB(tVI)
If DllGetVersion(tVI) = S_OK Then ComCtlVersion = tVI.dwMajorVersion
End Function

Private Sub UserControl_Resize() 'Handles UserControl.Resize
    If hCombo Then
        Dim rc As RECT
        Dim rcWnd As RECT
        GetClientRect UserControl.hWnd, rc
        SetWindowPos hCombo, 0, 0, 0, rc.Right, cyList * mDPI, SWP_NOMOVE Or SWP_NOZORDER
        With UserControl
        MoveWindow hCombo, 0, 0, .ScaleWidth, .ScaleHeight, 1
        GetWindowRect hCombo, rcWnd
        If (rcWnd.Bottom - rcWnd.Top) <> .ScaleHeight Or (rcWnd.Right - rcWnd.Left) <> .ScaleWidth Then
            .Extender.Height = .ScaleY((rcWnd.Bottom - rcWnd.Top), vbPixels, vbContainerSize)
        End If
        End With
    End If
End Sub

Private Sub UserControl_GotFocus() 'Handles UserControl.GotFocus
    DebugAppend "UserControl_GotFocus"
End Sub

Private Sub UserControl_ExitFocus() 'Handles UserControl.ExitFocus
    DebugAppend "UserControl_ExitFocus"
End Sub

Private Sub UserControl_EnterFocus() 'Handles UserControl.EnterFocus
    DebugAppend "UserControl_EnterFocus"
End Sub

Private Sub UserControl_Show() 'Handles UserControl.Show
    DebugAppend "UserControl_Show"
    If mInit = False Then
        mInit = True
        InitControl
    End If
End Sub

Private Sub UserControl_Terminate() 'Handles UserControl.Terminate
    If hLVW Then
        SendMessage hLVW, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal 0
        SendMessage hLVW, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal 0
        DestroyWindow hLVW
    End If
    ImageList_Destroy himlMain
    DestroyWindow hCombo
    DeleteObject hFont
    DeleteObject hFontBold
    hFont = 0
    hFontBold = 0
    If hTheme Then CloseThemeData hTheme
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) 'Handles UserControl.ReadProperties
    DebugAppend "UserControl_ReadProperties"
    cxList = PropBag.ReadProperty("DropdownWidth", mDefCX)
    cyList = PropBag.ReadProperty("DropdownHeight", mDefCY)
    mNotify = PropBag.ReadProperty("MonitorChanges", mDefNotify)
    mBk = PropBag.ReadProperty("BackColor", mDefBk)
    mEnabled = PropBag.ReadProperty("Enabled", mDefEnabled)
    mStyle = PropBag.ReadProperty("ComboStyle", mDefStyle)
    mListView = PropBag.ReadProperty("UseListView", mDefListView)
    mTrack = PropBag.ReadProperty("ListViewHotTrack", mDefTrack)
    mRaiseOnLoad = PropBag.ReadProperty("RaiseChangeOnLoad", mDefRaiseOnLoad)
    mLimitCX = PropBag.ReadProperty("NoExtendWidth", mDefLimitCX)
    cxyIcon = PropBag.ReadProperty("IconSize", mDefIcon)
    mNoRf = PropBag.ReadProperty("NoRefreshTipOnDrop", mDefNoRf)
    Set PropFont = PropBag.ReadProperty("Font", Nothing)
    mInit = True
    InitControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag) 'Handles UserControl.WriteProperties
    PropBag.WriteProperty "DropdownWidth", cxList, mDefCX
    PropBag.WriteProperty "DropdownHeight", cyList, mDefCY
    PropBag.WriteProperty "BackColor", mBk, mDefBk
    PropBag.WriteProperty "MonitorChanges", mNotify, mDefNotify
    PropBag.WriteProperty "Enabled", mEnabled, mDefEnabled
    PropBag.WriteProperty "ComboStyle", mStyle, mDefStyle
    PropBag.WriteProperty "UseListView", mListView, mDefListView
    PropBag.WriteProperty "ListViewHotTrack", mTrack, mDefTrack
    PropBag.WriteProperty "RaiseChangeOnLoad", mRaiseOnLoad, mDefRaiseOnLoad
    PropBag.WriteProperty "NoExtendWidth", mLimitCX, mDefLimitCX
    PropBag.WriteProperty "IconSize", cxyIcon, mDefIcon
    PropBag.WriteProperty "NoRefreshTipOnDrop", mNoRf, mDefNoRf
    PropBag.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
End Sub

Private Sub UserControl_InitProperties() 'Handles UserControl.InitProperties
    cxList = mDefCX
    cyList = mDefCY
    mBk = mDefBk
    mNotify = mDefNotify
    mEnabled = mDefEnabled
    mStyle = mDefStyle
    mTrack = mDefTrack
    mRaiseOnLoad = mDefRaiseOnLoad
    mLimitCX = mDefLimitCX
    mListView = mDefListView
    cxyIcon = mDefIcon
    mNoRf = mDefNoRf
    Set PropFont = Ambient.Font
    Debug.Print "InitProps->Font=" & Ambient.Font.Name
End Sub
Public Property Get NoRefreshTipOnDrop() As Boolean: NoRefreshTipOnDrop = mNoRf: End Property
Attribute NoRefreshTipOnDrop.VB_Description = "Don't reload the status information before showing the dropdown ListView."
Public Property Let NoRefreshTipOnDrop(ByVal value As Boolean): mNoRf = value: End Property
Public Property Get IconSize() As Long: IconSize = cxyIcon: End Property
Attribute IconSize.VB_Description = "Size of the icons in ListView mode. Must be set at design time."
Public Property Let IconSize(ByVal cxy As Long): cxyIcon = cxy: End Property
Public Property Get NoExtendWidth() As Boolean: NoExtendWidth = mLimitCX: End Property
Attribute NoExtendWidth.VB_Description = "Never extend the dropdown width beyond the width of the control."
Public Property Let NoExtendWidth(ByVal bValue As Boolean): mLimitCX = bValue: End Property
Public Property Get RaiseChangeOnLoad() As Boolean: RaiseChangeOnLoad = mRaiseOnLoad: End Property
Attribute RaiseChangeOnLoad.VB_Description = "Raise a PrinterChange even when the default printer is automatically selected on load. It will also be raised if the printer list is refreshed."
Public Property Let RaiseChangeOnLoad(ByVal bValue As Boolean): mRaiseOnLoad = bValue: End Property
Public Property Get UseListView() As Boolean: UseListView = mListView: End Property
Attribute UseListView.VB_Description = "Use a ListView with large icon and printer status instead of the normal dropdown listbox."
Public Property Let UseListView(ByVal bValue As Boolean): mListView = bValue: End Property
Public Property Get ListViewHotTrack() As Boolean: ListViewHotTrack = mTrack: End Property
Attribute ListViewHotTrack.VB_Description = "Have the selection follow the cursor when using the ListView dropdown."
Public Property Let ListViewHotTrack(ByVal bValue As Boolean): mTrack = bValue: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = mBk: End Property
Public Property Let BackColor(ByVal cr As OLE_COLOR)
    mBk = cr
    UserControl.BackColor = cr
End Property
Public Property Get ComboStyle() As UCPCType: ComboStyle = mStyle: End Property
Attribute ComboStyle.VB_Description = "Sets the type of combobox used. Cannnot be changed during runtime."
Public Property Let ComboStyle(ByVal value As UCPCType): mStyle = value: End Property
Public Property Get Enabled() As Boolean: Enabled = mEnabled: End Property
Attribute Enabled.VB_Description = "Sets whether the control is enabled."
Public Property Let Enabled(ByVal fEnable As Boolean)
    If fEnable <> mEnabled Then
        mEnabled = fEnable
        If hCombo Then
            If mEnabled Then
                EnableWindow hCombo, CTRUE
            Else
                EnableWindow hCombo, CFALSE
            End If
        End If
    End If
End Property
Public Property Get DropdownWidth() As Long: DropdownWidth = cxList: End Property
Attribute DropdownWidth.VB_Description = "The maximum width, when wider than the control but less than the untruncated item width."
Public Property Let DropdownWidth(ByVal value As Long)
    If value <> cxList Then
        cxList = value
        If Ambient.UserMode Then
            If Not mLimitCX Then
                SendMessage hCombo, CB_SETDROPPEDWIDTH, cxList, ByVal 0
            End If
        End If
    End If
End Property

Public Property Get DropdownHeight() As Long: DropdownHeight = cyList: End Property
Attribute DropdownHeight.VB_Description = "Sets the maximum height of the dropdown, when not limited by total item height or available screen space."
Public Property Let DropdownHeight(ByVal value As Long)
    If value <> cyList Then
        cyList = value
    End If
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Sets the font for the text. "
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
'DebugAppend "FontSet"
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As LongPtr
Set PropFont = NewFont
OldFontHandle = hFont
Set mIFMain = PropFont
Dim lftmp As LOGFONT
GetObjectW mIFMain.hFont, LenB(lftmp), lftmp
hFont = CreateFontIndirect(lftmp)
If hFontBold Then
    DeleteObject hFontBold
End If
lftmp.LFWeight = FW_BOLD
hFontBold = CreateFontIndirect(lftmp)
If hCombo <> 0 Then SendMessageW hCombo, WM_SETFONT, hFont, ByVal 1&
If hLVW <> 0 Then SendMessageW hLVW, WM_SETFONT, hFont, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property
Private Sub PropFont_FontChanged(ByVal PropertyName As String) 'Handles PropFont.FontChanged
    DebugAppend "FontChanged"
    Dim OldFontHandle As LongPtr
    OldFontHandle = hFont
    Set mIFMain = PropFont
    Dim lftmp As LOGFONT
    GetObjectW mIFMain.hFont, LenB(lftmp), lftmp
    hFont = CreateFontIndirect(lftmp)
    If hFontBold Then
        DeleteObject hFontBold
    End If
    lftmp.LFWeight = FW_BOLD
    hFontBold = CreateFontIndirect(lftmp)
    If hCombo <> 0 Then SendMessageW hCombo, WM_SETFONT, hFont, ByVal 1&
    If hLVW <> 0 Then SendMessageW hLVW, WM_SETFONT, hFont, ByVal 1&
    If OldFontHandle <> 0 Then DeleteObject OldFontHandle
    UserControl.PropertyChanged "Font"
End Sub

Public Property Get SelectedPrinter() As String
    If Ambient.UserMode Then
        If mIdxSel >= 0 Then
            SelectedPrinter = mPrinters(mIdxSel).sName
            Exit Property
        Else
            If mIdxDef >= 0 Then
                SelectedPrinter = mPrinters(mIdxDef).sName
                Exit Property
            End If
        End If
        Dim lRet As Long, cchDef As Long
        Dim sDef As String
        lRet = GetDefaultPrinterW(0, cchDef)
        If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
            sDef = String$(cchDef - 1, 0)
            lRet = GetDefaultPrinterW(StrPtr(sDef), cchDef)
        End If
        SelectedPrinter = sDef
    End If
End Property
Public Property Let SelectedPrinter(ByVal sName As String)
    If Ambient.UserMode Then
        If nPr Then
            Dim i As Long
            For i = 0 To UBound(mPrinters)
                If LCase$(sName) = LCase$(mPrinters(i).sName) Then
                    If i <> mIdxSel Then
                        mIdxSelPrev = mIdxSel
                        mIdxSel = i
                        If mIdxSelPrev <> mIdxSel Then
                            SendMessage hCombo, CB_SETCURSEL, mPrinters(i).cbi, ByVal 0
                            RaiseEvent PrinterChanged(mPrinters(i).sName, mPrinters(i).sParsingPath, mPrinters(i).sModel, mPrinters(i).sLocation, mPrinters(i).sLastStatus, mPrinters(i).bDefault)
                        End If
                    End If
                End If
            Next
        End If
    End If
End Property
Public Property Get SelectedIndex() As Long
Attribute SelectedIndex.VB_Description = "The index of the selected printer, suitable for GetPrinterInfo()."
    SelectedIndex = mIdxSel
End Property

Public Property Get PrinterCount() As Long: PrinterCount = nPr: End Property
Attribute PrinterCount.VB_Description = "Retrieves the name of the specified printer in the list."
Public Function Printers(ByVal index As Long) As String
Attribute Printers.VB_Description = "Retrieves the name of the specified printer in the list."
    If index <= UBound(mPrinters) Then
        Printers = mPrinters(index).sName
    End If
End Function
Public Sub GetPrinterInfo(ByVal index As Long, lpShellParsingPath As String, lpName As String, lpModel As String, lpLocation As String, lpLastStatusMessage As String, pbIsDefault As Boolean)
Attribute GetPrinterInfo.VB_Description = "Retrieves extended information about the specified printer."
    If index <= UBound(mPrinters) Then
        lpName = mPrinters(index).sName
        lpShellParsingPath = mPrinters(index).sParsingPath
        lpModel = mPrinters(index).sModel
        lpLocation = mPrinters(index).sLocation
        lpLastStatusMessage = mPrinters(index).sLastStatus
        pbIsDefault = mPrinters(index).bDefault
    End If
End Sub

Private Sub InitControl()
    DebugAppend "InitControl " & Ambient.UserMode
    Set Me.Font = PropFont
    InitImageLists
    pvCreateCombo
    If Ambient.UserMode Then
        SHLoadNonloadedIconOverlayIdentifiers
        RefreshPrinters
        pvCreateListView
        Subclass2 UserControl.hWnd, AddressOf ucPrinterUserControlWndProc, UserControl.hWnd, ObjPtr(Me)
    End If
End Sub
Private Sub pvCreateListView()
    Dim dwStyle As Long
    dwStyle = WS_CHILD Or WS_TABSTOP Or WS_BORDER Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_SINGLESEL
    hLVW = CreateWindowExW(0, StrPtr(WC_LISTVIEW), 0, dwStyle, 0, 0, 110, 110, UserControl.hWnd, 0, App.hInstance, ByVal 0)
    If hLVW Then
        Dim tCol As LVCOLUMNW
        tCol.Mask = LVCF_WIDTH Or LVCF_TEXT
        tCol.cchTextMax = Len(sCol0)
        tCol.pszText = StrPtr(sCol0)
        tCol.CX = UserControl.ScaleWidth
        SendMessage hLVW, LVM_INSERTCOLUMNW, 0, tCol
        
        tCol.Mask = LVCF_WIDTH Or LVCF_TEXT
        tCol.cchTextMax = Len(sCol1)
        tCol.pszText = StrPtr(sCol1)
        tCol.CX = UserControl.ScaleWidth
        SendMessage hLVW, LVM_INSERTCOLUMNW, 0, tCol
        
        SendMessage hLVW, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal himlMain
        SendMessage hLVW, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal himlSmall
        
        SetWindowTheme hLVW, StrPtr("Combobox"), 0
        'The ListView would have inherited the desktop font rather then
        'the one set by the user for this control:
        ' If hFont Then
        '     DeleteObject hFont
        ' End If
        ' Dim lftmp As LOGFONT
        ' GetObjectW mIFMain.hFont, LenB(lftmp), lftmp
        ' hFont = CreateFontIndirect(lftmp)
        ' If hFontBold Then
        '     DeleteObject hFontBold
        ' End If
        ' lftmp.LFWeight = FW_BOLD
        ' hFontBold = CreateFontIndirect(lftmp)
 
        ' SendMessage hLVW, WM_SETFONT, hFont, ByVal 0
        If Ambient.UserMode Then
 
           SetParent hLVW, 0
           Subclass2 hLVW, AddressOf ucPrinterLVWndProc, hLVW, ObjPtr(Me)
        End If
    End If
End Sub



Private Sub pvCreateCombo()
    Dim dwStyle As ComboBox_Styles
    dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL Or WS_TABSTOP
    If mStyle = UCPC_DropdownList Then
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
    Else
        dwStyle = dwStyle Or CBS_DROPDOWN
    End If
    Dim rc As RECT
    GetClientRect UserControl.hWnd, rc
    hCombo = CreateWindowExW(0, StrPtr(WC_COMBOBOXEX), 0, dwStyle, _
                            0, 0, rc.Right, cyList * mDPI, UserControl.hWnd, 0, App.hInstance, ByVal 0)

    hComboCB = SendMessage(hCombo, CBEM_GETCOMBOCONTROL, 0, ByVal 0&)
    hComboEd = SendMessage(hCombo, CBEM_GETEDITCONTROL, 0, ByVal 0&)

    If hComboEd Then SendMessage hComboEd, EM_SETREADONLY, 1&, ByVal 0&
    
    Call SendMessage(hCombo, CBEM_SETIMAGELIST, 0, ByVal himlSmall)
    DebugAppend "ImageListValid? " & himlSmall
    If Ambient.UserMode Then
        hTheme = OpenThemeData(hCombo, StrPtr("Combobox"))
        Subclass2 hCombo, AddressOf ucPrinterComboWndProc, hCombo, ObjPtr(Me)
        Subclass2 hComboCB, AddressOf ucPrinterComboWndProc, hComboCB, ObjPtr(Me)
        DebugAppend "hCombo=" & hCombo & ",hComboCB=" & hComboCB
        ' Dim tFilter As DEV_BROADCAST_DEVICEINTERFACE
        ' tFilter.dbcc_size = 32 'We can't use LenB because it uses the size above 28 to calculate
        '                         'the length of the string in the C-style variable array on the end.
        '                         'It's declared with a buffer since VB/tB don't support those, but if
        '                         'the buffer isn't in use, use what we'd get for sizeof() if it wasn't
        '                         'used in C++.
        ' tFilter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
        ' tFilter.dbcc_classguid = GUID_DEVINTERFACE_
        ' hNotify = RegisterDeviceNotification(hMain, tFilter, DEVICE_NOTIFY_WINDOW_HANDLE)
    Else
        Dim pidl As LongPtr
        Dim nIcon As Long
        SHGetFolderLocation 0, CSIDL_PRINTERS, 0, 0, pidl
        nIcon = GetIconIndexPidl(pidl, SHGFI_SMALLICON)
        CBX_InsertItem hCombo, Ambient.DisplayName, nIcon
        DebugAppend "Insert design mode icon " & nIcon
        SendMessage hCombo, CB_SETCURSEL, 0, ByVal 0
        CoTaskMemFree pidl
    End If

    If mEnabled = False Then
        EnableWindow hCombo, CFALSE
    End If
End Sub
Private Function CBX_InsertItem(ByVal hCBoxEx As LongPtr, sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1, Optional lParam As Long = 0, Optional iItem As Long = -1, Optional iIndent As Long = 0, Optional iImageSel As Long = -1) As Long

    Dim cbxi As COMBOBOXEXITEMW

    With cbxi
    .Mask = CBEIF_TEXT
    .cchTextMax = Len(sText)
    .pszText = StrPtr(sText)
    If iImage <> -1 Then
        .Mask = .Mask Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE
        .iImage = iImage
    End If
    If iOverlay <> -1 Then
        .iOverlay = iOverlay
    End If
    If lParam Then
        .Mask = .Mask Or CBEIF_LPARAM
        .lParam = lParam
    End If
    If iIndent Then
        .Mask = .Mask Or CBEIF_INDENT
        .iIndent = iIndent
    End If
    If iImageSel <> -1 Then
        .Mask = .Mask
        .iSelectedImage = iImageSel
    Else
        .iSelectedImage = iImage
    End If

    .iItem = iItem

    End With

    CBX_InsertItem = CLng(SendMessage(hCBoxEx, CBEM_INSERTITEMW, 0, cbxi))

End Function
Private Function GetCBXItemlParam(hWnd As LongPtr, i As Long) As LongPtr
    Dim cbxi As COMBOBOXEXITEMW
    With cbxi
    .Mask = CBEIF_LPARAM
    .iItem = i
    End With
    If SendMessage(hWnd, CBEM_GETITEMW, 0, cbxi) Then
    GetCBXItemlParam = cbxi.lParam
    Else
    GetCBXItemlParam = -1
    End If
End Function

 
Public Sub RefreshPrinters()
    mIdxSelPrev = mIdxSel
    If mIdxSelPrev >= 0 Then
        mLabelSelPrev = mPrinters(mIdxSel).sName
    End If
 
    Dim bFlag As Boolean
    If nPr Then
        mPrintersOld = mPrinters
        nPrOld = nPr
        bFlag = True
    End If
    
    DoPrinterEnum
    
    If mIdxSel = -1 Then mIdxSel = mIdxDef
    If mIdxSelPrev = -1 Then
        mIdxSelPrev = mIdxDef
        mLabelSelPrev = mPrinters(mIdxSel).sName
    End If
    
    RefreshPrintersCombo
    

End Sub

Private Function PrinterColChanged() As Boolean
    If nPr <> nPrOld Then
        PrinterColChanged = True
        Exit Function
    End If
    Dim i As Long
    For i = 0 To UBound(mPrinters)
        If mPrinters(i).sName <> mPrintersOld(i).sName Then
            PrinterColChanged = True
            Exit Function
        End If
    Next
End Function


Private Function GetPropertyKeyDisplayString(pps As IPropertyStore, pkProp As PROPERTYKEY, Optional bFixChars As Boolean = True) As String
'Gets the string value of the given canonical property; e.g. System.Company, System.Rating, etc
'This would be the value displayed in Explorer if you added the column in details view
'<EhHeader>
On Error GoTo e0
'</EhHeader>
Dim lpsz As LongPtr
Dim ppd As IPropertyDescription
If ((pps Is Nothing) = False) Then
    PSGetPropertyDescription pkProp, IID_IPropertyDescription, ppd
    If (ppd Is Nothing) Then
'        DebugAppend "GetPropertyKeyDisplayString->Could not obtain IPropertyDescription, will attempt alternative."
        Dim vrr As Variant, vbr As Variant
        pps.GetValue pkProp, vrr
        PropVariantToVariant vrr, vbr
        If (VarType(vbr) And vbArray) = vbArray Then
            Dim i As Long
            For i = LBound(vbr) To UBound(vbr)
                GetPropertyKeyDisplayString = GetPropertyKeyDisplayString & CStr(vbr(i)) & "; "
            Next i
            If Len(GetPropertyKeyDisplayString) > 2 Then
                GetPropertyKeyDisplayString = Left$(GetPropertyKeyDisplayString, Len(GetPropertyKeyDisplayString) - 2)
            End If
        Else
            GetPropertyKeyDisplayString = CStr(vbr)
        End If
    Else
        Dim hr As Long
        hr = PSFormatPropertyValue(ObjPtr(pps), ObjPtr(ppd), PDFF_DEFAULT, lpsz)
'        DebugAppend "prophr=0x" & Hex$(hr)
        SysReAllocStringW VarPtr(GetPropertyKeyDisplayString), lpsz
        CoTaskMemFree lpsz
    End If
    If bFixChars Then
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H200E), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H200F), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H202A), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H202C), "")
    End If
    Set ppd = Nothing
Else
    DebugAppend "GetPropertyKeyDisplayString.Error->PropertyStore is not set."
        
End If
'<EhFooter>
Exit Function
    
e0:
    DebugAppend "ucShellBrowse.GetPropertyKeyDisplayString->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Function PKEY_Printer_IsDefault() As PROPERTYKEY
    'System.Printer.Default
    '{FE9E4C12-AACB-4AA3-966D-91A29E6128B5, 3}
    Static pk As PROPERTYKEY
    If pk.fmtid.Data1 = 0 Then
        Call DEFINE_PROPERTYKEY(pk, &HFE9E4C12, &HAACB, &H4AA3, &H96, &H6D, &H91, &HA2, &H9E, &H61, &H28, &HB5, 3)
    End If
    PKEY_Printer_IsDefault = pk
End Function
Private Function PKEY_Printer_Location() As PROPERTYKEY
    'System.Printer.Default
    '{FE9E4C12-AACB-4AA3-966D-91A29E6128B5, 4}
    Static pk As PROPERTYKEY
    If pk.fmtid.Data1 = 0 Then
        Call DEFINE_PROPERTYKEY(pk, &HFE9E4C12, &HAACB, &H4AA3, &H96, &H6D, &H91, &HA2, &H9E, &H61, &H28, &HB5, 4)
    End If
    PKEY_Printer_Location = pk
End Function
Private Function PKEY_Printer_Model() As PROPERTYKEY
    'System.Printer.Default
    '{FE9E4C12-AACB-4AA3-966D-91A29E6128B5, 5}
    Static pk As PROPERTYKEY
    If pk.fmtid.Data1 = 0 Then
        Call DEFINE_PROPERTYKEY(pk, &HFE9E4C12, &HAACB, &H4AA3, &H96, &H6D, &H91, &HA2, &H9E, &H61, &H28, &HB5, 5)
    End If
    PKEY_Printer_Model = pk
End Function
Private Function PKEY_Printer_Status() As PROPERTYKEY
    'System.Printer.Default
    '{FE9E4C12-AACB-4AA3-966D-91A29E6128B5, 7}
    Static pk As PROPERTYKEY
    If pk.fmtid.Data1 = 0 Then
        Call DEFINE_PROPERTYKEY(pk, &HFE9E4C12, &HAACB, &H4AA3, &H96, &H6D, &H91, &HA2, &H9E, &H61, &H28, &HB5, 7)
    End If
    PKEY_Printer_Status = pk
End Function



Private Sub DoPrinterEnum()
    'On Error GoTo e0
    DebugAppend "DoPrinterEnum::Entry"
    ReDim mPrinters(0): nPr = 0
    Dim i As Long
    Dim pFolder As IShellItem
    Dim pEnum As IEnumShellItems
    Dim pPrinter As IShellItem, pPrinter2 As IShellItem2
    Dim pps As IPropertyStore
    Dim upi As IParentAndItem
    Dim psf As IShellFolder
    Dim pOverlay As IShellIconOverlay
    Dim pidlRel As LongPtr, pidlPar As LongPtr, pidlFQ As LongPtr
    Dim lpName As LongPtr, lpParse As LongPtr
    Dim lpTip As LongPtr
    Dim pkm As New KnownFolderManager
    Dim pkf As IKnownFolder
    Dim dwAtr As Long
    Dim pcl As Long
    Dim sDef As String
    Dim cchDef As Long
    Dim lRet As Long
    Dim nDefFB As Long 'Fallback method
    Dim nIcon As Long
    nDefFB = -1
    Dim bSetDef As Boolean
    lRet = GetDefaultPrinterW(0, cchDef)
    If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
        sDef = String$(cchDef - 1, 0)
        lRet = GetDefaultPrinterW(StrPtr(sDef), cchDef)
    End If
    DebugAppend "looking for default printer: " & sDef
    
    pkm.GetFolder FOLDERID_PrintersFolder, pkf
    If (pkf Is Nothing) = False Then
        pkf.GetShellItem KF_FLAG_DEFAULT, IID_IShellItem, pFolder
        If (pFolder Is Nothing) = False Then
            pFolder.BindToHandler 0, BHID_EnumItems, IID_IEnumShellItems, pEnum
            Do While pEnum.Next(1, pPrinter, pcl) = S_OK
                lpName = 0: lpParse = 0: pidlRel = 0: pidlPar = 0: pidlFQ = 0
                Set psf = Nothing: Set pOverlay = Nothing
                ReDim Preserve mPrinters(nPr)
                pPrinter.GetDisplayName SIGDN_NORMALDISPLAY, lpName
                mPrinters(nPr).sName = LPWSTRtoStr(lpName)
                DebugAppend "AddPrinter " & mPrinters(nPr).sName
                pPrinter.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpParse
                mPrinters(nPr).sParsingPath = LPWSTRtoStr(lpParse)
                Set upi = pPrinter
                upi.GetParentAndItem pidlPar, psf, pidlRel
                mPrinters(nPr).sInfoTip = GenerateInfoTip(pPrinter, psf, pidlRel)
                'DebugAppend "AddPrinterTip " & mPrinters(nPr).sInfoTip
                On Error Resume Next
                pPrinter.GetAttributes SFGAO_LINK Or SFGAO_SHARE, dwAtr
                pidlFQ = ILCombine(pidlPar, pidlRel)
                nIcon = GetIconIndexPidl(pidlFQ, SHGFI_LARGEICON)
                mPrinters(nPr).nIcon = nIcon
                mPrinters(nPr).nIconLV = TranslateIcon(nIcon, pPrinter, dwAtr, pidlPar, pidlFQ, cxyIcon * mDPI, cxyIcon * mDPI, pidlRel)
                CoTaskMemFree pidlFQ
                'DebugAppend "SetPrinterIcon " & mPrinters(nPr).nIcon
                mPrinters(nPr).nOvr = -1
                Set pOverlay = psf
                If (pOverlay Is Nothing) = False Then
                    pOverlay.GetOverlayIconIndex pidlRel, mPrinters(nPr).nOvr
                End If
                If (mPrinters(nPr).nOvr > 15) Or (mPrinters(nPr).nOvr < 0) Then
                    'Overlay icons are a mess. On Win7 there's a bunch in root that return 16, which is invalid
                    'and will cause a crash later one, and doesn't show anything. Shares never get shown so I'm
                    'going to manually set those
                    mPrinters(nPr).nOvr = -1
                    If (dwAtr And SFGAO_SHARE) = SFGAO_SHARE Then
                        mPrinters(nPr).nOvr = 1
                    End If
                    If (dwAtr And SFGAO_LINK) = SFGAO_LINK Then
                        mPrinters(nPr).nOvr = 2
                    End If
                End If
                EnsureOverlay mPrinters(nPr).nOvr
                If Len(sDef) Then
                    If mPrinters(nPr).sName = sDef Then
                        mPrinters(nPr).bDefault = True
                        mIdxDef = nPr
                        bSetDef = True
                    End If
                End If
                'Fallback method for determining default printer.
                'Likely to be removed in future versions after confirming
                'reliability of the primary method above.
                Set pPrinter2 = pPrinter
                Dim vrDef As Variant
                lRet = pPrinter2.GetProperty(PKEY_Printer_IsDefault, vrDef)
                If SUCCEEDED(lRet) Then
                    'This property is an LPWSTR; only the default will have it
                    'So no need to dereference-- it would be localized anyway.
                    If IsEmpty(vrDef) = False Then
                        nDefFB = nPr
                    End If
                End If
                
                pPrinter2.GetPropertyStore GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pps
                mPrinters(nPr).sModel = GetPropertyKeyDisplayString(pps, PKEY_Printer_Model)
                mPrinters(nPr).sLocation = GetPropertyKeyDisplayString(pps, PKEY_Printer_Location)
                mPrinters(nPr).sLastStatus = GetPropertyKeyDisplayString(pps, PKEY_Printer_Status)
                Set pps = Nothing
                
                On Error GoTo e0
                nPr = nPr + 1
            Loop
        End If
    End If
    If bSetDef = False Then
        If nDefFB >= 0 Then
            mPrinters(nDefFB).bDefault = True
            mIdxDef = nDefFB
        End If
    End If
    If mRaiseOnLoad Then
        RaiseEvent PrinterChanged(mPrinters(mIdxDef).sName, mPrinters(mIdxDef).sParsingPath, mPrinters(mIdxDef).sModel, mPrinters(mIdxDef).sLocation, mPrinters(mIdxDef).sLastStatus, mPrinters(mIdxDef).bDefault)
    End If
    DebugAppend "DoPrinterEnum::Done, count=" & nPr
    Exit Sub
e0:
    DebugAppend "Unexpected error in DoPrinterEnum->" & Err.Number & ": " & Err.Description
End Sub
 
Private Function GetIconIndexPidl(ByVal pidl As LongPtr, uType As Long) As Long
Dim sfi As SHFILEINFOW
If SHGetFileInfoW(ByVal pidl, 0, sfi, LenB(sfi), SHGFI_SYSICONINDEX Or SHGFI_PIDL Or uType) Then
    GetIconIndexPidl = sfi.iIcon
End If
End Function

Private Function GenerateInfoTip(si As IShellItem, psfCur As IShellFolder, pidlRel As LongPtr) As String
 
Dim sTip As String
On Error GoTo e0
 
If (si Is Nothing = False) Then
        
    Dim pqi As IQueryInfo
    si.BindToHandler 0&, BHID_SFUIObject, IID_IQueryInfo, pqi
    If (pqi Is Nothing) Then
'        DebugAppend "GenerateInfoTip::Try alternate..."
        If (psfCur Is Nothing) = False Then
            psfCur.GetUIObjectOf hLVW, 1&, pidlRel, IID_IQueryInfo, 0&, pqi
        End If
    End If
    If (pqi Is Nothing) = False Then
        Dim lpTip As LongPtr, sQITip As String
        Dim dwFlags As QITipFlags
        dwFlags = QITIPF_LINKUSETARGET Or QITIPF_USESLOWTIP Or QITIPF_SINGLELINE
        pqi.GetInfoTip dwFlags, lpTip
        sQITip = LPWSTRtoStr(lpTip)
'        DebugAppend "QITIPF_USESLOWTIPGenerateInfoTip::Exit->UseSlowTip=" & sQITip, 11
        GenerateInfoTip = Replace(Replace(Replace(Replace(Replace(sQITip, vbCrLf, ", "), vbLf, ", "), vbTab, vbNullString), "  ", " "), "  ", " ")
        Exit Function
    Else
        DebugAppend "Failed to get IQueryInfo"
    End If
        
    Dim lpp As Long
    Dim si2p As IShellItem2
    Dim pl As IPropertyDescriptionList
    Dim pd As IPropertyDescription
    Dim lpn As LongPtr, sPN As String
        
    Set si2p = si
    Dim pst As IPropertyStore
    si2p.GetPropertyDescriptionList PKEY_PropList_InfoTip, IID_IPropertyDescriptionList, pl
    If (pl Is Nothing) = False Then
        pl.GetCount lpp
'        DebugAppend "InfoTip Cnt=" & lpp
        If lpp Then
            Dim stt As String
            si2p.GetPropertyStore GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pst
            If (pst Is Nothing) = False Then
'                DebugAppend "PropsList=" & GetPropertyKeyDisplayString(pst, PKEY_PropList_InfoTip)
                'We could just parse that; but going through IPropertyDescriptionList automatically skips
                'fields where there's no data (an error is raised, hence the e1/resume next)GPS_DEFAULT Or
                On Error GoTo e1
                Dim i As Long
                For i = 0 To (lpp - 1)
                    pl.GetAt i, IID_IPropertyDescription, pd
                    If (pd Is Nothing) = False Then
                        stt = GetPropertyDisplayString(pst, pd, PKEY_Null)
                        If stt <> "" Then
                            pd.GetDisplayName lpn
                            sPN = LPWSTRtoStr(lpn)
                            stt = sPN & ": " & stt
                            If sTip = "" Then
                                sTip = stt
                            Else
                                sTip = sTip & vbCrLf & stt
                            End If
'                            DebugAppend "Prop=" & stt
                            stt = ""
                        Else
                            DebugAppend "Prop=(empty)"
                        End If
                        Set pd = Nothing
                    Else
                        DebugAppend "Prop=(missing)"
                    End If
                Next i
                Set pst = Nothing
            End If
        Else
            DebugAppend "lpp=" & lpp
        End If
    Else
        DebugAppend "No proplist"
    End If
Else
    DebugAppend "No IShellItem"
End If
'DebugAppend "GenerateInfoTip::Exit->Regular=" & sTip, 11
GenerateInfoTip = Replace(Replace(Replace(Replace(Replace(sTip, vbCrLf, ", "), vbLf, ", "), vbTab, vbNullString), "  ", " "), "  ", " ")
Exit Function
e0:
DebugAppend "GenerateInfoTip->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
Exit Function
e1:
DebugAppend "GenerateInfoTip->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
Resume Next
End Function

Private Function GetPropertyDisplayString(pps As IPropertyStore, ppd As IPropertyDescription, BackupKey As PROPERTYKEY, Optional bFixChars As Boolean = True) As String
'Same as above if you already have the IPropertyDescription (caller is responsible for freeing it too)
Dim lpsz As LongPtr
On Error GoTo e0
If (pps Is Nothing) = False Then
    If (ppd Is Nothing) Then
        DebugAppend "GetPropertyDisplayString->Could not obtain IPropertyDescription, will attempt alternative."
        If IsEqualPKEY(BackupKey, PKEY_Null) = False Then
            Dim vrr As Variant, vbr As Variant
            pps.GetValue BackupKey, vrr
            PropVariantToVariant vrr, vbr
            If (VarType(vbr) And vbArray) = vbArray Then
                Dim i As Long
                For i = LBound(vbr) To UBound(vbr)
                    GetPropertyDisplayString = GetPropertyDisplayString & CStr(vbr(i)) & "; "
                Next i
                GetPropertyDisplayString = Left$(GetPropertyDisplayString, Len(GetPropertyDisplayString) - 2)
            Else
                GetPropertyDisplayString = CStr(vbr)
            End If
        End If
    Else
        PSFormatPropertyValue ObjPtr(pps), ObjPtr(ppd), PDFF_DEFAULT, lpsz
        SysReAllocStringW VarPtr(GetPropertyDisplayString), lpsz
        CoTaskMemFree lpsz
    End If
    If bFixChars Then
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H202A), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H202C), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H200E), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H200F), "")
    End If
Else
    DebugAppend "GetPropertyDisplayString.Error->PropertyStore or PropertyDescription is not set."
        
End If
Exit Function
e0:
DebugAppend "GetPropertyDisplayString.Error->" & Err.Description & ", 0x" & Hex$(Err.Number), 3
End Function


Public Sub RefreshInfoTips()
    If nPr Then
        Dim i As Long
        Dim sTmp As String
        Dim pPrinter As IShellItem
        Dim upi As IParentAndItem
        Dim pidlPar As LongPtr
        Dim pidlRel As LongPtr
        Dim psf As IShellFolder
        For i = 0 To UBound(mPrinters)
            SHCreateItemFromParsingName StrPtr(mPrinters(i).sParsingPath), Nothing, IID_IShellItem, pPrinter
            If (pPrinter Is Nothing) = False Then
            Set upi = pPrinter
                upi.GetParentAndItem pidlPar, psf, pidlRel
                mPrinters(i).sInfoTip = GenerateInfoTip(pPrinter, psf, pidlRel)
            End If
        Next
    End If
End Sub

Private Sub ShowListView()
    'This sub is in *desperate* need of cleanup
    
    If mNoRf = False Then RefreshInfoTips 'Info tips may change often
'CTDBG
    ' Dim nPrDbg As Long = nPr
    '   nPr = 12
    Dim rcCombo As RECT
    Dim rcLVI As RECT
    Dim cxCombo As Long
    Dim cxSet As Long
    Dim cxIdeal As Long
    GetWindowRect hCombo, rcCombo

    Dim mi As MONITORINFO
    Dim hMonitor As LongPtr
    hMonitor = MonitorFromWindow(hCombo, MONITOR_DEFAULTTOPRIMARY)
    mi.cbSize = LenB(mi)
    GetMonitorInfoW hMonitor, mi
    
    cxCombo = (rcCombo.Right - rcCombo.Left)
    If mLimitCX Then
        cxSet = cxCombo
    Else
        cxIdeal = FindMaxWidth()
        If cxIdeal < cxCombo Then
            cxSet = cxCombo
        Else
            cxSet = cxIdeal
        End If
        If (cxList < cxIdeal) And (cxList > cxCombo) Then cxSet = cxList
    End If
    DebugAppend "SizeLV::cxIdeal=" & cxIdeal & ",cxCombo=" & cxCombo & ",cxList=" & cxList
    
    Dim cyIdeal As Long
    Dim cyTile As Long
    Dim cyMaxAvail As Long
    Dim cyaUp As Long, cyaDown As Long
    cyaUp = rcCombo.Top
    cyaDown = mi.rcMonitor.Bottom - (rcCombo.Bottom - rcCombo.Top)
    cyMaxAvail = IIf(cyaDown > cyaUp, cyaDown, cyaUp)
    
    
    cyTile = (32 + 4) * mDPI
    cyIdeal = cyTile * nPr
    
    Dim bSB As Boolean
    If ((cyList > 0) And (cyIdeal > cyList)) Or (cyIdeal > cyMaxAvail) Then 'Scrollbar!!
        DebugAppend "SizeLv::Adjust for scrollbar"
         Dim cxsb As Long
        cxsb = GetSystemMetrics(SM_CXVSCROLL) * mActualZoom
        cxIdeal = cxIdeal + cxsb
        If cxIdeal < cxCombo Then
            cxSet = cxCombo
        Else
            cxSet = cxIdeal
            bSB = True
        End If
        If (cxList < cxIdeal) And (cxList > cxCombo) Then
            cxSet = cxList
            bSB = False 'Not so fast
        End If
    End If
     
    Dim tLVI As LVTILEVIEWINFO
    Dim tsz As Size
    Call SendMessage(hLVW, LVM_SETVIEW, LV_VIEW_TILE, ByVal 0&)
    
    tLVI.cbSize = LenB(tLVI)
    tLVI.dwMask = LVTVIM_COLUMNS Or LVTVIM_TILESIZE '
    tLVI.dwFlags = LVTVIF_FIXEDWIDTH Or LVTVIF_FIXEDHEIGHT
    If bSB Then 'Exclude scrollbar adjustment from tile size *only if* using ideal width
        tsz.CX = (cxSet - (2 * mDPI)) - cxsb
    Else
        tsz.CX = cxSet - (2 * mDPI)
    End If
    tsz.cy = cyTile
    tLVI.SizeTile = tsz
    DebugAppend "Set tile cx=" & tsz.CX
    tLVI.cLines = 2
    Call SendMessage(hLVW, LVM_SETTILEVIEWINFO, 0, tLVI)
    
    
    
    If nPr Then
        Dim lvi As LVITEMW
        Dim i As Long
        Dim nSetSel As Long
        
        For i = 0 To UBound(mPrinters)
            lvi.Mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_IMAGE
            lvi.cchTextMax = Len(mPrinters(i).sName)
            lvi.pszText = StrPtr(mPrinters(i).sName)
            lvi.iImage = mPrinters(i).nIconLV
            lvi.lParam = i
            lvi.iItem = i
            lvi.iSubItem = 0
            mPrinters(i).lvi = CLng(SendMessage(hLVW, LVM_INSERTITEMW, 0, lvi))
            ' If i = mIdxSel Then
            '     nSetSel = mPrinters(i).lvi
            ' End If
            
            lvi.Mask = LVIF_TEXT
            lvi.iItem = mPrinters(i).lvi
            lvi.iSubItem = 1
            lvi.cchTextMax = Len(mPrinters(i).sInfoTip)
            lvi.pszText = StrPtr(mPrinters(i).sInfoTip)
            SendMessage hLVW, LVM_SETITEMW, 0, lvi
            'DebugAppend "LVAddPrinter@ " & mPrinters(i).lvi & "::" & mPrinters(i).sName & ": " & mPrinters(i).sInfoTip
        Next i
'CTDBG
        'SIZING TEST :: Test computer only has 4; test more
        ' For i = 0 To UBound(mPrinters)
        '     lvi.Mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_IMAGE
        '     lvi.cchTextMax = Len(mPrinters(i).sName)
        '     lvi.pszText = StrPtr(mPrinters(i).sName)
        '     lvi.iImage = mPrinters(i).nIconLV
        '     lvi.lParam = i
        '     lvi.iItem = i
        '     lvi.iSubItem = 0
        '     mPrinters(i).lvi = CLng(SendMessage(hLVW, LVM_INSERTITEMW, 0, lvi))
        '     ' If i = mIdxSel Then
        '     '     nSetSel = mPrinters(i).lvi
        '     ' End If
            
        '     lvi.Mask = LVIF_TEXT
        '     lvi.iItem = mPrinters(i).lvi
        '     lvi.iSubItem = 1
        '     lvi.cchTextMax = Len(mPrinters(i).sInfoTip)
        '     lvi.pszText = StrPtr(mPrinters(i).sInfoTip)
        '     SendMessage hLVW, LVM_SETITEMW, 0, lvi
        '     DebugAppend "LVAddPrinter@ " & mPrinters(i).lvi & "::" & mPrinters(i).sName & ": " & mPrinters(i).sInfoTip
        ' Next i
        ' For i = 0 To UBound(mPrinters)
        '     lvi.Mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_IMAGE
        '     lvi.cchTextMax = Len(mPrinters(i).sName)
        '     lvi.pszText = StrPtr(mPrinters(i).sName)
        '     lvi.iImage = mPrinters(i).nIconLV
        '     lvi.lParam = i
        '     lvi.iItem = i
        '     lvi.iSubItem = 0
        '     mPrinters(i).lvi = CLng(SendMessage(hLVW, LVM_INSERTITEMW, 0, lvi))
        '     ' If i = mIdxSel Then
        '     '     nSetSel = mPrinters(i).lvi
        '     ' End If
            
        '     lvi.Mask = LVIF_TEXT
        '     lvi.iItem = mPrinters(i).lvi
        '     lvi.iSubItem = 1
        '     lvi.cchTextMax = Len(mPrinters(i).sInfoTip)
        '     lvi.pszText = StrPtr(mPrinters(i).sInfoTip)
        '     SendMessage hLVW, LVM_SETITEMW, 0, lvi
        '     DebugAppend "LVAddPrinter@ " & mPrinters(i).lvi & "::" & mPrinters(i).sName & ": " & mPrinters(i).sInfoTip
        ' Next i
        
        
        SetTileInfo hLVW
        
        DebugAppend "SizeLV::cxSet=" & cxSet & ",cxCombo=" & cxCombo
        ListView_GetItemRect hLVW, 0, rcLVI, LVIR_ICON
        DebugAppend "SizeLV::IconBounds.Left=" & rcLVI.Left & ",Top=" & rcLVI.Top & ",Right=" & rcLVI.Right & ",Bottom=" & rcLVI.Bottom
        ListView_GetItemRect hLVW, 0, rcLVI, LVIR_LABEL
        DebugAppend "SizeLV::LabelBounds.Left=" & rcLVI.Left & ",Top=" & rcLVI.Top & ",Right=" & rcLVI.Right & ",Bottom=" & rcLVI.Bottom
        ListView_GetItemRect hLVW, 0, rcLVI, LVIR_BOUNDS
        DebugAppend "SizeLV::Bounds.Left=" & rcLVI.Left & ",Top=" & rcLVI.Top & ",Right=" & rcLVI.Right & ",Bottom=" & rcLVI.Bottom
        
        
        Dim cySet As Long
        Dim cyMin As Long
        cyMin = (rcLVI.Bottom - rcLVI.Top) + (2 * smCYEdge)

        cySet = ((rcLVI.Bottom - rcLVI.Top) * nPr) + (3 * nPr)
'CTDBG
        'nPr = nPrDbg
        If (cyList > 0) And (cyList < cySet) Then cySet = cyList
        If cySet < cyMin Then cySet = cyMin
        SetWindowPos hLVW, HWND_TOPMOST, rcCombo.Left, rcCombo.Top + (rcCombo.Bottom - rcCombo.Top), cxSet, cySet, SWP_NOZORDER
        SetWindowPos hLVW, HWND_TOPMOST, rcCombo.Left, rcCombo.Top + (rcCombo.Bottom - rcCombo.Top), rcCombo.Left + cxSet, rcCombo.Top + (rcCombo.Bottom - rcCombo.Top) + cySet, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOREPOSITION

        'Set WS_EX_TOOLWINDOW otherwise it will add an icon to the taskbar
        Dim dwExStyle As WindowStylesEx
        dwExStyle = CLng(GetWindowLong(hLVW, GWL_EXSTYLE))
        dwExStyle = dwExStyle Or WS_EX_PALETTEWINDOW 'includes WS_EX_TOOLWINDOW
        SetWindowLong hLVW, GWL_EXSTYLE, dwExStyle
        
        Dim dwLVExStyle As LVStylesEx
        dwLVExStyle = CLng(SendMessage(hLVW, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0))
        If mTrack Then
            dwLVExStyle = dwLVExStyle Or LVS_EX_TRACKSELECT Or LVS_EX_FULLROWSELECT
            SendMessage hLVW, LVM_SETHOVERTIME, 0, ByVal 1
        Else
            dwLVExStyle = dwLVExStyle And Not LVS_EX_TRACKSELECT
        End If
        Call SendMessage(hLVW, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal dwLVExStyle)
         
        Dim fUp As Boolean

        If (rcCombo.Bottom + cySet) > mi.rcMonitor.Bottom Then
            fUp = True
        End If

        Dim bCBAnim As BOOL
        SystemParametersInfo SPI_GETCOMBOBOXANIMATION, 0, bCBAnim, 0
        If bCBAnim Then
            Const CMS_QANIMATION = 165
            AnimateWindow hLVW, CMS_QANIMATION, IIf(fUp, AW_VER_NEGATIVE, AW_VER_POSITIVE) Or AW_SLIDE
        Else
            ShowWindow hLVW, SW_SHOWNA
        End If

        RedrawWindow hLVW, vbNullPtr, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
        UpdateWindow hLVW
        
        Dim rcActual As RECT
        GetWindowRect hLVW, rcActual
        DebugAppend "ListView final pos=" & rcActual.Left & ", " & rcActual.Top & ", " & rcActual.Right & ", " & rcActual.Bottom
        DebugAppend "Width should be: " & rcCombo.Left + cxSet & " - " & rcCombo.Left
        DebugAppend "Height is: " & (rcActual.Bottom - rcActual.Top) & ", should be: " & cySet
        
        DebugAppend "mIdxSel=" & mIdxSel & ", mIdxDef=" & mIdxDef & ", mPrinters@mIdxSel=" & mPrinters(mIdxSel).sName & ", mPrinters@mIdxDef=" & mPrinters(mIdxDef).sName
        ListView_SetSelectedItem hLVW, mIdxSel
        ListView_EnsureVisible hLVW, mIdxSel, 0
        
        UserControl.SetFocus
        
        SetFocusAPI hLVW

        'SetCapture hLVW
        Set gUCPrinterHookInst = Me
        gUCPrinterHookWindow = hLVW
        gUCPrinterHookHandle = SetWindowsHookEx(WH_MOUSE, AddressOf ucPrinterMouseHookProc, 0, App.ThreadID)
        If gUCPrinterHookHandle = 0 Then
            DebugAppend "Hook error: " & Err.LastDllError
        End If
        
        DebugAppend "ListView count=" & ListView_GetItemCount(hLVW), 2
        

    End If
    
End Sub

Private Function FindMaxWidth() As Long
    Dim i As Long
    Dim cx1 As Long, cx2 As Long
    If nPr = 0 Then
        FindMaxWidth = UserControl.ScaleWidth
    Else
        'Temporarily set the font to bold to calculate width...
        Dim hfOrig As LongPtr, hfBold As LongPtr
        Dim lf As LOGFONT
        hfOrig = SendMessage(hLVW, WM_GETFONT, 0, ByVal 0)
        If hfOrig Then
            GetObjectW hfOrig, LenB(lf), lf
            lf.LFWeight = FW_BOLD
            hfBold = CreateFontIndirect(lf)
            SendMessage hLVW, WM_SETFONT, hfBold, ByVal 1
        End If
        For i = 0 To UBound(mPrinters)
            cx2 = CLng(SendMessage(hLVW, LVM_GETSTRINGWIDTHW, 0, ByVal StrPtr(mPrinters(i).sInfoTip)))
            DebugAppend "CalcMaxWidth(mAZ=" & mActualZoom & "), cx(" & mPrinters(i).sInfoTip & ")=" & cx2 & ", from uc=" & UserControl.TextWidth(mPrinters(i).sInfoTip) & ", from API on UC=" & TextWidthW(UserControl.hDC, mPrinters(i).sInfoTip) & ", from API on LV=" & TextWidthW(GetDC(hLVW), mPrinters(i).sInfoTip)
            If cx2 > cx1 Then cx1 = cx2
            'cx2 = UserControl.TextWidth(mPrinters(i).sName) 'Probably never, but just in case
            cx2 = CLng(SendMessage(hLVW, LVM_GETSTRINGWIDTHW, 0, ByVal StrPtr(mPrinters(i).sName)))
            If cx2 > cx1 Then cx1 = cx2
        Next
        If hfOrig Then
            SendMessage hLVW, WM_SETFONT, hfOrig, ByVal 1
            DeleteObject hfBold
        End If
        FindMaxWidth = cx1 + ((smCXEdge * mActualZoom) * 2) + (32 * mDPI + 8) + (8 * mActualZoom) 'Add border + large icon + margin
        DebugAppend "CalcMaxWidth::Calc with mActualZoom"
 
    End If
End Function

Private Function TextWidthW(ByVal hDC As LongPtr, ByVal sString As String) As Long
  Dim lPtr As LongPtr
  Dim s As Size
  If LenB(sString) Then
    lPtr = StrPtr(sString)
    If Not (lPtr = 0) Then
      GetTextExtentPoint32W hDC, lPtr, Len(sString), s
      TextWidthW = s.CX
    End If
  End If
End Function

Private Sub RefreshPrintersCombo()
    If nPr Then
        SendMessage hCombo, CB_RESETCONTENT, 0, ByVal 0
        Dim nIdx As Long
        Dim i As Long
        For i = 0 To UBound(mPrinters)
            mPrinters(i).cbi = CBX_InsertItem(hCombo, mPrinters(i).sName, mPrinters(i).nIcon, mPrinters(i).nOvr, i)
            DebugAppend "FillCombo mPrinters(" & i & ").cbi=" & mPrinters(i).cbi, 2
            If mPrinters(i).sName = mLabelSelPrev Then
                nIdx = i
            End If
        Next
        SendMessage hCombo, CB_SETCURSEL, nIdx, ByVal 0
    End If
End Sub

Private Sub SetTileInfo(ByVal hWnd As LongPtr)
    Dim tLVT As LVTILEINFO
    Dim lCol() As Long
    ReDim lCol(0)
    lCol(0) = 1
    Dim ct As Long
    Dim i As Long
    ct = CLng(SendMessage(hWnd, LVM_GETITEMCOUNT, 0, ByVal 0&))
    For i = 0 To ct - 1
        tLVT.cbSize = LenB(tLVT)
        tLVT.iItem = i
        tLVT.cColumns = UBound(lCol) + 1
        tLVT.puColumns = VarPtr(lCol(0))
        Call SendMessage(hWnd, LVM_SETTILEINFO, 0, tLVT)
    Next i
End Sub

Private Function GetSysImageList(uFlags As SHGFI_flags) As LongPtr
    Dim sfi As SHFILEINFOW
    Dim sSys As String
    Dim L As Long
    sSys = String$(MAX_PATH, 0)
    L = GetWindowsDirectoryW(StrPtr(sSys), MAX_PATH)
    If L Then
        sSys = Left$(sSys, L)
    Else
        sSys = Left$(Environ("WINDIR"), 3)
    End If
    ' Any valid file system path can be used to retrieve system image list handles.
    GetSysImageList = SHGetFileInfoW(ByVal StrPtr(sSys), 0, sfi, LenB(sfi), SHGFI_SYSICONINDEX Or uFlags)
End Function

Private Function PrinterIndexFromListIndex(ByVal lvi As Long) As Long
    Dim i As Long
    For i = 0 To UBound(mPrinters)
        If mPrinters(i).lvi = lvi Then
            PrinterIndexFromListIndex = i
            Exit Function
        End If
    Next i
    PrinterIndexFromListIndex = -1
End Function

Private Function IsEqualPKEY(pk1 As PROPERTYKEY, pk2 As PROPERTYKEY) As Boolean
    IsEqualPKEY = (CompareMemory(pk1, pk2, LenB(pk1)) = LenB(pk1))
End Function

Private Function OLEFontIsEqual(ByVal Font As StdFont, ByVal FontOther As StdFont) As Boolean
If Font Is Nothing Then
    If FontOther Is Nothing Then OLEFontIsEqual = True
ElseIf FontOther Is Nothing Then
    If Font Is Nothing Then OLEFontIsEqual = True
Else
    If Font.Name = FontOther.Name And Font.Size = FontOther.Size And Font.Charset = FontOther.Charset And Font.Weight = FontOther.Weight And _
    Font.Underline = FontOther.Underline And Font.Italic = FontOther.Italic And Font.Strikethrough = FontOther.Strikethrough Then
        OLEFontIsEqual = True
    End If
End If
End Function

Private Function MakeTrue( _
                ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function

Private Function FindTopLevelWindow() As LongPtr
    Dim hWndCur As LongPtr
    Dim hWndPar As LongPtr
    Dim sClass As String
    Dim nLen As Long
    Const nMaxIter As Long = 99 'Overwhelming likelihood of infinte loop
    Dim i As Long
    Dim IsIDE As Boolean
    Debug.Assert MakeTrue(IsIDE)
    
    hWndCur = UserControl.ContainerHwnd
    Do
        hWndPar = GetParent(hWndCur)
        sClass = String$(255, 0)
        nLen = GetClassNameW(hWndPar, StrPtr(sClass), 255)
        If nLen Then
            sClass = Left$(sClass, nLen)
            If IsIDE Then
                If sClass = "ThunderMain" Then
                    FindTopLevelWindow = hWndPar
                    Exit Function
                End If
            Else
                If sClass = "ThunderRT6Main" Then
                    FindTopLevelWindow = hWndPar
                    Exit Function
                End If
            End If
        End If
        hWndCur = hWndPar
        i = i + 1: If i > nMaxIter Then Exit Do
    Loop
End Function
    
#If TWINBASIC Then
Public Sub CloseDropdown(Optional ByVal hwndFrom As LongPtr)
#Else
Public Sub CloseDropdown(Optional ByVal hwndFrom As Long)
#End If
    DebugAppend "CloseDropdown, bLVVis=" & bLVVis & ", hwndFrom=" & hwndFrom & "|" & hCombo & "|" & hComboCB & "|" & hComboEd
    If bLVVis Then
        ShowWindow hLVW, SW_HIDE
        SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
        bLVVis = False
        UnhookMouse
        If hwndFrom Then
            If (hwndFrom = hCombo) Or (hwndFrom = hComboCB) Then bFlagSuppressReopen = True
        End If
    Else
        SendMessage hCombo, CB_SHOWDROPDOWN, 0, ByVal 0
    End If
End Sub

Private Sub UnhookMouse()
    If gUCPrinterHookHandle Then
        UnhookWindowsHookEx gUCPrinterHookHandle
        gUCPrinterHookHandle = 0
        Set gUCPrinterHookInst = Nothing
        gUCPrinterHookWindow = 0
    End If
End Sub

Private Function GetLVItemlParam(hwndLV As LongPtr, iItem As Long) As LongPtr
  Dim lvi As LVITEM
  
  lvi.Mask = LVIF_PARAM
  lvi.iItem = iItem
  If SendMessage(hwndLV, LVM_GETITEM, 0, lvi) Then
    GetLVItemlParam = lvi.lParam
  End If

End Function

Private Function Subclass2(hWnd As LongPtr, lpfn As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
    If uId = 0 Then uId = hWnd
    Subclass2 = SetWindowSubclass(hWnd, lpfn, uId, dwRefData):      Debug.Assert Subclass2
End Function
Private Function UnSubclass2(hWnd As LongPtr, ByVal lpfn As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpfn, pid)
End Function
Private Function PtrCbWndProc() As LongPtr
    PtrCbWndProc = FARPROC(AddressOf ucPrinterComboWndProc)
End Function
Private Function PtrUCWndProc() As LongPtr
    PtrUCWndProc = FARPROC(AddressOf ucPrinterUserControlWndProc)
End Function
Private Function PtrLVWndProc() As LongPtr
    PtrLVWndProc = FARPROC(AddressOf ucPrinterLVWndProc)
End Function
Private Function FARPROC(ByVal pfn As LongPtr) As LongPtr
    FARPROC = pfn
End Function

#If TWINBASIC Then
Public Function zzCBWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr
#Else
Public Function zzCBWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long) As Long
#End If

Select Case uMsg

    Case WM_NOTIFYFORMAT
        zzCBWndProc = NFR_UNICODE
        Exit Function
        
    Case WM_NCLBUTTONUP
        DebugAppend "CB NCLBU"
        
    Case CB_SHOWDROPDOWN
        'If hWnd = hComboCB Then
        ' DebugAppend "CB_SHOWDROPDOWN"
        ' zzCBWndProc = 1 'Cancel the actual dropdown.
        ' Exit Function
        'End If
        
    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
        DebugAppend "WM_LBUTTONDOWN, bLVVis=" & bLVVis
        mMouseDown = True
    ' SendMessage hComboCB, CB_SHOWDROPDOWN, 0, ByVal 0
        If mListView Then
            If bFlagSuppressReopen Then
                bFlagSuppressReopen = False
            Else
                If bLVVis Then
                    ShowWindow hLVW, 0
                    SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                    bLVVis = False
                    UnhookMouse
                Else
                    ShowListView
                    bLVVis = True
                End If
            End If
            zzCBWndProc = 1 'Cancel the actual dropdown.
            Exit Function
        End If

    Case WM_LBUTTONUP
        mMouseDown = False
    
    Case WM_KEYDOWN
        If wParam = VK_F4 Then
            If mListView Then
                zzCBWndProc = DefSubclassProc(hWnd, WM_LBUTTONDOWN, 1, 1000)
                Exit Function
            End If
        End If
        
    Case WM_COMMAND
        Dim lCode As Long
        lCode = HIWORD(CLng(wParam))
        Select Case lCode
            Case CBN_DROPDOWN
                ' DebugAppend "CBN_DROPDOWN"
                ' SendMessage hComboCB, CB_SHOWDROPDOWN, 0, ByVal 0
                ' ListView1.Visible = True
                ' SetWindowPos ListView1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOREPOSITION
                ' SetFocusAPI ListView1.hWnd
                ' zzCBWndProc = 1 'Cancel the actual dropdown.
                ' Exit Function
        End Select
        
    Case WM_DESTROY
        Call UnSubclass2(hWnd, PtrCbWndProc, uIdSubclass)
End Select

zzCBWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
Exit Function
e0:
DebugAppend "CBWndProc->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)

End Function
#If TWINBASIC Then
Public Function zzUCWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr
#Else
Public Function zzUCWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long) As Long
#End If

Select Case uMsg
' Case CB_SHOWDROPDOWN
    
'     DebugAppend "CB_SHOWDROPDOWN on Parent"
'     zzUCWndProc = 1 'Cancel the actual dropdown.
'     Exit Function

    Case WM_NOTIFYFORMAT
        zzUCWndProc = NFR_UNICODE
        Exit Function
        
    Case WM_KEYDOWN
        If (wParam = VK_ESCAPE) Then
            If bLVVis Then
                CloseDropdown
            End If
        ElseIf (wParam = VK_F4) Then
            If mListView Then
                If bLVVis Then
                    CloseDropdown
                Else
                    zzUCWndProc = DefSubclassProc(hWnd, WM_LBUTTONDOWN, 1, 1000)
                    Exit Function
                End If
            Else
                Dim bDrop As Long
                bDrop = CLng(SendMessage(hCombo, CB_GETDROPPEDSTATE, 0, ByVal 0))
                SendMessage hCombo, CB_SHOWDROPDOWN, IIf(bDrop, 0, 1), ByVal 0
            End If
        End If
        
    Case WM_SETFOCUS
        DebugAppend "WM_SETFOCUS on UCWndProc, hwnd " & hWnd & "|" & hLVW

    Case WM_KILLFOCUS
        DebugAppend "WM_KILLFOCUS on UCWndProc, hwnd " & hWnd & "|" & hLVW
        If bLVVis = True Then
            CloseDropdown hWnd
        End If
    Case WM_NCLBUTTONUP
        DebugAppend "UC NCLBU"
    
    Case WM_NOTIFY
        Dim tNMH As NMHDR
        CopyMemory tNMH, ByVal lParam, LenB(tNMH)
        If tNMH.hwndFrom = hLVW Then
            Select Case tNMH.code
                Case LVN_HOTTRACK
                    If mTrack Then
                        Dim nmlv As NMLISTVIEW
                        CopyMemory nmlv, ByVal lParam, LenB(nmlv)
                        If (nmlv.iItem <> -1) And (nmlv.iItem <> mLastHT) Then
                            mLastHT = nmlv.iItem
                        End If
                        ListView_SetSelectedItem hLVW, mLastHT
                    End If
                
                Case LVN_KEYDOWN
                    If bLVVis Then
                        Dim nmkd As NMLVKEYDOWN
                        CopyMemory nmkd, ByVal lParam, cbnmlvkd
                        Select Case nmkd.wVKey
                            Case VK_ESCAPE, VK_F4
                                DebugAppend "VK_ESCAPE, VK_F4 from LV on UCWndProc"
                                ShowWindow hLVW, SW_HIDE
                                SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                                bLVVis = False
                                UnhookMouse
                        End Select
                    End If
                    
                Case NM_KILLFOCUS
                    If tNMH.hwndFrom = hLVW Then
                        DebugAppend "NM_KILLFOCUS from LV on UCWndProc"
                        ShowWindow hLVW, SW_HIDE
                        SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                        bLVVis = False
                        UnhookMouse
                    End If
                
                Case NM_SETFOCUS
                    If tNMH.hwndFrom = hLVW Then DebugAppend "NM_SETFOCUS from LV on UCWndProc"
                    
                Case NM_CLICK, NM_DBLCLK, NM_RETURN
                    DebugAppend "NM_CLICK from LV on UCWndProc"
                    Dim nLVSel As Long, lp As Long
                    nLVSel = CLng(ListView_GetSelectedItem(hLVW))
                    If nLVSel >= 0 Then
                        mIdxSelPrev = mIdxSel
                        mIdxSel = nLVSel
                        lp = CLng(GetLVItemlParam(hLVW, mIdxSel))
                        DebugAppend "NM_CLICK from hLVW on " & mIdxSel
                        ShowWindow hLVW, SW_HIDE
                        SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                        bLVVis = False
                        UnhookMouse
                        DebugAppend "NM_CLICK lp=" & lp & ", (lp).sName+" & mPrinters(lp).sName
                        SendMessage hCombo, CB_SETCURSEL, mPrinters(lp).cbi, ByVal 0
                        If mIdxSelPrev <> mIdxSel Then
                            RaiseEvent PrinterChanged(mPrinters(lp).sName, mPrinters(lp).sParsingPath, mPrinters(lp).sModel, mPrinters(lp).sLocation, mPrinters(lp).sLastStatus, mPrinters(lp).bDefault)
                        End If
                    Else
                        'Canceled; don't update or raise changed
                        ShowWindow hLVW, SW_HIDE
                        SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                        bLVVis = False
                        UnhookMouse
                    End If
                    SetFocusAPI hCombo
                    
                Case NM_CUSTOMDRAW
                    Dim NMLVCD As NMLVCUSTOMDRAW
                    CopyMemory NMLVCD, ByVal lParam, LenB(NMLVCD)
                    With NMLVCD.NMCD
                        Select Case .dwDrawStage
                            Case CDDS_PREPAINT
                                ' lReturn = CDRF_NOTIFYITEMDRAW
                                ' bHandled = True
                                zzUCWndProc = CDRF_NOTIFYITEMDRAW
                                Exit Function
                                
                            Case CDDS_ITEMPREPAINT
                                Dim nItem As Long
                                nItem = CLng(GetLVItemlParam(hLVW, CLng(.dwItemSpec))) 'PrinterIndexFromListIndex(.dwItemSpec)
                                If nItem >= 0 Then
                                    If mPrinters(nItem).bDefault Then
                                        SelectObject .hDC, hFontBold
                                        'DebugAppend "DrawDefault " & nItem & ", font=" & hFontBold
                                    Else
                                        SelectObject .hDC, hFont
                                        'DebugAppend "DrawStd " & nItem
                                    End If
                                    CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
                                    ' lReturn = CDRF_NOTIFYSUBITEMDRAW Or CDRF_NEWFONT
                                    ' bHandled = True
                                    zzUCWndProc = CDRF_NEWFONT
                                    Exit Function
                                End If
                        End Select
                    End With
            End Select
        End If
        
        
    Case WM_COMMAND
        If lParam = hCombo Then
        Dim lCode As Long
        lCode = HIWORD(CLng(wParam))
        Select Case lCode
            Case CBN_SELCHANGE
                Dim nIdx As Long
                Dim nSel As Long
                nSel = CLng(SendMessage(hCombo, CB_GETCURSEL, 0, ByVal 0))
                nIdx = -1
                nIdx = CLng(GetCBXItemlParam(hCombo, nSel))
                If nIdx >= 0 Then
                    RaiseEvent PrinterChanged(mPrinters(nIdx).sName, mPrinters(nIdx).sParsingPath, mPrinters(nIdx).sModel, mPrinters(nIdx).sLocation, mPrinters(nIdx).sLastStatus, mPrinters(nIdx).bDefault)
                End If
                    
            Case CBN_DROPDOWN
                DebugAppend "CBN_DROPDOWN on Parent"
                zzUCWndProc = 1 'Cancel the actual dropdown.
                Exit Function
        End Select
        End If
    Case WM_DESTROY
        Call UnSubclass2(hWnd, PtrUCWndProc, uIdSubclass)
    
    End Select
    zzUCWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    Exit Function
e0:
    DebugAppend "UCWndProc->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)

End Function

#If TWINBASIC Then
Public Function zzLVWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr
#Else
Public Function zzLVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long) As Long
#End If

Select Case uMsg
    
    Case WM_NOTIFYFORMAT
        zzLVWndProc = NFR_UNICODE
        Exit Function
        
    ' Case OCM_NOTIFY
    '     DebugAppend "Receiving OCM_NOTIFY"
    Case WM_KEYUP
        DebugAppend "KeyUp on LVWndProc"
        
    Case WM_NOTIFY
        Dim tNMH As NMHDR
        CopyMemory tNMH, ByVal lParam, LenB(tNMH)
        If tNMH.hwndFrom = hLVW Then
            DebugAppend "Receiving Notify from hLVW in LVWndProc"
            Select Case tNMH.code
                Case NM_KILLFOCUS
                    If tNMH.hwndFrom = hLVW Then
                        DebugAppend "NM_KILLFOCUS from LV on LVWndProc"
                        ShowWindow hLVW, SW_HIDE
                        SendMessage hLVW, LVM_DELETEALLITEMS, 0, ByVal 0
                        bLVVis = False
                        UnhookMouse
                    End If
                            
                Case NM_SETFOCUS
                    If tNMH.hwndFrom = hLVW Then DebugAppend "NM_SETFOCUS from LV on LVWndProc"
            End Select
        End If
    Case WM_SETFOCUS
        DebugAppend "WM_SETFOCUS on LVWndProc, hwnd " & hWnd & "|" & hLVW

    Case WM_KILLFOCUS
        DebugAppend "WM_KILLFOCUS on LVWndProc, hwnd " & hWnd & "|" & hLVW
        
    ' Case WM_NCPAINT
    '     Dim hdc As LongPtr
    '     hdc = GetDCEx(hWnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN)
    '     Dim rc As RECT
    '     GetClientRect hWnd, rc
    '     DrawThemeBackground hTheme, hdc, CP_BORDER, CBXS_HOT, rc, vbNullPtr
        
    '     ReleaseDC(hWnd, hdc)
        
    Case WM_DESTROY
        Call UnSubclass2(hWnd, PtrLVWndProc, uIdSubclass)
        


End Select
zzLVWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

