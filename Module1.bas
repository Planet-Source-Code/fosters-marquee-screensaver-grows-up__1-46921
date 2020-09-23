Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1
Public Const MAX_PATH = 260
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Type HOSTENT
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLen As Integer
  hAddrList As Long
End Type

Public Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To MAX_WSADescription) As Byte
  szSystemStatus(0 To MAX_WSASYSStatus) As Byte
  wMaxSockets As Integer
  wMaxUDPDG As Integer
  dwVendorInfo As Long
End Type

Public Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type
Public Const SRCCOPY = &HCC0020

Public Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Dim bStop As Boolean
Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

Public Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Public Type udtChar
    xPos As Integer
    yPos As Integer
    Size As Integer
    Color As Long
    Speed As Long
    Text As String
    TextLen As Long
    TextWidth As Integer
    TextHeight As Integer
End Type

Public Type LUID
   lowpart As Long
   highpart As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public iTiming As Integer
Public iMaxSpeed As Integer
Public iFontSize As Integer
Public iNumDropChars As Integer
Public sFreeText As String
Public bDropShadow As Boolean
Public bBold As Boolean
Public bItalic As Boolean
Public sFont As String
Public bVert As Boolean
Public iContent As Integer
Public iColor As Integer
Public bTransparent As Boolean
Public bRandomColors As Boolean

Sub InitialiseRegistry()
    If Len(GetSetting(App.Title, "Settings", "Timing")) = 0 Then SaveSetting App.Title, "Settings", "Timing", 20
    If Len(GetSetting(App.Title, "Settings", "Speed")) = 0 Then SaveSetting App.Title, "Settings", "Speed", 12
    If Len(GetSetting(App.Title, "Settings", "NumItems")) = 0 Then SaveSetting App.Title, "Settings", "NumItems", 20
    If Len(GetSetting(App.Title, "Settings", "FontSize")) = 0 Then SaveSetting App.Title, "Settings", "FontSize", 28
    If Len(GetSetting(App.Title, "Settings", "FontBold")) = 0 Then SaveSetting App.Title, "Settings", "FontBold", 0
    If Len(GetSetting(App.Title, "Settings", "FontItalic")) = 0 Then SaveSetting App.Title, "Settings", "FontItalic", 0
    If Len(GetSetting(App.Title, "Settings", "DropShadow")) = 0 Then SaveSetting App.Title, "Settings", "DropShadow", 0
    If Len(GetSetting(App.Title, "Settings", "Font")) = 0 Then SaveSetting App.Title, "Settings", "Font", "Arial"
    If Len(GetSetting(App.Title, "Settings", "Color")) = 0 Then SaveSetting App.Title, "Settings", "Color", 1
    If Len(GetSetting(App.Title, "Settings", "Direction")) = 0 Then SaveSetting App.Title, "Settings", "Direction", 0
    If Len(GetSetting(App.Title, "Settings", "Screencolor")) = 0 Then SaveSetting App.Title, "Settings", "ScreenColor", 0
    If Len(GetSetting(App.Title, "Settings", "ColorType")) = 0 Then SaveSetting App.Title, "Settings", "ColorType", 0
    If Len(GetSetting(App.Title, "Settings", "Content")) = 0 Then SaveSetting App.Title, "Settings", "Content", 0
    If Len(GetSetting(App.Title, "Settings", "FreeText")) = 0 Then SaveSetting App.Title, "Settings", "FreeText", "Marquee Screensaver"
End Sub

Sub GetParms()


    iTiming = CInt(GetSetting(App.Title, "Settings", "Timing"))
    iMaxSpeed = CInt(GetSetting(App.Title, "Settings", "Speed"))
    iNumDropChars = CInt(GetSetting(App.Title, "Settings", "NumItems"))
    iFontSize = CInt(GetSetting(App.Title, "Settings", "FontSize"))
    sFreeText = GetSetting(App.Title, "Settings", "FreeText")
    bDropShadow = IIf(CInt(GetSetting(App.Title, "Settings", "DropShadow")) = 0, False, True)
    bBold = IIf(CInt(GetSetting(App.Title, "Settings", "FontBold")) = 0, False, True)
    bItalic = IIf(CInt(GetSetting(App.Title, "Settings", "FontItalic")) = 0, False, True)
    sFont = GetSetting(App.Title, "Settings", "Font")
    iColor = CInt(GetSetting(App.Title, "Settings", "Color"))
    bVert = IIf(CInt(GetSetting(App.Title, "Settings", "Direction")) = 0, False, True)
    bTransparent = IIf(CInt(GetSetting(App.Title, "Settings", "ScreenColor")) = 0, False, True)
    bRandomColors = IIf(CInt(GetSetting(App.Title, "Settings", "ColorType")) = 0, False, True)
    iContent = CInt(GetSetting(App.Title, "Settings", "Content"))
End Sub
