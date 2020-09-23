VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCapture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3540
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2940
      Top             =   2100
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   3600
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DropChar() As udtChar
Dim iScreenWidth As Integer
Dim iScreenHeight As Integer
Private Sub Form_Click()
    Unload Me
    End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Static X0 As Integer, Y0 As Integer
'-----------------------------------------------------------------
    If (RunMode = RM_NORMAL) Then           ' Determine screen saver mode
        If ((X0 = 0) And (Y0 = 0)) Or _
           ((Abs(X0 - x) < 5) And (Abs(Y0 - y) < 5)) Then ' small mouse movement...
            X0 = x                          ' Save current x coordinate
            Y0 = y                          ' Save current y coordinate
            Exit Sub                        ' Exit
        End If
    
        Unload Me
        End ' Large mouse movement (terminate screensaver)
    End If
End Sub
Function GetIPAddress() As String
Dim sHostName As String * 256
Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
      " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
    End If
    
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    
    If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
      "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
    End If
    
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4
    
    ReDim tmpIPAddr(1 To HOST.hLen)
    
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    
    For i = 1 To HOST.hLen
    sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup

End Function

Public Function GetIPHostName() As String
Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
      GetIPHostName = ""
      Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPHostName = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
      " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function

Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H100 And &HFF&

End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function



Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
      MsgBox "Socket error occurred in Cleanup."
    End If

End Sub



Public Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim sLoByte As String
Dim sHiByte As String
    
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
      MsgBox "This application requires a minimum of " & _
      CStr(MIN_SOCKETS_REQD) & " supported sockets."
      SocketsInitialize = False
      Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
    (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
      " is not supported by 32-bit Windows Sockets."
      SocketsInitialize = False
      Exit Function
    End If
    
    'must be OK, so lets do it
    SocketsInitialize = True

End Function

Private Sub Form_Load()
    
    GetParms
    
    If bTransparent Then picCapture.Picture = CaptureScreen()

    If (RunMode = RM_NORMAL) Then ShowCursor 0
    InitDeskDC DeskDC, DeskBmp, DispRec
    Randomize Timer
    
    'picBlank.Move 0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY
    'picText.Move 0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY
    picText.Font = sFont
    
    'Me.Show
    BitBlt Me.hdc, 0, 0, picText.Width, picText.Height, picCapture.hdc, 0, 0, vbSrcCopy
    
    SetupChars
    
    Timer1.Interval = iTiming
    Timer1.Enabled = True

End Sub

Private Sub Form_Resize()
    Timer1.Enabled = False
    iScreenWidth = Me.Width \ Screen.TwipsPerPixelX
    iScreenHeight = Me.Height \ Screen.TwipsPerPixelY
    picBlank.Move 0, 0, iScreenWidth, iScreenHeight
    picText.Move 0, 0, iScreenWidth, iScreenHeight
    Erase DropChar
    SetupChars
    Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------
    Dim Idx As Integer                          ' Array index
'-----------------------------------------------------------------
    ' [* YOU MUST TURN OFF THE TIMER BEFORE DESTROYING THE SPRITE OBJECT *]
    Timer1.Enabled = False                     ' [* YOU MAY DEADLOCK!!! *]
'   Set gSpriteCollection = Nothing             ' Not sure if this would work...

    DelDeskDC DeskDC                            ' Cleanup the DeskDC (Memleak will occure if not done)
    
    If (RunMode = RM_NORMAL) Then ShowCursor -1 ' Show MousePointer
    Screen.MousePointer = vbDefault             ' Reset MousePointer
    End
'-----------------------------------------------------------------
End Sub
Private Sub Timer1_Timer()
Dim iCol As Integer
Dim iSize As Integer
Dim iX As Integer

    If Not bTransparent Then
        BitBlt picText.hdc, 0, 0, picText.Width, picText.Height, picBlank.hdc, 0, 0, vbSrcCopy 'send blank screen to bufer
    Else
        BitBlt picText.hdc, 0, 0, picText.Width, picText.Height, picCapture.hdc, 0, 0, vbSrcCopy 'send captured screen to buffer
    End If
    For iX = 0 To iNumDropChars - 1
        With DropChar(iX)
            picText.FontSize = .Size
            
            If bItalic Then picText.FontItalic = True
            If bBold Then picText.FontBold = True
            
            If bDropShadow And bTransparent Then
                SetTextColor picText.hdc, vbBlack
                TextOut picText.hdc, .xPos + 1, .yPos + 1, .Text, .TextLen
            End If
            
            'colored text
            SetTextColor picText.hdc, .Color
            TextOut picText.hdc, .xPos, .yPos, .Text, .TextLen
            
            If bVert Then
                .yPos = .yPos + .Speed
                'wavey
                '.xPos = .xPos + (Int(Rnd * 4) - 2)
                If .yPos > (picText.Height) Then
                    .yPos = .TextHeight * -1
                    .xPos = Int(Rnd * iScreenWidth)
                End If
            Else
                .xPos = .xPos - .Speed
                'wavey
                '.yPos = .yPos + (Int(Rnd * 4) - 2)
                If (.xPos + .TextWidth) < 0 Then
                    .xPos = iScreenWidth + .TextWidth
                    .yPos = Int(Rnd * iScreenHeight)
                End If
            End If
        End With
    Next iX
    BitBlt Me.hdc, 0, 0, picText.Width, picText.Height, picText.hdc, 0, 0, vbSrcCopy 'buffer to screen
    
End Sub
Sub SetupChars()
Dim iX As Integer
Dim lCol As Long
Dim sRunning() As String
Dim iSelectColor As Long
    sRunning = ReturnRunningProcesses
    If iContent = 2 Then iNumDropChars = UBound(sRunning) + 1 'number of scroll items is equal to running processes
    ReDim DropChar(iNumDropChars)
    For iX = 0 To iNumDropChars - 1
        With DropChar(iX)
            Select Case iContent
            Case 0
                .Text = sFreeText
            Case 1
                .Text = GetIPAddress
            Case 2
                .Text = LCase(sRunning(iX))
            Case 3
                .Text = GetIPHostName
            End Select
            .Size = Int(Rnd * iFontSize) + 8
            .Speed = Int(Rnd * iMaxSpeed) + 2
            lCol = .Speed * (255 / iMaxSpeed)
            If bRandomColors Then
                .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            Else
                Select Case iColor
                    Case 0 'white
                        .Color = RGB(lCol, lCol, lCol)
                    Case 1 'green
                        .Color = RGB(lCol, lCol * 2, lCol)
                    Case 2 'red
                        .Color = RGB(lCol * 2, lCol, lCol)
                    Case 3 'blue
                        .Color = RGB(lCol, lCol, lCol * 2)
                    Case 4 'yellow
                        .Color = RGB(lCol * 2, lCol * 2, lCol)
                    Case 5 'cyan
                        .Color = RGB(lCol, lCol * 2, lCol * 2)
                End Select
                
            End If
            picText.FontSize = .Size
            .TextLen = Len(.Text)
            .TextWidth = picText.TextWidth(.Text)
            .TextHeight = picText.TextHeight(.Text)
            If bVert Then
                .xPos = Int(Rnd * iScreenWidth)
                .yPos = Int(Rnd * iScreenHeight) * -1
            Else
                .xPos = (iScreenWidth + .TextWidth) + Int(Rnd * iScreenWidth)
                .yPos = Int(Rnd * iScreenHeight)
            End If
        End With
    Next
End Sub
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE

   If Client Then
      hDCSrc = GetDC(hWndSrc)
   Else
      hDCSrc = GetWindowDC(hWndSrc)
   End If


   hDCMemory = CreateCompatibleDC(hDCSrc)
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

   hBmp = SelectObject(hDCMemory, hBmpPrev)

   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   Set CreateBitmapPicture = IPic
End Function

Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long

   hWndScreen = GetDesktopWindow()

   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Function ReturnRunningProcesses() As String()
Const MaxPIDs = 256
Dim dwPIDs(1 To MaxPIDs) As Long, szPIDs As Long, cb2 As Long
Dim i As Long, hPr As Long, cproc As Long
Dim Modules(1 To MaxPIDs) As Long, lr As Long
Dim mName As String * MAX_PATH
Dim sT() As String

    ReDim sT(0)
    szPIDs = MaxPIDs * 4
    If EnumProcesses(dwPIDs(1), szPIDs, cb2) <> 0 Then
        cproc = cb2 / 4
        For i = 1 To cproc
            hPr = OpenProcess(PROCESS_QUERY_INFORMATION _
               Or PROCESS_VM_READ, 0, dwPIDs(i))
            If hPr <> 0 Then
                If EnumProcessModules(hPr, Modules(1), MaxPIDs, cb2) <> 0 Then
                    mName = Space(MAX_PATH)
                    lr = GetModuleFileNameExA(hPr, Modules(1), mName, MAX_PATH)
                    If (mName <> "") Then
                        ReDim Preserve sT(UBound(sT) + 1)
                        sT(UBound(sT) - 1) = Mid(C23(mName), InStrRev(C23(mName), "\") + 1) & " (" & dwPIDs(i) & ")"
                    End If
                End If
            End If
            CloseHandle (hPr)
        Next i
    End If
    If UBound(sT) > 0 Then ReDim Preserve sT(UBound(sT) - 1)
    ReturnRunningProcesses = sT
End Function

Function C23(A As String) As String
Dim i As Long
    i = InStr(1, A, Chr(0), vbBinaryCompare)
    If (i <> 0) Then
        C23 = Mid(A, 1, i - 1)
    Else
        C23 = ""
    End If
End Function
