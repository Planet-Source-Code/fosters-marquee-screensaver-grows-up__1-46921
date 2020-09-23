VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Custom Marquee Screensaver"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6930
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5160
      TabIndex        =   31
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   5160
      TabIndex        =   30
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text Colour"
      Height          =   915
      Left            =   780
      TabIndex        =   28
      Top             =   5880
      Width           =   4215
      Begin VB.PictureBox picColorColl 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   1515
         TabIndex        =   34
         Top             =   240
         Width           =   1545
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   1260
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   40
            Top             =   60
            Width           =   195
         End
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1020
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   39
            Top             =   60
            Width           =   195
         End
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   780
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   38
            Top             =   60
            Width           =   195
         End
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   540
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   37
            Top             =   60
            Width           =   195
         End
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   300
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   36
            Top             =   60
            Width           =   195
         End
         Begin VB.PictureBox picColorSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   60
            ScaleHeight     =   165
            ScaleWidth      =   165
            TabIndex        =   35
            Top             =   60
            Width           =   195
         End
      End
      Begin VB.OptionButton optColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Random Colour"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   540
         Width           =   2295
      End
      Begin VB.OptionButton optColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Monochrome (Multi Shade)"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Screen"
      Height          =   675
      Left            =   780
      TabIndex        =   26
      Top             =   3180
      Width           =   4215
      Begin VB.CheckBox chkDropShadow 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Drop Shadow Text"
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   300
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton optScreen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transparent"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton optScreen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Black"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Content"
      Height          =   1935
      Left            =   780
      TabIndex        =   25
      Top             =   3900
      Width           =   4215
      Begin VB.TextBox txtNumItems 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1980
         TabIndex        =   33
         Text            =   "20"
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtFreeText 
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Text            =   "Marquee Screensaver"
         Top             =   1080
         Width           =   3855
      End
      Begin VB.OptionButton optContent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Host Name"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   14
         Top             =   540
         Width           =   1695
      End
      Begin VB.OptionButton optContent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Running Processes"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   300
         Width           =   1695
      End
      Begin VB.OptionButton optContent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP Address"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton optContent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Free text"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Scroll Items"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Free text"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Font"
      Height          =   1635
      Left            =   780
      TabIndex        =   22
      Top             =   1500
      Width           =   4215
      Begin VB.TextBox txtMaxFontSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3540
         TabIndex        =   7
         Text            =   "28"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkItalic 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Italic"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1260
         Width           =   675
      End
      Begin VB.CheckBox chkBold 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bold"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1260
         Width           =   675
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Maximum Font Size"
         Height          =   195
         Left            =   1980
         TabIndex        =   29
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label lblFontSample 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marquee Screensaver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   435
         Left            =   240
         TabIndex        =   24
         Top             =   660
         Width           =   3795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Speed and Direction"
      Height          =   1395
      Left            =   780
      TabIndex        =   18
      Top             =   60
      Width           =   4215
      Begin VB.OptionButton optDir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vertical"
         Height          =   195
         Index           =   1
         Left            =   3180
         TabIndex        =   3
         Top             =   1020
         Width           =   915
      End
      Begin VB.OptionButton optDir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horizontal"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   1020
         Width           =   1035
      End
      Begin VB.TextBox txtTimer 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Text            =   "20"
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtMaxSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Text            =   "12"
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Timing Interval (ms)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum speed"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll Direction"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1020
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   1425
      Left            =   240
      Picture         =   "Form2.frx":0000
      Top             =   5340
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":1686
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBold_Click()
    lblFontSample.FontBold = (chkBold.Value = vbChecked)
End Sub

Private Sub chkItalic_Click()
    lblFontSample.FontItalic = (chkItalic.Value = vbChecked)
End Sub

Private Sub cmbFont_Click()
    lblFontSample.FontName = cmbFont.Text
End Sub

Private Sub Command1_Click()
Dim iX As Integer
    If Not IsNumeric(txtTimer) Then
        MsgBox "Timing should be a numeric", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If CInt(txtTimer) < 0 Then
        MsgBox "Timing should be a positive integer", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtMaxSpeed) Then
        MsgBox "Speed should be a numeric", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If CInt(txtMaxSpeed) < 0 Then
        MsgBox "Speed should be a positive integer", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtMaxFontSize) Then
        MsgBox "Font size should be a numeric", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If CInt(txtMaxFontSize) < 0 Then
        MsgBox "Font size should be a positive integer", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtNumItems) Then
        MsgBox "Number of scroll items should be a numeric", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    
    If CInt(txtNumItems) < 0 Then
        MsgBox "Number of scroll items should be a positive integer", vbExclamation, App.Title
        txtTimer.SetFocus
        Exit Sub
    End If
    

    'save stuff
    SaveSetting App.Title, "Settings", "Timing", txtTimer
    SaveSetting App.Title, "Settings", "Speed", txtMaxSpeed
    SaveSetting App.Title, "Settings", "NumItems", txtNumItems
    SaveSetting App.Title, "Settings", "FontSize", txtMaxFontSize
    SaveSetting App.Title, "Settings", "FontBold", chkBold.Value
    SaveSetting App.Title, "Settings", "FontItalic", chkItalic.Value
    SaveSetting App.Title, "Settings", "DropShadow", chkDropShadow.Value
    SaveSetting App.Title, "Settings", "Font", cmbFont.Text
    
    For iX = 0 To picColorSelect.Count - 1
        If picColorSelect(iX).Height = 195 Then
            SaveSetting App.Title, "Settings", "Color", iX
            Exit For
        End If
    Next
    
    SaveSetting App.Title, "Settings", "Direction", IIf(optDir(0).Value, 0, 1)
    SaveSetting App.Title, "Settings", "ScreenColor", IIf(optScreen(0).Value, 0, 1)
    SaveSetting App.Title, "Settings", "ColorType", IIf(optColor(0).Value, 0, 1)
    
    For iX = 0 To optContent.Count - 1
        If optContent(iX).Value Then
            SaveSetting App.Title, "Settings", "Content", iX
            Exit For
        End If
    Next
    
    SaveSetting App.Title, "Settings", "FreeText", txtFreeText
    
    Unload Me
    End
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
Dim iX As Integer

    cmbFont.Clear
    For iX = 0 To Printer.FontCount - 1
        cmbFont.AddItem Printer.Fonts(iX)
    Next iX

    txtTimer = GetSetting(App.Title, "Settings", "Timing")
    txtMaxSpeed = GetSetting(App.Title, "Settings", "Speed")
    txtNumItems = GetSetting(App.Title, "Settings", "NumItems")
    txtMaxFontSize = GetSetting(App.Title, "Settings", "FontSize")
    txtFreeText = GetSetting(App.Title, "Settings", "FreeText")
    
    chkDropShadow.Value = CInt(GetSetting(App.Title, "Settings", "DropShadow"))
    
    
    chkBold.Value = CInt(GetSetting(App.Title, "Settings", "FontBold"))
    chkItalic.Value = CInt(GetSetting(App.Title, "Settings", "FontItalic"))
    cmbFont.Text = GetSetting(App.Title, "Settings", "Font")
    lblFontSample.Font = cmbFont.Text
    
    picColorSelect_Click CInt(GetSetting(App.Title, "Settings", "Color"))
    
    optDir(CInt(GetSetting(App.Title, "Settings", "Direction"))).Value = True
    optScreen_Click CInt(GetSetting(App.Title, "Settings", "ScreenColor"))
    optScreen(CInt(GetSetting(App.Title, "Settings", "ScreenColor"))).Value = True
    optColor_Click CInt(GetSetting(App.Title, "Settings", "ColorType"))
    optColor(CInt(GetSetting(App.Title, "Settings", "ColorType"))).Value = True
    optContent_Click CInt(GetSetting(App.Title, "Settings", "Content"))
    optContent(CInt(GetSetting(App.Title, "Settings", "Content"))).Value = True
    


End Sub

Private Sub optColor_Click(Index As Integer)
    picColorColl.Visible = optColor(0).Value
End Sub

Private Sub optContent_Click(Index As Integer)
    txtFreeText.Enabled = optContent(0).Value
    
End Sub

Private Sub optScreen_Click(Index As Integer)
    chkDropShadow.Visible = optScreen(1).Value
End Sub

Private Sub picColorSelect_Click(Index As Integer)
Dim iX As Integer
    For iX = 0 To picColorSelect.Count - 1
        If iX = Index Then
            picColorSelect(iX).Height = 195
        Else
            picColorSelect(iX).Height = 100
        End If
    Next
End Sub
