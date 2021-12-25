VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Capture and Print"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   11745
   StartUpPosition =   3  '系統預設值
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   3720
      TabIndex        =   8
      Top             =   3720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ExecProg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "IntTime"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Setting"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Path"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GamePath"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Keep a life"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Debug"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ProgStatus"
      Tab(4).Control(1)=   "DebugText2"
      Tab(4).Control(2)=   "DebugText1"
      Tab(4).Control(3)=   "SkillPicture"
      Tab(4).Control(4)=   "PPPicture"
      Tab(4).ControlCount=   5
      Begin VB.TextBox Text1 
         Height          =   2535
         Left            =   -74640
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox IntTime 
         Height          =   270
         Left            =   720
         TabIndex        =   18
         Text            =   "250"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox ProgStatus 
         Height          =   375
         Left            =   -74520
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox DebugText2 
         Height          =   375
         Left            =   -74520
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox DebugText1 
         Height          =   375
         Left            =   -74520
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.PictureBox SkillPicture 
         Height          =   3015
         Left            =   -69480
         ScaleHeight     =   2955
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox PPPicture 
         Height          =   735
         Left            =   -74520
         ScaleHeight     =   675
         ScaleWidth      =   2115
         TabIndex        =   13
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox GamePath 
         Height          =   375
         Left            =   -74280
         TabIndex        =   12
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton ExecProg 
         Caption         =   "Run"
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "ms"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   6000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save image"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   3120
      ScaleHeight     =   3315
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear picture box"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print picturebox"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Capture active window after 2 s"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Capture client area of form"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Capture form including title bar"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capture entire screen"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0


' For Screen Capture
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

' For Key in

' For Process control
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200

Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const BELOW_NORMAL_PRIORITY_CLASS = 16384
Private Const ABOVE_NORMAL_PRIORITY_CLASS = 32768
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100


' Common Parameter
Dim r As Long
Dim hFrgWnd As Long
Dim hRANDC As Long


Private Sub Check1_Click()
    Timer1.Interval = Val(IntTime.Text)
    Timer1.Enabled = Check1.Enabled
    
End Sub

' Capture the entire screen.
Private Sub Command1_Click()
   Set Picture1.Picture = CaptureScreen()
End Sub

' Capture the entire form including title and border.
Private Sub Command2_Click()
    Set Picture1.Picture = CaptureForm(Me)
End Sub

' Capture the client area of the form.
Private Sub Command3_Click()
    Set Picture1.Picture = CaptureClient(Me)
End Sub

' Capture the active window after two seconds.
Private Sub Command4_Click()
    MsgBox "Two seconds after you close this dialog the active window will be captured."
    ' Wait for two seconds.
    Dim EndTime As Date
    EndTime = DateAdd("s", 2, Now)
    Do Until Now > EndTime
       DoEvents
       Loop
    Set Picture1.Picture = CaptureActiveWindow()
    ' Set focus back to form.
    Me.SetFocus
End Sub

' Print the current contents of the picture box.
    Private Sub Command5_Click()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
End Sub

' Clear out the picture box.
Private Sub Command6_Click()
    Set Picture1.Picture = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CreateBitmapPicture
'    - Creates a bitmap type Picture object from a bitmap and
'      palette.
'
' hBmp
'    - Handle to a bitmap.
'
' hPal
'    - Handle to a Palette.
'    - Can be null if the bitmap doesn't use a palette.
'
' Returns
'    - Returns a Picture object containing the bitmap.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
   ' IPicture requires a reference to "Standard OLE Types."
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   ' Fill in with IDispatch Interface ID.
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Fill Pic with necessary parts.
   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   ' Return the new Picture object.
   Set CreateBitmapPicture = IPic
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureWindow
'    - Captures any portion of a window.
'
' hWndSrc
'    - Handle to the window to be captured.
'
' Client
'    - If True CaptureWindow captures from the client area of the
'      window.
'    - If False CaptureWindow captures from the entire window.
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture.
'    - Dimensions need to be specified in pixels.
'
' Returns
'    - Returns a Picture object containing a bitmap of the specified
'      portion of the window that was captured.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
'

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

   ' Depending on the value of Client get the proper device context.
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window.
   End If

   ' Create a memory device context for the copy process.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties.
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   ' If the screen has a palette make a copy and realize it.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC.
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureScreen
'    - Captures the entire screen.
'
' Returns
'    - Returns a Picture object containing a bitmap of the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long

   ' Get a handle to the desktop window.
   hWndScreen = GetDesktopWindow()

   ' Call CaptureWindow to capture the entire desktop give the handle
   ' and return the resulting Picture object.

   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
'    - Captures an entire form including title bar and border.
'
' frmSrc
'    - The Form object to capture.
'
' Returns
'    - Returns a Picture object containing a bitmap of the entire
'      form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureForm(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the entire form given its window
   ' handle and then return the resulting Picture object.
   Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureClient
'    - Captures the client area of a form.
'
' frmSrc
'    - The Form object to capture.
'
' Returns
'    - Returns a Picture object containing a bitmap of the form's
'      client area.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureClient(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the client area of the form given
   ' its window handle and return the resulting Picture object.
   Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureActiveWindow
'    - Captures the currently active window on the screen.
'
' Returns
'    - Returns a Picture object containing a bitmap of the active
'      window.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureActiveWindow() As Picture

    Dim hWndActive As Long
    Dim r As Long
    
    Dim RectActive As RECT
    
    ' Get a handle to the active/foreground window.
    hWndActive = GetForegroundWindow()
    
    ' Get the dimensions of the window.
    r = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given its
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintPictureToFitPage
'    - Prints a Picture object as big as possible.
'
' Prn
'    - Destination Printer object.
'
' Pic
'    - Source Picture object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    
    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation.
    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait   ' Taller than wide.
    Else
        Prn.Orientation = vbPRORLandscape  ' Wider than tall.
    End If
    
    ' Calculate device independent Width-to-Height ratio for picture.
    PicRatio = Pic.Width / Pic.Height
    
    ' Calculate the dimentions of the printable area in HiMetric.
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to Height ratio for printer.
    PrnRatio = PrnWidth / PrnHeight
    
    ' Scale the output to the printable area.
    If PicRatio >= PrnRatio Then
        ' Scale picture to fit full width of printable area.
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area.
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    ' Print the picture using the PaintPicture method.
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub
'-------------------------------------------------------------------


Private Sub Command7_Click()

CommonDialog1.DefaultExt = ".BMP"
CommonDialog1.Filter = "Bitmap Image (*.bmp)|*.bmp"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    SavePicture Picture1.Picture, CommonDialog1.FileName
End If

End Sub


Private Sub Command8_Click()

End Sub

Private Sub ExecProg_Click()
    Dim r As Double
    r = ShellExecute(Me.hwnd, "open", GamePath.Text, "", "", nShowCmd)
    r = ShellExecute(, vbNormalFocus)
    
End Sub

Private Sub Form_Load()
    Dim ProcessID As Long
    Dim hProc As Long
    Const fdwAccess1 As Long = PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION

    hRANDC = 0

    ProcessID = GetCurrentProcessId()
    hProc = OpenProcess(fdwAccess1, 0&, ProcessID)

    If hProc Then
        ' Attempt to set new priority.
        Call SetPriorityClass(hProc, BELOW_NORMAL_PRIORITY_CLASS)
        Call CloseHandle(hProc)
    End If
    
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Timer1_Timer()
    Dim r As Long
    Dim hDCSrc As Long
    Dim hWndSrc As Long
           
    ProgStatus.Text = "Can not detect RAN"
    
    hFrgWnd = GetForegroundWindow()
    hDCSrc = GetWindowDC(hWndSrc)
    r = BitBlt(PPPicture.hDC, 0, 0, PPPicture.Picture.Width, PPPicture.Picture.Height, hDCSrc, 0, 0, vbSrcCopy)
    r = ReleaseDC(hWndSrc, hDCSrc)
    
End Sub
