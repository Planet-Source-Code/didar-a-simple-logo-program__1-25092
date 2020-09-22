VERSION 5.00
Begin VB.Form frmSrc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "General Corporation"
   ClientHeight    =   405
   ClientLeft      =   5865
   ClientTop       =   0
   ClientWidth     =   1785
   Icon            =   "frmSrc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2280
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   1440
      Top             =   2760
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   -5880
      ScaleHeight     =   1635
      ScaleWidth      =   8835
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Corporation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   5
         Top             =   120
         Width           =   1710
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40

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


Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long


Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
    ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long _
    ) As Long


Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, _
    ByVal iCapabilitiy As Long) As Long

Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, _
    ByVal hObject As Long) As Long


Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, _
    ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, _
    ByVal YSrc As Long, ByVal dwRop As Long) As Long


Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long





Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, _
    ByVal hPalette As Long, ByVal bForceBackground As Long) As Long


Private Declare Function RealizePalette Lib "GDI32" ( _
    ByVal hDC As Long) As Long


Private Declare Function GetWindowDC Lib "user32" ( _
    ByVal hwnd As Long) As Long


Private Declare Function GetDC Lib "user32" ( _
    ByVal hwnd As Long) As Long


Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long


Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
    ByVal hDC As Long) As Long


Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type


Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
    PicDesc As PicBmp, RefIID As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture
    Dim r As Long


Dim Pic As PicBmp
' IPicture requires a reference to "Stan
'     dard OLE Types"
Dim IPic As IPicture
Dim IID_IDispatch As GUID
' Fill in with IDispatch Interface ID


With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With

With Pic
    .Size = Len(Pic) ' Length of structure
    .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
    .hBmp = hBmp ' Handle To bitmap
    .hPal = hPal ' Handle To palette (may be null)
End With
' Create Picture object
r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
' Return the new Picture object
Set CreateBitmapPicture = IPic
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal Client As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture
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

' Depending on the value of Client get t
'     he proper device context


If Client Then
    hDCSrc = GetDC(hWndSrc) ' Get device context For client area
Else
    hDCSrc = GetWindowDC(hWndSrc) ' Get device context For entire window
End If
' Create a memory device context for the
'     copy process
hDCMemory = CreateCompatibleDC(hDCSrc)
' Create a bitmap and place it in the me
'     mory DC
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
' Get screen properties
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette
' If the screen has a palette make a cop
'     y and realize it


    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)

' Copy the on-screen image into the memo
'     ry DC
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
LeftSrc, TopSrc, vbSrcCopy)
' Remove the new copy of the the on-scre
'     en image
hBmp = SelectObject(hDCMemory, hBmpPrev)
' If the screen has a palette get back t
'     he palette that was selected
' in previously


If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If
' Release the device context resources b
'     ack to the system
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
' Call CreateBitmapPicture to create a p
'     icture object from the bitmap
' and palette handles. Then return the r
'     esulting picture object.
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function


Private Sub form_load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 400, 0, 0, 0, SWP_SHOWWINDOW
Me.Width = 1785
Me.Height = 405
End Sub
Private Sub Timer2_Timer()

Me.Left = 5850
Me.Height = 1
Me.Width = 1
Set Picture1.Picture = CaptureWindow(hWndScreen, False, 0, 0, _
    Screen.Width \ Screen.TwipsPerPixelX, _
    Screen.Height \ Screen.TwipsPerPixelY)
frmSrc.Top = -35
Me.Width = 1785
Me.Height = 405
End Sub
Private Sub Timer1_Timer()
Me.Width = 1785
Me.Height = 405
Me.Left = 17000
Me.Top = 13000
End Sub
