Attribute VB_Name = "Puzzle_Bas"

Option Explicit

Private Const mc_strModuleID        As String = "Puzzle_Bas."
Private Const RASTERCAPS            As Long = 38
Private Const RC_PALETTE            As Long = &H100
Private Const SIZEPALETTE           As Long = 104

Public Const SPI_GETDESKWALLPAPER   As Long = 115

Public lngBoxes                     As Long
Public lngBoxesOOO                  As Long
Public lngH                         As Long
Public lngLoop1                     As Long
Public lngLoop2                     As Long
Public lngMsgResp                   As Long
Public lngPerRow                    As Long
Public lngPicTo                     As Long
Public lngSH                        As Long
Public lngSW                        As Long
Public lngSwap                      As Long
Public lngSwap1                     As Long
Public lngSwap2                     As Long
Public lngSwap3                     As Long
Public lngW                         As Long

Public strMsg                       As String
Public strOldPaper                  As String

Public Type PALETTEENTRY
    peRed                           As Byte
    peGreen                         As Byte
    peBlue                          As Byte
    peFlags                         As Byte
End Type

Public Type LOGPALETTE
    palVersion                      As Integer
    palNumEntries                   As Integer
    palPalEntry(255)                As PALETTEENTRY        'Enough for 256 colors
End Type

Public Type GUID
    Data1                           As Long
    Data2                           As Integer
    Data3                           As Integer
    Data4(7)                        As Byte
End Type

Public Type PicBmp
    size                            As Long
    PicType                         As Long
    hBmp                            As Long
    hPal                            As Long
    Reserved                        As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture

Dim IID_IDispatch                   As GUID
Dim IPic                            As IPicture
Dim Pic                             As PicBmp

    On Error GoTo ExitHandler

'   Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

'   Fill picture info
    With Pic
        .size = Len(Pic)     'Length of structure
        .PicType = vbPicTypeBitmap          'Type of Picture (bitmap)
        .hBmp = hBmp         'Handle to bitmap
        .hPal = hPal         'Handle to palette (may be null)
    End With

'   Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
'   Return the new picture
    Set CreateBitmapPicture = IPic

ExitHandler:

End Function

Public Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

Dim HasPaletteScrn                  As Long
Dim hBmp                            As Long
Dim hBmpPrev                        As Long
Dim hDCMemory                       As Long
Dim hPal                            As Long
Dim hPalPrev                        As Long
Dim LogPal                          As LOGPALETTE
Dim PaletteSizeScrn                 As Long

    On Error GoTo ExitHandler

'   Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
'   Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
'   Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

'   Does our picture use a palette?
    HasPaletteScrn = (GetDeviceCaps(hDCSrc, RASTERCAPS)) And RC_PALETTE   'Palette
'   What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)   'Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
'       Set the palette version
        With LogPal
            .palVersion = &H300
'           Number of palette entries
            .palNumEntries = 256
'           Retrieve the system palette entries
            GetSystemPaletteEntries hDCSrc, 0, 256, .palPalEntry(0)
'       Create the palette
        End With             'LogPal

        hPal = CreatePalette(LogPal)
'       Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
'       Realize the palette
        RealizePalette hDCMemory
    End If

'   Copy the source image to our compatible device context
    BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy
'   Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
'       Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

'   Delete our memory DC
    DeleteDC hDCMemory
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)

ExitHandler:

End Function
