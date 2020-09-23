Attribute VB_Name = "ReSizer"
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This file came from:                                         ****
'   ****                                                                 ****
'   ****    Title:       Load and Resize Pictures with GDI+              ****
'   ****                                                                 ****
'   ****    By:          Miltiadis Kritikos                              ****
'   ****                                                                 ****
'   ****    Description: This module uses a combination of GDI+ and      ****
'   ****    GDI API calls.  Its functionallity is limited to loading     ****
'   ****    images and resizing them.  It can load all the picture       ****
'   ****    formats supported by LoadPicture plus the PNG file           ****
'   ****    format.                                                      ****
'   ****                                                                 ****
'   ****    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=48352&lngWId=1
'   ****                                                                 ****
'   ****    See the bottom of this Module for an example of "How to      ****
'   ****    Use IT!".                                                    ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit

Private Type GUID
    Data1                                           As Long
    Data2                                           As Integer
    Data3                                           As Integer
    Data4(7)                                        As Byte
End Type

Private Type PICTDESC
    size                                            As Long
    Type                                            As Long
    hBmp                                            As Long
    hPal                                            As Long
    Reserved                                        As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion                                  As Long
    DebugEventCallback                              As Long
    SuppressBackgroundThread                        As Long
    SuppressExternalCodecs                          As Long
End Type

Private Type PWMFRect16
    Left                                            As Long
    Top                                             As Long
    Right                                           As Long
    Bottom                                          As Long
End Type

Private Type wmfPlaceableFileHeader
    Key                                             As Long
    hMf                                             As Integer
    BoundingBox                                     As PWMFRect16
    Inch                                            As Integer
    Reserved                                        As Long
    CheckSum                                        As Integer
End Type

'GDI and GDI+ constants
Private Const mc_strModuleID                        As String = "ReSizer."
Private Const PLANES                                As Long = 14          'Number of planes
Private Const BITSPIXEL                             As Long = 12          'Number of bits per pixel
Private Const PATCOPY                               As Long = &HF00021    '(DWORD) dest = pattern
Private Const PICTYPE_BITMAP                        As Long = 1           'Bitmap type
Private Const InterpolationModeHighQualityBicubic   As Long = 7
Private Const GDIP_WMF_PLACEABLEKEY                 As Long = &H9AC6CDD7
Private Const UnitPixel                             As Long = 2

'GDI Functions
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'GDI+ functions
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

'Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(ByVal hBitmap As Long) As IPicture

Dim IID_IDispatch                                   As GUID
Dim Pic                                             As PICTDESC
Dim IPic                                            As IPicture

'   Fill in OLE IDispatch Interface ID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With  'IID_IDispatch

'   Fill Pic with necessary parts
    With Pic
        .size = Len(Pic)     'Length of structure
        .Type = PICTYPE_BITMAP              'Type of Picture (bitmap)
        .hBmp = hBitmap      'Handle to bitmap
    End With  'Pic

'   Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic

End Function

Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, ByVal lngWidth As Long, ByVal lngHeight As Long)

    On Error GoTo ErrorHandler

'   Fills in the wmfPlacable header
    With WmfHeader
        .BoundingBox.Right = lngWidth
        .BoundingBox.Bottom = lngHeight
        .Inch = 1440
        .Key = GDIP_WMF_PLACEABLEKEY
    End With  'WmfHeader

ErrorHandler:

End Sub

Public Sub FreeGDIPlus(Token As Long)

    On Error GoTo ErrorHandler
'   Frees GDI Plus
    GdiplusShutdown Token
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "FreeGDIPlus(" & Token & ")"

End Sub

Private Sub gdipResize(Img As Long, lngHDC As Long, lngWidth As Long, lngHeight As Long, Optional ByVal RetainRatio As Boolean = False)

'Resize the picture using GDI plus
Dim DesRatio                                        As Double             'Destination rect Ratio
Dim DestHeight                                      As Long               'Destination image Height
Dim DestWidth                                       As Long               'Destination image Width
Dim Graphics                                        As Long               'Graphics Object Pointer
Dim OrHeight                                        As Long               'Original Image Height
Dim OrRatio                                         As Double             'Original Image Ratio
Dim OrWidth                                         As Long               'Original Image Width

    On Error GoTo ErrorHandler
    GdipCreateFromHDC lngHDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic

    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight

        OrRatio = OrWidth / OrHeight
        DesRatio = lngWidth / lngHeight

'       Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, lngWidth, lngHeight * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, lngWidth / OrRatio, lngHeight)
        GdipDrawImageRectRectI Graphics, Img, (lngWidth - DestWidth) / 2, (lngHeight - DestHeight) / 2, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, lngWidth, lngHeight
    End If
    GdipDeleteGraphics Graphics
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "gdipResize(" & Img & ", " & lngHDC & ", " & lngWidth & ", " & lngHeight & ", " & RetainRatio & ")"

End Sub

Private Sub GetBitmap(lngHDC As Long, hBitmap As Long)

    On Error GoTo ErrorHandler
'   Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
    hBitmap = SelectObject(lngHDC, hBitmap)
    DeleteDC lngHDC
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "GetBitmap(" & lngHDC & ", " & hBitmap & ")"

End Sub

Private Sub InitDC(lngHDC As Long, hBitmap As Long, lngBackColor As Long, lngWidth As Long, lngHeight As Long)

'Initialises the hDC to draw
Dim hBrush                                          As Long

    On Error GoTo ErrorHandler
'   Create a memory DC and select a bitmap into it, fill it in with the backcolor
    lngHDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(lngWidth, lngHeight, GetDeviceCaps(lngHDC, PLANES), GetDeviceCaps(lngHDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(lngHDC, hBitmap)
    hBrush = CreateSolidBrush(lngBackColor)
    hBrush = SelectObject(lngHDC, hBrush)
    PatBlt lngHDC, 0, 0, lngWidth, lngHeight, PATCOPY
    DeleteObject SelectObject(lngHDC, hBrush)
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "InitDC(" & lngHDC & ", " & hBitmap & ", " & lngBackColor & ", " & lngWidth & ", " & lngHeight & ")"

End Sub

Public Function InitGDIPlus() As Long

'Initialises GDI Plus
Dim Token                                           As Long
Dim gdipInit                                        As GdiplusStartupInput

    On Error GoTo ExitHandler
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token

ExitHandler:

End Function

Public Function ReSize(Handle As Long, PicType As PictureTypeConstants, lngWidth As Long, lngHeight As Long, Optional lngBackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture

'Returns a resized version of the picture
Dim Img                                             As Long
Dim hDC                                             As Long
Dim hBitmap                                         As Long
Dim WmfHeader                                       As wmfPlaceableFileHeader

'   Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
        GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
        FillInWmfHeader WmfHeader, lngWidth, lngHeight
        GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
        GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
'       Does not return a valid Image object
        GdipCreateBitmapFromHICON Handle, Img
    End Select

'   Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, lngBackColor, lngWidth, lngHeight
        gdipResize Img, hDC, lngWidth, lngHeight, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set ReSize = CreatePicture(hBitmap)
    End If

End Function
