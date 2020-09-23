Attribute VB_Name = "m2DRender"
'*******************************************************************************
'm2DRender.bas  by  Mario DiCristofano                      Email: MNSX7@aol.com
'*******************************************************************************
' all images should have a white background so image masks will work properly. *
'      this excludes images rendered with the render_maskedimage function      *
'*******************************************************************************
'This module contains several function for rendering graphics, storing them in
'memory, and keeping track of the images stored in memory, as well as getting
'and changing the current display's settings.  It can be used as a 2D rendering
'engine, whether for games or just to add animations to your windows-based
'programs.
'*******************************************************************************
'Module Structure Information:                           (all arrays start at 1)
'
'The module deals strongly with device contexts - this is where the images loaded
'into memory are stored.  With the particular way we use device contexts in this
'module, you can think of them as 'one big bitmap'.  By storing the dimensions
'of the images we load into the device contexts, we can do some quick math and
'determine where in the 'one big bitmap' each individual image is stored.  The
'dimensions of images stored in the ImageDC device context are kept in the
'ImageDimensions%() array.  The dimensions of images stored in the BackGroundDC
'device context are kept in the BackGroundImageDimensions%() array.  We also
'store the dimensions of the device contexts themselves:  ImageDC's dimensions
'are kept in the ImageDCDimensions%() array, while BackGroundDC's dimensions are
'stored in the BackGroundImageDCDimensions%() array.  The ImageDimensions%() and
'BackGroundImageDimensions%() arrays use a table format:
'
'   For instance, if you wanted to get the width AND height of the first
'   bitmap in the ImageDC device context, it would look like this:
'
'       gWidth% = ImageDimensions%(1, 1)
'       gHeight% = ImageDimensions%(1, 2)
'
'   The first number in the array specifies the Image's index in the ImageDC
'   device context.  The second number specifies either 1 for Width, or 2 for
'   Height.
'
'The BackGroundImageDimensions%() arrays works the same way.  The arrays that
'hold the dimensions of the device contexts themselves are different.  They only
'need to store two numbers, the Width and Height.  The first position in the
'array (1) holds the width, the second position in the array (2) holds the height.
'*******************************************************************************
'Each function contains a detailed explanation of it's parameters,
'return values, and examples, as well as the purpose of the function.
'*******************************************************************************
Option Explicit
Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
    'BitBlt Raster Operations:
    Public Const BLACKNESS = &H42 'Fill the destination rectangle with the color whose index is 0 in the physical palette (which is black by default).
    Public Const DSTINVERT = &H550009 'Invert the colors in the destination rectangle.
    Public Const MERGECOPY = &HC000CA 'Merge the colors of the source rectangle with the specified pattern using the bitwise AND operator.
    Public Const MERGEPAINT = &HBB0226 'Merge the colors of the inverted source rectangle with the colors of the destination rectangle using the bitwise OR operator.
    Public Const NOTSRCCOPY = &H330008 'Copy the inverted source rectangle to the destination rectangle.
    Public Const NOTSRCERASE = &H1100A6 'Combine the colors of the source and destination rectangles using the bitwise OR operator and then invert the resulting color.
    Public Const PATCOPY = &HF00021 'Copy the specified pattern into the destination bitmap.
    Public Const PATINVERT = &H5A0049 'Combine the colors of the specified pattern with the colors of the destination rectangle using the bitwise XOR operator.
    Public Const PATPAINT = &HFB0A09 'Combine the colors of the specified pattern with the colors of the inverted source rectangle using the bitwise OR operator. Combine the result of that operation with the colors of the destination rectangle using the bitwise OR operator.
    Public Const SRCAND = &H8800C6 'Combine the colors of the source and destination rectangles using the bitwise AND operator.
    Public Const SRCCOPY = &HCC0020 'Copy the source rectangle directly into the destination rectangle.
    Public Const SRCERASE = &H440328 'Combine the inverted colors of the destination rectangle with the colors of the source rectange using the bitwise AND operator.
    Public Const SRCINVERT = &H660046 'Combine the colors of the source and destination rectangles using the bitwise XOR operator.
    Public Const SRCPAINT = &HEE0086 'Combine the colors of the source and destination rectangles using the bitwise OR operator.
    Public Const WHITENESS = &HFF0062 'Fill the destination rectangle with the color whose index is 1 in the physical palette (which is white by default).
Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
    'CDS are flags used when changing display modes, DISP are function return values:
    Public Const CDS_UPDATEREGISTRY = &H1 'Save the new settings to the registry and also change the settings dynamically.
    Public Const CDS_TEST = &H2 'Test to see if the new settings are supported by the hardware, without actually changing the settings. The function's return value will indicate any problems that may have occured.
    Public Const CDS_FULLSCREEN = &H4 'Go into full-screen mode. This setting cannot be saved.
    Public Const CDS_GLOBAL = &H8 'Save the new settings for all users. The CDS_UPDATEREGISTRY flag must also be specified.
    Public Const CDS_SET_PRIMARY = &H10 'Make this device the primary display device.
    Public Const CDS_RESET = &H40000000 'Change the settings even if they are the same as the current settings.
    Public Const CDS_SETRECT = &H20000000
    Public Const CDS_NORESET = &H10000000 'Save the settings to the registry, but do not make them take effect yet. The CDS_UPDATEREGISTRY flag must also be specified.
    Public Const DISP_CHANGE_SUCCESSFUL = 0 'The display settings were successfully changed.
    Public Const DISP_CHANGE_RESTART = 1 'The computer must be restarted for the display changes to take effect.
    Public Const DISP_CHANGE_FAILED = -1 'The display driver failed the specified graphics mode.
    Public Const DISP_CHANGE_BADMODE = -2 'The specified graphics mode is not supported.
    Public Const DISP_CHANGE_NOTUPDATED = -3 'Windows NT/2000: The settings could not be written to the registry.
    Public Const DISP_CHANGE_BADFLAGS = -4 'An invalid set of flags was specified.
    Public Const DISP_CHANGE_BADPARAM = -5 'An invalid parameter was specified.
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpszDriver As String, ByVal lpszDevice As String, ByVal lpszOutput As Long, lpInitData As Any) As Long
Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long
    'Determine which data to get for iModeNum:
    Public Const ENUM_CURRENT_SETTINGS = -1 'Retrieves current display settings
    Public Const ENUM_REGISTRY_SETTINGS = -2 'Retrieves setting stored in the registry
Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Declare Function GetTickCount Lib "kernel32.dll" () As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal iImageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fFlags As Long) As Long
    'Constants for iImageType:
    Public Const IMAGE_BITMAP = 0 'Specifies image is a bitmap.
    Public Const IMAGE_ICON = 1 'Specifies image is an icon.
    'Constants for fFlags:
    Public Const IM_DEFAULTCOLOR = &H0 'Load with default colors.
    Public Const IM_LOADFROMFILE = &H10 'Load from file
    Public Const IM_DEFAULTSIZE = &H40 'Load with normal size.
Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)
Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    'Supported raster operations:
    'SRCAND constant declared under BITBLT declaration.
    'SRCCOPY constant declared under BITBLT declaration.
    'SRCERASE constant declared under BITBLT declaration.
    'SRCINVERT constant declared under BITBLT declaration.
    'SRCPAINT constant declared under BITBLT declaration.
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPixel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

'm2DRender! specific:
Public Enum m2DR_Type
    L_BACKGROUND = 0
    L_IMAGE = 1
End Enum
Public Enum m2DR_StretchType
    R_IMAGE = 0
    R_IMAGE_MASK = 1
    R_IMAGE_CREATE_MASK = 2
End Enum
Public ImageDCDimensions(1 To 2) As Integer
Public BackGroundDCDimensions(1 To 2) As Integer
'------------------------------------------------------------------------------
Public DisplayModes() As String 'Holds list of supported display modes
Public BackGroundDC As Long 'Holds handle to device context that contains backgrounds
Public BackGroundImageDimensions() As Integer 'Holds background image dimensions of images loaded into the BackGroundDC device context
Public ImageDC As Long 'Holds handle to device context that contains images
Public MaskDC As Long 'Holds handle to device context that contains masks for ImageDC
Public ImageDimensions() As Integer 'Holds dimensions of images loaded in device context ImageDC.
'------------------------------------------------------------------------------

Public Function Change_DisplayMode(ForceDisplayChange As Boolean, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nColorDepth As Integer, ByVal nRefreshRate As Long, Optional GSDM_String As String = "") As Boolean
'*******************************************************************************
'This function changes the current display mode's resolution (width and height)
'color depth (16 or 32 bit) and the monitor's refresh rate.
'
'Parameters:
'ForceDisplayChange = If you set this to true, it will change the display
    'settings even if the settings are the same as the current display.  False
    'will not change the display if the settings are the same.
'nWidth = New screen width
'nHeight = New screen height
'nColorDepth = Bits Per Pixel.  16 for 16-bit color, 32 for 32-bit color
'nRefreshRate = Specify the monitor's refresh rate
'GSDM_String = If you have previously made a call to Get_SupportedDisplayModes,
'   then you can feed this function the string from array DisplayModes$() that
'   contains the settings you want to switch to.  It will then parse the string
'   and change the settings.  The nWidth, nHeight, nColorDepth & nRefreshRate
'   parameters are ignored if a string is specified here.  ForceDisplayChange
'   still gets checked.  If this parameter is left blank, it defaults to the
'   settings specified using the parameters explained above.
'
'Returns:
'True = Display change successful
'False = Bad parameter, unsupported display mode, error writing display settings
'   to the registry, or if a string was given for GSDM_String, there might have
'   been an error parsing the string.
'
'Examples:
'
'1) To change the display to user specified settings:
'
'       bReturn = Change_DisplayMode(False, 640, 480, 32, 70)
'
'2) To change the display mode to that of a string returned by the function
'   Get_SupportedDisplayModes().  Where DisplayModes$(1) is the first entry
'   in the array compiled by Get_SupportedDisplayModes().
'
'       bReturn = Change_DisplayMode(False, 0, 0, 0, 0, DisplayModes$(1))
'
'*******************************************************************************
Dim iFindRefresh1 As Integer, iFindRefresh2 As Integer
Dim iWidth As Integer, iHeight As Integer, iColor As Integer
Dim lProcess As Long, CurrentDisp As DEVMODE, iRefresh As Integer
Dim iFindWidth As Integer, iFindHeight As Integer, iFindColor As Integer

'Get the current display mode
lProcess& = EnumDisplaySettings(vbNullString, ENUM_CURRENT_SETTINGS, CurrentDisp)
If lProcess& = 0 Then GoTo ERROR_HANDLER

If GSDM_String$ = "" Then 'No string specified, use given values
    'Enter new settings into DEVMODE structure
    CurrentDisp.dmPelsWidth = nWidth&
    CurrentDisp.dmPelsHeight = nHeight&
    CurrentDisp.dmBitsPerPixel = nColorDepth%
    CurrentDisp.dmDisplayFrequency = nRefreshRate&
Else
    'Process the string returned by Get_SupportedDisplayModes function
    'and get values
    iFindWidth% = InStr(1, GSDM_String$, "x")
    iFindHeight% = InStr((iFindWidth% + 1), GSDM_String$, "x")
    If iFindWidth% = 0 Or iFindHeight% = 0 Then GoTo ERROR_HANDLER
    iWidth% = Left(GSDM_String$, (iFindWidth% - 1))
    iHeight% = Mid(GSDM_String$, (iFindWidth% + 1), iFindHeight% - (iFindWidth% + 1))
    
    iFindColor% = InStr((iFindHeight% + 1), GSDM_String$, " ")
    If iFindColor% = 0 Then GoTo ERROR_HANDLER
    iColor% = Mid(GSDM_String$, (iFindHeight% + 1), iFindColor% - (iFindHeight% + 1))
    
    iFindRefresh1% = InStr((iFindColor% + 1), GSDM_String$, " ")
    iFindRefresh2% = InStr((iFindRefresh1% + 1), GSDM_String$, " ")
    If iFindRefresh1% = 0 Or iFindRefresh2% = 0 Then GoTo ERROR_HANDLER
    iRefresh% = Mid(GSDM_String$, (iFindRefresh1% + 1), iFindRefresh2% - (iFindRefresh1% + 1))
    
    'Set values we pulled from the string into the DEVMODE structure
    CurrentDisp.dmPelsWidth = iWidth%
    CurrentDisp.dmPelsHeight = iHeight%
    CurrentDisp.dmBitsPerPixel = iColor%
    CurrentDisp.dmDisplayFrequency = iRefresh%
End If

'Test the new settings without changing them
lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
If lProcess& = 0 Then 'Test passed
    If ForceDisplayChange = True Then 'Force display change
        lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_UPDATEREGISTRY + CDS_SET_PRIMARY + CDS_RESET)
        If lProcess& <> 0 Then GoTo ERROR_HANDLER
    ElseIf ForceDisplayChange = False Then 'Don't force change
        lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_UPDATEREGISTRY + CDS_SET_PRIMARY)
        If lProcess& <> 0 Then GoTo ERROR_HANDLER
    End If
Else
    GoTo ERROR_HANDLER
End If

Change_DisplayMode = True
Exit Function

ERROR_HANDLER:
Change_DisplayMode = False
End Function
Public Function Get_ImageXPosFromDC(ByVal ImageIndex As Integer, DeviceContext As m2DR_Type) As Long
'*******************************************************************************
'This function calculates the X coordinate of an image based on it's index
'in the ImageDC device context.  It does so by looping through all dimensions
'in the ImageDimensions() array until it reaches the specified index, and
'adding their widths together.
'
'Parameters:
'ImageIndex = Index of image to find
'DeviceContext = Specify which device context the image is in
'
'Returns:
'The X coordinate of the specified image as a long value.
'
'Example:
'
'1) To retrieve the X coordinate of the 3rd image loaded into the ImageDC
'   device context
'
'       TheXCoordinate& = Get_ImageXPosFromDC(3, L_IMAGE)
'
'*******************************************************************************
Dim lPlaceKeeper As Long, iImageWidth As Long

'If the image is the first in the index, the coordinate = 0.  Exit.
If ImageIndex% = 1 Then
    Get_ImageXPosFromDC& = 0
    Exit Function
End If

If DeviceContext = L_IMAGE Then 'Find in ImageDC device context.
    'Loop through all dimensions of array and add up total width.
    For lPlaceKeeper& = 1 To (ImageIndex% - 1)
        iImageWidth& = (iImageWidth& + ImageDimensions%(lPlaceKeeper&, 1))
    Next lPlaceKeeper&
ElseIf DeviceContext = L_BACKGROUND Then 'Find in BackGroundDC device context.
    For lPlaceKeeper& = 1 To (ImageIndex% - 1)
        iImageWidth& = (iImageWidth& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
    Next lPlaceKeeper&
End If

'Total width is where the X position is in the device context.
Get_ImageXPosFromDC& = iImageWidth&
End Function
Public Function Get_SupportedDisplayModes(Optional Optimal_RefreshRate As Boolean = False) As Boolean
'*******************************************************************************
'This function returns a list of supported display modes.  The list is stored
'in a public array called DisplayModes$().  You will have to access this array
'to get at the data this function compiled.
'
'------------------------------------------------------------------------------
'Tests the following resolutions in 16 and 32 bit color:
'   320x240,640x480,800x600,1024x768,1280x1024,1600x1280,1920x1440
'Tests for the following refresh rates:
'   60 Hz, 70 Hz, 72 Hz, 75 Hz, 85 Hz, 90 Hz, 95 Hz
'------------------------------------------------------------------------------
'
'Parameters:
'Optimal_RefreshRate = If set to true, the function will return only the
'   fastest refresh rate for the given resolution and color depth.  If left
'   blank or set to False, the function returns all screen resolutions, color
'   depthes and all refresh rates.
'
'Returns:
'True = List of display modes stored in DisplayModes$() public array
'False = Error enumerating current display modes, function failed
'
'Example:
'
'1) To fill the public array DisplayModes$() with all the supported settings
'
'       bReturn = Get_SupportedDisplayModes()
'
'You will need to access the DisplayModes$() array to retrieve the data returned
'by this function.  A For/Next adding each index in the array to a list works,
'if you want the user to be able to view the settings.  You can use this in
'conjuction with the Change_DisplaySettings function, feeding it a string
'returned by this function, chosen from a listbox to let the user select a new
'display mode.  Remember, ListBox and ComboBox indexes start at 0, the
'DisplayModes$() array's index starts at 1!
'*******************************************************************************
Dim lProcess As Long, lPlaceKeeper As Long
Dim TempArray(1 To 98) As String, CurrentDisp As DEVMODE

lPlaceKeeper& = 1
'Fill the DEVMODE structure with the current display mode's settings
lProcess& = EnumDisplaySettings(vbNullString, ENUM_CURRENT_SETTINGS, CurrentDisp)
If lProcess& = 0 Then
    Get_SupportedDisplayModes = False
    Exit Function
End If

'Test each individual resolution, color depth, and refresh rate and record
'successful settings to the TempArray$().  lPlaceKeeper& keeps track of the
'current array position, so we don't overwrite any other values.

If Optimal_RefreshRate = False Then 'Get all supported display modes
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "320x240x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "640x480x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "800x600x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1024x768x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 60 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 70 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 72 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 75 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 85 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 90 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 95 Hz"
    lPlaceKeeper& = (lPlaceKeeper& + 1)
    '------------------------------------------------------------------------------
ElseIf Optimal_RefreshRate = True Then 'Go through the refresh rates backwards, and record the fastest setting only.
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G320_240_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G640_480_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G800_600_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1024_768_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1280_1024_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1600_1280_16D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_16D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 16: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x16 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1900_1440_16D:
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G320_240_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 320: CurrentDisp.dmPelsHeight = 240
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "320x240x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G320_240_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G640_480_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 640: CurrentDisp.dmPelsHeight = 480
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "640x480x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G640_480_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G800_600_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 800: CurrentDisp.dmPelsHeight = 600
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "800x600x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G800_600_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1024_768_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1024: CurrentDisp.dmPelsHeight = 768
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1024x768x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1024_768_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1280_1024_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1280: CurrentDisp.dmPelsHeight = 1024
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1280x1024x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1280_1024_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1600_1280_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1600: CurrentDisp.dmPelsHeight = 1280
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1600x1280x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1600_1280_32D:
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 95
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 95 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 90
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 90 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 85
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 85 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 75
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 75 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 72
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 72 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 70
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 70 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
        GoTo G1900_1440_32D
    End If
    '------------------------------------------------------------------------------
    CurrentDisp.dmPelsWidth = 1920: CurrentDisp.dmPelsHeight = 1440
    CurrentDisp.dmBitsPerPixel = 32: CurrentDisp.dmDisplayFrequency = 60
    lProcess& = ChangeDisplaySettings(CurrentDisp, CDS_TEST)
    If lProcess& = 0 Then
        TempArray$(lPlaceKeeper&) = "1920x1440x32 @ 60 Hz"
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
G1900_1440_32D:
End If

'Count number of valid display modes returned
lPlaceKeeper& = 0
For lProcess& = 1 To 98
    If TempArray$(lProcess&) <> "" Then
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
Next lProcess&

'Redimension public array that holds display information to number of valid
'display modes returned
ReDim DisplayModes$(1 To lPlaceKeeper&)

'Transfer valid display modes from TempArray() to properly sized DisplayModes()
'main array
lPlaceKeeper& = 1
For lProcess& = 1 To 98
    If TempArray$(lProcess&) <> "" Then
        DisplayModes$(lPlaceKeeper&) = TempArray$(lProcess&)
        lPlaceKeeper& = (lPlaceKeeper& + 1)
    End If
Next lProcess&

Get_SupportedDisplayModes = True 'Success
End Function
Public Function Initialize_m2DRender(ImageDCWidth As Long, ImageDCHeight As Long, Optional Create_ImageMaskDC As Boolean = False, Optional BackGroundDCWidth As Long = 0&, Optional BackGroundDCHeight As Long = 0&) As Long
'*******************************************************************************
'This function initializes the module, creating the ImageDC device context and
'the BackGroundDC device context.  You must specifiy the initial starting sizes
'This function should be called before rendering or loading any images for the
'first time.
'
'Parameters:
'ImageDCWidth = The initial width of the ImageDC device context
'ImageDCHeight = The initial height of the ImageDC device context
'Create_ImageMaskDC = Specify True to create the MaskDC device context or
'   False not to create the MaskDC device context.  If True, the MaskDC's
'   dimensions will always be the same as those of ImageDC's dimensions.
'BackGroundDCWidth = The initial width of the BackGroundDC device context
'   This is optional, and if not given or a 0 value is specified, the device
'   context is not created
'BackGroundDCHeight = The initial height of the BackGroundDC device context
'   This is optional, and if not given or a 0 value is specified, the device
'   context is not created
'
'
'Returns:
'SUCCESS = The time taken in ticks to initialize the module
'FAILURE = -1 long value
'
'Example:
'
'1) Initialize the module, and have it create an ImageDC device context that is
'   100x100 pixels, have it create a device context to store masks of any images
'   loaded into ImageDC device context, and a BackGroundDC device context that is
'   250x100 pixels.
'
'       lTicksElapsed& = Initialize_m2DRender(100, 100, True, 250, 100)
'
'*******************************************************************************
Dim MaskBMP As Long, TempMaskDC As Long
Dim ImageBMP As Long, CurrentDisp As Long, MergeObject As Long
Dim BackGroundBMP As Long, BackGroundMerge As Long, lStartTime As Long

'Redimension main image dimension arrays, so they have a UBound() of 0
ReDim ImageDimensions%(0, 0)
ReDim BackGroundImageDimensions%(0, 0)

lStartTime& = GetTickCount()
'Create a device context equal to the current display
CurrentDisp& = CreateDC("DISPLAY", 0&, 0&, 0&)
'Create the main device context used to store images
ImageDC& = CreateCompatibleDC(GetDC(CurrentDisp&))
'Create a bitmap compatible with the current display settings
ImageBMP& = CreateCompatibleBitmap(CurrentDisp&, ImageDCWidth&, ImageDCHeight&)
'Merge the compatible bitmap into the main device context
MergeObject& = SelectObject(ImageDC&, ImageBMP&)
If ImageBMP& = 0 Or MergeObject& = 0 Then GoTo ERROR_HANDLER

If ImageDC& <> 0 Then
    ImageDCDimensions%(1) = ImageDCWidth&
    ImageDCDimensions%(2) = ImageDCHeight&
    ReDim ImageDimensions%(0, 1 To 2)

    'Free resources
    Call DeleteObject(ImageBMP&)
    Call DeleteObject(MergeObject&)
End If

'If specified, create the device context to hold image masks
If Create_ImageMaskDC = True Then
    'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
    'bitmap and merge it into the TempMaskDC
    TempMaskDC& = CreateCompatibleDC&(ImageDC&)
    ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
    MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
    
    'Free resources
    Call DeleteObject(ImageBMP&)
    Call DeleteObject(MergeObject&)
    
    'Create a device context compatible with the temporary device context,
    'TempMaskDC.  This means it will be monochrome (black and white).  Create
    'a properly sized bitmap equal to the specified ImageDC device context's
    'width and height.  Merge that bitmap into the MaskDC device context.
    'This created a device context equal in size to ImageDC, but in black and
    'white.
    MaskDC& = CreateCompatibleDC(TempMaskDC&)
    ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCWidth&, ImageDCHeight&)
    MergeObject = SelectObject(MaskDC&, ImageBMP&)
    
    'Free resources: delete temporary DC
    Call DeleteDC(TempMaskDC&)
    Call DeleteObject(ImageBMP&)
    Call DeleteObject(MergeObject&)
Else 'Not creating the MaskDC device context
    MaskDC& = 0
End If
'Create background device context also?
If BackGroundDCWidth& <> 0& And BackGroundDCHeight& <> 0& Then
    BackGroundDC& = CreateCompatibleDC(CurrentDisp&)
    BackGroundBMP& = CreateCompatibleBitmap(CurrentDisp&, BackGroundDCWidth&, BackGroundDCHeight&)
    BackGroundMerge& = SelectObject(BackGroundDC&, BackGroundBMP&)
    If BackGroundBMP& = 0 Or BackGroundMerge& = 0 Then GoTo ERROR_HANDLER
        
    If BackGroundDC& <> 0 Then
        BackGroundDCDimensions%(1) = BackGroundDCWidth&
        BackGroundDCDimensions%(2) = BackGroundDCHeight&
        ReDim BackGroundImageDimensions%(0, 1 To 2)
            
        Call DeleteObject(BackGroundBMP&)
        Call DeleteObject(BackGroundMerge&)
    End If
End If

Call DeleteDC(CurrentDisp&)
Initialize_m2DRender& = (GetTickCount() - lStartTime&) 'Success!
Exit Function
    
ERROR_HANDLER:
'Free up resources for failed processes.
Call DeleteDC(ImageDC&)
Call DeleteDC(CurrentDisp&)
Call DeleteDC(BackGroundDC&)
Call DeleteDC(MaskDC&)
Call DeleteObject(ImageBMP&)
Call DeleteObject(BackGroundBMP&)
Call DeleteObject(MergeObject&)
Call DeleteObject(BackGroundMerge&)
Initialize_m2DRender& = -1
End Function
Public Function Load_ImageIntoMemory(ImageFile As String, WhatToLoad As m2DR_Type, Optional CreateImageMask As Boolean = False) As Long
'*******************************************************************************
'This function loads an image into the specified device context.  It first
'makes sure that the specified device context exists, and if it doesn't exist,
'it creates the device context.  If the device context already exists, but
'it is not big enough to hold the old images and the new image that is being
'loaded, it resizes the device context to hold all the images.  The dimensions
'of the image being loaded are also stored in either the ImageDimensions() or
'the BackGroundImageDimensions() arrays, depending on which device context you
'are loading an image to.
'
'Parameters:
'ImageFile = the path to the image to load.
'WhatToLoad = Choose either L_BACKGROUND to load a image to the BackGroundDC
'   device context or L_IMAGE to load an image to the ImageDC device context.
'CreateImageMask = If set to True, a mask will be created for the image being
'   loaded and stored in the device context MaskDC.  If no MaskDC device
'   context exists, one will be created.  This parameter is ignored if the
'   device context (WhatToLoad=L_BACKGROUND) is the BackGroundDC device context.
'
'Returns:
'Successful = returns tick count taken to complete operation.
'Failure = -1 long value
'
'Example:
'
'1) Load a bitmap into the ImageDC device context, and create a mask for the
'   image.  Putting the image filename into a string is not necessary.
'
'       IMG$ = App.Path & "\TestBMP.bmp"
'       lTicksElapsed& = Load_ImageIntoMemory(IMG$, L_IMAGE, True)
'
'*******************************************************************************
Dim lBitBlt As Long, TempMaskDC As Long
Dim lMaxHeight As Long, TempDC2 As Long, ImageBMP As Long
Dim ImageDCXPosition As Long, CurrentDisp As Long
Dim TempBMP As Long, TempMerge As Long
Dim BMPInfo As BITMAP, lStartTime As Long, lCount As Long
Dim TempDimensions() As Integer, MergeObject As Long
Dim TempDC As Long, TempImage As Long, lPlaceKeeper As Long

'Record the tick count before process starts
lStartTime& = GetTickCount()

'Get the current display's settings
CurrentDisp& = CreateDC("DISPLAY", 0&, 0&, 0&)
'Create a temporary device context
TempDC& = CreateCompatibleDC(CurrentDisp&)
'Call the function to load the image into memory
TempImage& = LoadImage(0&, ImageFile$, IMAGE_BITMAP, 0&, 0&, IM_LOADFROMFILE + IM_DEFAULTSIZE + IM_DEFAULTCOLOR)
'Get the loaded image's dimensions, and save them in the BITMAP structure.
Call GetObject(TempImage&, Len(BMPInfo), BMPInfo)
'Merge the loaded image into the device context
MergeObject& = SelectObject(TempDC&, TempImage&)
If TempDC& = 0 Or TempImage& = 0 Or MergeObject& = 0 Then
    GoTo ERROR_HANDLER
End If

'If we're loading an image to the ImageDC device context then:
If WhatToLoad = L_IMAGE Then
    'See if ImageDC device context exists.  If not, create it.
    If ImageDC& = 0 Then 'create compatible device context and bitmap.  Merge.
        ImageDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, BMPInfo.bmWidth, BMPInfo.bmHeight)
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Record new ImageDC device context dimensions
        ImageDCDimensions%(1) = BMPInfo.bmWidth
        ImageDCDimensions%(2) = BMPInfo.bmHeight
        
        'Free resources
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Check if we're creating a mask.  If we are, create the MaskDC device
        'context, since it can't exist if ImageDC doesn't exist, and we just
        'created ImageDC device context.
        If CreateImageMask = True Then
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create a device context compatible with the temporary device context,
            'TempMaskDC.  This means it will be monochrome (black and white).  Create
            'a properly sized bitmap equal to the specified ImageDC device context's
            'width and height.  Merge that bitmap into the MaskDC device context.
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
        End If
    End If
    
    'If no MaskDC device context exists, and we need to create an image mask,
    'then we create the MaskDC device context.
    If CreateImageMask = True And MaskDC& = 0 Then
        'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
        'bitmap and merge it into the TempMaskDC
        TempMaskDC& = CreateCompatibleDC&(ImageDC&)
        ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
        MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
        'Free resources
        Call DeleteObject(ImageBMP&)
        Call DeleteObject(MergeObject&)
            
        'Create a device context compatible with the temporary device context,
        'TempMaskDC.  This means it will be monochrome (black and white).  Create
        'a properly sized bitmap equal to the specified ImageDC device context's
        'width and height.  Merge that bitmap into the MaskDC device context.
        MaskDC& = CreateCompatibleDC(TempMaskDC&)
        ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
        MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
        'Free resources
        Call DeleteDC(TempMaskDC&)
        Call DeleteObject(ImageBMP&)
        Call DeleteObject(MergeObject&)
    End If
    
    'Add up total width currently in ImageDC device context and get tallest
    'height
    lCount& = 0
    lMaxHeight& = 0
    For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
        lCount& = (lCount& + ImageDimensions%(lPlaceKeeper&, 1))
        If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    'If new image is taller than tallest image already in device context
    'then set the tallest image to that of the new image to be loaded
    If BMPInfo.bmHeight > lMaxHeight& Then lMaxHeight& = BMPInfo.bmHeight
    
    'If new width or height exceeds that of ImageDC device context then
    'resize the device context to hold the new image
    If (lCount& + BMPInfo.bmWidth) > ImageDCDimensions%(1) Or lMaxHeight& > ImageDCDimensions%(2) Then
        'Create temporary device context and bitmap, resized to hold the new
        'information.  Merge the bitmap into the device context.
        TempDC2& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, (lCount& + BMPInfo.bmWidth), lMaxHeight&)
        TempMerge& = SelectObject(TempDC2&, TempBMP&)
        If TempDC2& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy images already stored in the device context to the temporary
        'device context.
        lBitBlt& = BitBlt(TempDC2&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        
        'Free resources
        Call DeleteDC(ImageDC&) 'Delete ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Create new ImageDC device context and newly sized bitmap compatible
        'with the TempDC2 device context, merge bitmap into ImageDC
        ImageDC& = CreateCompatibleDC(TempDC2&)
        TempBMP& = CreateCompatibleBitmap(TempDC2&, (lCount& + BMPInfo.bmWidth), lMaxHeight&)
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Record new ImageDC device context dimensions
        ImageDCDimensions%(1) = (lCount& + BMPInfo.bmWidth)
        ImageDCDimensions%(2) = lMaxHeight&
        
        'Copy contents of TempDC2& back over to the properly sized ImageDC
        'device context
        lBitBlt& = BitBlt(ImageDC&, 0, 0, lCount&, lMaxHeight&, TempDC2&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        
        'Free resources
        Call DeleteDC(TempDC2&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'If a MaskDC device context already exists, then we also need to
        'resize that device context, regardless of whether the CreateImageMask
        'parameter is true or false.  The ImageDC device context and MaskDC
        'device context must always be the same size.
        If MaskDC& <> 0 Then
            Call DeleteDC(MaskDC&) 'Delete the previous MaskDC device context
            
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            If TempMaskDC& = 0 Or ImageBMP& = 0 Or MergeObject& = 0 Then GoTo ERROR_HANDLER
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create a device context compatible with the temporary device context,
            'TempMaskDC.  This means it will be monochrome (black and white).  Create
            'a properly sized bitmap equal to the specified ImageDC device context's
            'width and height.  Merge that bitmap into the MaskDC device context.
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            If MaskDC& = 0 Or ImageBMP& = 0 Or MergeObject& = 0 Then GoTo ERROR_HANDLER
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Copy the contents of ImageDC to the black and white MaskDC device
            'context, turning all images into masks.
            lBitBlt& = BitBlt(MaskDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        End If
    End If
    
    '--------------------------------------------------------------------------
    'Since the public array (ImageDimensions) holds all the dimension info, we
    'need to transfer the old data to a temporary array, redimensions the main
    'array to a size that will hold all the old data + the new entry, and then
    'transfer the data from the temporary array back to the main array.
    ReDim TempDimensions%(1 To (UBound(ImageDimensions%()) + 1), 1 To 2)
    
    For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
        TempDimensions%(lPlaceKeeper&, 1) = ImageDimensions%(lPlaceKeeper&, 1)
        TempDimensions%(lPlaceKeeper&, 2) = ImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    TempDimensions%((UBound(TempDimensions%())), 1) = BMPInfo.bmWidth
    TempDimensions%((UBound(TempDimensions%())), 2) = BMPInfo.bmHeight
    
    ReDim ImageDimensions%(1 To UBound(TempDimensions%()), 1 To 2)
    
    For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
        ImageDimensions%(lPlaceKeeper&, 1) = TempDimensions%(lPlaceKeeper&, 1)
        ImageDimensions%(lPlaceKeeper&, 2) = TempDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    'At this point, the main array has been redimensioned to hold all the data
    '& the data has been transferred from the temporary array back to the main
    'array.  The newly loaded image dimensions have also been added to the main
    'array.
    '--------------------------------------------------------------------------
    
    'Blit the temporary device context's contents to the proper X coordinate
    'in the main ImageDC device context.
    lBitBlt& = BitBlt(ImageDC&, Get_ImageXPosFromDC(UBound(ImageDimensions%()), L_IMAGE), 0, BMPInfo.bmWidth, BMPInfo.bmHeight, TempDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    
    'If creating a mask also, then copy the image to the MaskDC device context
    'also
    If CreateImageMask = True Then
        lBitBlt& = BitBlt(MaskDC&, Get_ImageXPosFromDC(UBound(ImageDimensions%()), L_IMAGE), 0, BMPInfo.bmWidth, BMPInfo.bmHeight, TempDC&, 0, 0, SRCCOPY)
    End If
    
ElseIf WhatToLoad = L_BACKGROUND Then 'Loading a background to BackGroundImageDC device context
    
    'See if BackGroundDC device context exists.  If not, create it.
    If BackGroundDC& = 0 Then 'create compatible device context and bitmap.  Merge.
        BackGroundDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, BMPInfo.bmWidth, BMPInfo.bmHeight)
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Record new BackGroundDC device context dimensions
        BackGroundDCDimensions%(1) = BMPInfo.bmWidth
        BackGroundDCDimensions%(2) = BMPInfo.bmHeight
        
        'Free resources
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
    End If
    
    lCount& = 0
    lMaxHeight& = 0
    'Add up total width currently in BackGroundDC device context and get tallest
    'height
    For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
        lCount& = (lCount& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
        If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    'If new image is taller than tallest image already in device context
    'then set the tallest image to that of the new image to be loaded
    If BMPInfo.bmHeight > lMaxHeight& Then lMaxHeight& = BMPInfo.bmHeight
    
    'If new width or height exceeds that of BackGroundDC device context then
    'resize the device context to hold the new image
    If (lCount& + BMPInfo.bmWidth) > BackGroundDCDimensions%(1) Or lMaxHeight& > BackGroundDCDimensions%(2) Then
        'Create temporary device context and bitmap, resized to hold the new
        'information.  Merge the bitmap into the device context.
        TempDC2& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, (lCount& + BMPInfo.bmWidth), lMaxHeight&)
        TempMerge& = SelectObject(TempDC2&, TempBMP&)
        If TempDC2& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy images already stored in the device context to the temporary
        'device context.
        lBitBlt& = BitBlt(TempDC2&, 0, 0, lCount&, lMaxHeight&, BackGroundDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        
        'Free resources
        Call DeleteDC(BackGroundDC&) 'Delete BackGroundDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Create new BackGroundDC device context and newly sized bitmap compatible
        'with the TempDC2 device context, merge bitmap into BackGroundDC
        BackGroundDC& = CreateCompatibleDC(TempDC2&)
        TempBMP& = CreateCompatibleBitmap(TempDC2&, (lCount& + BMPInfo.bmWidth), lMaxHeight&)
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Record new BackGroundDC device context dimensions
        BackGroundDCDimensions%(1) = (lCount& + BMPInfo.bmWidth)
        BackGroundDCDimensions%(2) = lMaxHeight&
        
        'Copy contents of TempDC2& back over to the properly sized BackGroundDC
        'device context
        lBitBlt& = BitBlt(BackGroundDC&, 0, 0, lCount&, lMaxHeight&, TempDC2&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        
        'Free resources
        Call DeleteDC(TempDC2&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
    End If
    
    '-------------------------------------------------------------------------
    'Since the public array (BackGroundImageDimensions) holds all the dimension
    'info, we need to transfer the old data to a temporary array, redimensions
    'the main array to a size that will hold all the old data + the new entry,
    'and then transfer the data from the temporary array back to the main array
    ReDim TempDimensions%(1 To (UBound(BackGroundImageDimensions%()) + 1), 1 To 2)
    
    For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
        TempDimensions%(lPlaceKeeper&, 1) = BackGroundImageDimensions%(lPlaceKeeper&, 1)
        TempDimensions%(lPlaceKeeper&, 2) = BackGroundImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    TempDimensions%((UBound(TempDimensions%())), 1) = BMPInfo.bmWidth
    TempDimensions%((UBound(TempDimensions%())), 2) = BMPInfo.bmHeight
    
    ReDim BackGroundImageDimensions%(1 To UBound(TempDimensions%()), 1 To 2)
    
    For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
        BackGroundImageDimensions%(lPlaceKeeper&, 1) = TempDimensions%(lPlaceKeeper&, 1)
        BackGroundImageDimensions%(lPlaceKeeper&, 2) = TempDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    'At this point, the main array has been redimensioned to hold all the data
    '& the data has been transferred from the temporary array back to the main
    'array.  The newly loaded image dimensions have also been added to the main
    'array.
    '--------------------------------------------------------------------------
    
    'Blit the temporary device context's contents to the proper X coordinate
    'in the main BackGroundImageDC device context.
    lBitBlt& = BitBlt(BackGroundDC&, Get_ImageXPosFromDC(UBound(BackGroundImageDimensions%()), L_BACKGROUND), 0, BMPInfo.bmWidth, BMPInfo.bmHeight, TempDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
End If

'Sucess; Free resources and calculate ticks taken to process.
Call DeleteDC(CurrentDisp&)
Call DeleteDC(TempDC&)
Call DeleteDC(TempDC2&)
Call DeleteObject(TempImage&)
Call DeleteObject(MergeObject&)
Load_ImageIntoMemory = (GetTickCount() - lStartTime&)
Exit Function

ERROR_HANDLER:
'Free resources
Call DeleteDC(CurrentDisp&)
Call DeleteDC(TempDC&)
Call DeleteDC(TempDC2&)
Call DeleteObject(TempImage&)
Call DeleteObject(MergeObject&)
Load_ImageIntoMemory = -1
End Function
Public Function Remove_ImageFromMemory(ImageIndex As Integer, DeviceContext As m2DR_Type) As Boolean
'*******************************************************************************
'This function removes an Image from the device context.  It then calculates
'the total width of all remaining images, recreates a device context that is
'the exact size of the remaining images, and copies the remaining images to
'the new device context.
'
'Parameters:
'ImageIndex = Index of image to remove
'DeviceContext = Specifies whether removing an image from ImageDC or BackGroundDC
'   device contexts.  If the DeviceContext is L_IMAGE (ImageDC device context)
'   then the mask from MaskDC will be removed and the MaskDC device context
'   itself will be automatically resized, and ALL images will have masks created
'   for them, even if you specified NOT to CreateImageMasks when the call was
'   made to Load_ImageIntoMemory.
'
'
'Returns:
'True = if successful
'False = if failed
'
'Example:
'
'1) Remove the image that has an index of 2 (the second image loaded into the
'   device context) from the BackGroundDC.
'
'   bReturn = Remove_ImageFromMemory(2, L_BACKGROUND)
'
'*******************************************************************************
Dim PlaceHolderDC As Long, TempArray() As Integer
Dim lMaxHeight As Long, lBitBlt As Long, tPos As Long
Dim TempDC As Long, TempBMP As Long, TempMerge As Long
Dim CurrentDisp As Long, lPlaceKeeper&, lCount As Long
Dim TempMaskDC As Long, ImageBMP As Long, MergeObject As Long

tPos& = 0
If DeviceContext = L_IMAGE Then 'Removing an Image from the ImageDC device context
    
    'Make sure the image trying to be removed exists
    If ImageIndex% > UBound(ImageDimensions%) Then
        Remove_ImageFromMemory = False
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    'If image index to remove is the only one or there are 0 images, delete the device context and redimension the ImageDimensions() array
    If (UBound(ImageDimensions%())) <= 0 Or UBound(ImageDimensions%()) = 1 Then
        ReDim ImageDimensions%(0, 1 To 2)
        
        Call DeleteDC(ImageDC&)
        ImageDC& = 0
        ImageDCDimensions%(1) = 0
        ImageDCDimensions%(2) = 0
        Call DeleteDC(MaskDC&)
        MaskDC& = 0
        Remove_ImageFromMemory = True
        Exit Function
    '--------------------------------------------------------------------------
    'If there are only two images in the index, remove the specified image and copy the other
    ElseIf UBound(ImageDimensions%()) = 2 Then
        'Set the Image's Index to the one other than the one being removed,
        'so we can get accurate dimensions for that one.
        If ImageIndex% = 1 Then
            ImageIndex% = 2
        ElseIf ImageIndex% = 2 Then
            ImageIndex% = 1
        End If

        'Create a temporary DC and bitmap, compatible with the current display's
        'settings, and merge the bitmap into the device context.
        CurrentDisp& = CreateDC("DISPLAY", 0&, 0&, 0&)
        TempDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2))
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new ImageDC Dimensions
        ImageDCDimensions%(1) = ImageDimensions%(ImageIndex%, 1)
        ImageDCDimensions%(2) = ImageDimensions%(ImageIndex%, 2)
        
        'Reverse the ImageIndex%() position to that of the index of the array
        'to remove again.
        If ImageIndex% = 1 Then
            tPos& = 0
        ElseIf ImageIndex% = 2 Then
            tPos& = (ImageDimensions%(1, 1) + 1)
        End If
        
        'If the MaskDC device context exists, we need to resize it, to do this
        'we have to recreate the device context and a properly sized bitmap.
        If MaskDC& <> 0 Then
            Call DeleteDC(MaskDC&) 'Delete MaskDC device context
            
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create the MaskDC device context, compatible with the monochrome
            'TempMaskDC device context.
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Copy the image mask that is not being removed to the MaskDC device
            'context
            lBitBlt& = BitBlt(MaskDC&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, tPos&, 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
         End If

        'Blit the remaining source image (NOT THE ONE BEING REMOVED), to the
        'temporary device context.
        lBitBlt& = BitBlt(TempDC&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, tPos&, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Make sure image copied.
        
        'Free resources for next use
        Call DeleteDC(ImageDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate main ImageDC device context compatible with current display,
        'create compatible bitmap to the same size as the one remaining image,
        'merge newly sized bitmap into the device context.
        ImageDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2))
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        'Make sure nothing failed
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the image we placed in the TempDC back to the main ImageDC device
        'context we newly sized.
        lBitBlt& = BitBlt(ImageDC&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Make sure image copied OK
        
        'Set variable to hold remaining image dimensions.  It's an array, only
        'because if there's more than 1 remaining image, the array gets used.
        ReDim TempArray%(1 To 1, 1 To 2)
        TempArray%(1, 1) = ImageDimensions%(ImageIndex%, 1)
        TempArray%(1, 2) = ImageDimensions%(ImageIndex%, 2)
        
        'Redimension public ImageDimensions% variable to new size
        ReDim ImageDimensions(1 To 1, 1 To 2)
        ImageDimensions%(1, 1) = TempArray%(1, 1)
        ImageDimensions%(1, 2) = TempArray%(1, 2)
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteDC(CurrentDisp&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        Remove_ImageFromMemory = True 'Success
        Exit Function
    End If
    '--------------------------------------------------------------------------
    'If the process gets this far then:
    '--------------------------------------------------------------------------
    'There are 3 or more images in the index.  Figure out which one is being
    'removed, and copy the rest of them.
    '--------------------------------------------------------------------------
    If ImageIndex% = 1 Then '3 or more entries and index is the first image.
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 2 To UBound(ImageDimensions%())
            lCount& = (lCount& + ImageDimensions%(lPlaceKeeper&, 1))
            If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(ImageDC&)
        TempBMP& = CreateCompatibleBitmap(ImageDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new ImageDC Dimensions
        ImageDCDimensions%(1) = lCount&
        ImageDCDimensions%(2) = lMaxHeight&
        
        'If the MaskDC device context exists, it also needs to be resized
        'and have the other images copied into it.
        If MaskDC& <> 0 Then
            Call DeleteDC(MaskDC&) 'Delete MaskDC device context
            
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create the MaskDC device context, compatible with the monochrome
            'TempMaskDC device context.
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Copy remaining mask images not to be removed to the MaskDC device
            'context
            lBitBlt& = BitBlt(MaskDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, ImageDimensions%(1, 1), 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        End If
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, ImageDimensions%(1, 1), 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied OK?
        
        'Free resources for use
        Call DeleteDC(ImageDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        ImageDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        'Make sure nothing failed
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(ImageDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(ImageDimensions%()) - 1), 1 To 2)
        
        'Transfer image dimensions from main array to properly sized temporary
        'array.
        For lPlaceKeeper& = 2 To UBound(ImageDimensions%())
            TempArray%((lPlaceKeeper& - 1), 1) = ImageDimensions%(lPlaceKeeper&, 1)
            TempArray%((lPlaceKeeper& - 1), 2) = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Redimension the main ImageDimensions() array to the proper size
        ReDim ImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
            ImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            ImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
        
    ElseIf ImageIndex% = UBound(ImageDimensions%()) Then 'Image is last in the index
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 1 To (UBound(ImageDimensions%()) - 1)
            lCount& = (lCount& + ImageDimensions%(lPlaceKeeper&, 1))
            If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(ImageDC&)
        TempBMP& = CreateCompatibleBitmap(ImageDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new ImageDC Dimensions
        ImageDCDimensions%(1) = lCount&
        ImageDCDimensions%(2) = lMaxHeight&
        
        'If the MaskDC device context exists, it also needs to be resized
        'and have the other images copied into it.
        If MaskDC& <> 0 Then
            Call DeleteDC(MaskDC&) 'Delete MaskDC device context
            
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create the MaskDC device context, compatible with the monochrome
            'TempMaskDC device context
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Copy remaining mask images not to be removed to the MaskDC device
            'context
            lBitBlt& = BitBlt(MaskDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        End If
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied OK?
        
        'Free resources for use
        Call DeleteDC(ImageDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        ImageDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        'Make sure nothing failed
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(ImageDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(ImageDimensions%()) - 1), 1 To 2)
        
        'Transfer image dimensions from main array to properly sized temporary
        'array.
        For lPlaceKeeper& = 1 To (UBound(ImageDimensions%()) - 1)
            TempArray%(lPlaceKeeper&, 1) = ImageDimensions%(lPlaceKeeper&, 1)
            TempArray%(lPlaceKeeper&, 2) = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&

        'Redimension the main ImageDimensions() array to the proper size
        ReDim ImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
            ImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            ImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
        
    Else 'Index is not 1 or the last (UBound())
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 1 To ImageIndex%
            If lPlaceKeeper& <> ImageIndex% Then
                lCount& = (lCount& + ImageDimensions%(lPlaceKeeper&, 1))
                If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
            End If
        Next lPlaceKeeper&
        
        'Get width of second half of device context and store in tPos
        'also update height of an image larger then the biggest so far was
        'found
        For lPlaceKeeper& = (ImageIndex + 1) To UBound(ImageDimensions%())
            tPos& = (tPos& + ImageDimensions%(lPlaceKeeper&, 1))
            If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(ImageDC&)
        TempBMP& = CreateCompatibleBitmap(ImageDC&, (lCount& + tPos&), lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new ImageDC Dimensions
        ImageDCDimensions%(1) = (lCount& + tPos&)
        ImageDCDimensions%(2) = lMaxHeight&
        
        'If the MaskDC device context exists, it also needs to be resized
        'and have the other images copied into it.
        If MaskDC& <> 0 Then
            Call DeleteDC(MaskDC&) 'Delete MaskDC device context
            
            'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
            'bitmap and merge it into the TempMaskDC
            TempMaskDC& = CreateCompatibleDC&(ImageDC&)
            ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
            MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Create the MaskDC device context, compatible with the monochrome
            'TempMaskDC device context.
            MaskDC& = CreateCompatibleDC(TempMaskDC&)
            ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
            MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
            'Free resources
            Call DeleteDC(TempMaskDC&)
            Call DeleteObject(ImageBMP&)
            Call DeleteObject(MergeObject&)
            
            'Copy remaining mask images not to be removed to the MaskDC device
            'context
            lBitBlt& = BitBlt(MaskDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
            lBitBlt& = BitBlt(MaskDC&, lCount&, 0, tPos&, lMaxHeight&, ImageDC&, (lCount& + ImageDimensions%(ImageIndex%, 1)), 0, SRCCOPY)
            If lBitBlt& = 0 Then GoTo ERROR_HANDLER
        End If
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'First half of DC copied OK?
        lBitBlt& = BitBlt(TempDC&, lCount&, 0, tPos&, lMaxHeight&, ImageDC&, (lCount& + ImageDimensions%(ImageIndex%, 1)), 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Second half of DC copied OK?
        
        'Free resources for use
        Call DeleteDC(ImageDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        ImageDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, (lCount& + tPos&), lMaxHeight&)
        TempMerge& = SelectObject(ImageDC&, TempBMP&)
        'Make sure nothing failed
        If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(ImageDC&, 0, 0, (lCount& + tPos&), lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(ImageDimensions%()) - 1), 1 To 2)
        
        'Transfer first half of image dimensions up to ImageIndex (to remove)
        'from main array to properly sized temporary array.
        For lPlaceKeeper& = 1 To (ImageIndex% - 1)
            TempArray%(lPlaceKeeper&, 1) = ImageDimensions%(lPlaceKeeper&, 1)
            TempArray%(lPlaceKeeper&, 2) = ImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        'Copy second half of image dimensions after ImageIndex (to remove)
        'from main array to the temporary array.
        tPos& = lPlaceKeeper&
        For lPlaceKeeper& = (ImageIndex% + 1) To UBound(ImageDimensions%())
            TempArray%(tPos&, 1) = ImageDimensions%(lPlaceKeeper&, 1)
            TempArray%(tPos&, 2) = ImageDimensions%(lPlaceKeeper&, 2)
            tPos& = (tPos& + 1)
        Next lPlaceKeeper&
        
        'Redimension the main ImageDimensions() array to the proper size
        ReDim ImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
            ImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            ImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
    End If
    
ElseIf DeviceContext = L_BACKGROUND Then 'Removing an image from the BackGroundDC device context.  This code is the same as above, with variables and DCs switched.
        
    'Make sure the image trying to be removed exists
    If ImageIndex% > UBound(BackGroundImageDimensions%) Then
        Remove_ImageFromMemory = False
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    'If image index to remove is the only one or there are 0 images, delete the device context and redimension the ImageDimensions() array
    If (UBound(BackGroundImageDimensions%())) <= 0 Or UBound(BackGroundImageDimensions%()) = 1 Then
        ReDim BackGroundImageDimensions%(0, 1 To 2)
        
        Call DeleteDC(BackGroundDC&)
        BackGroundDC& = 0
        BackGroundDCDimensions%(1) = 0
        BackGroundDCDimensions%(2) = 0
        Remove_ImageFromMemory = True
        Exit Function
    '--------------------------------------------------------------------------
    'If there are only two images in the index, remove the specified image and copy the other
    ElseIf UBound(BackGroundImageDimensions%()) = 2 Then
        'Set the Image's Index to the one other than the one being removed,
        'so we can get accurate dimensions for that one.
        If ImageIndex% = 1 Then
            ImageIndex% = 2
        ElseIf ImageIndex% = 2 Then
            ImageIndex% = 1
        End If

        'Create a temporary DC and bitmap, compatible with the current display's
        'settings, and merge the bitmap into the device context.
        CurrentDisp& = CreateDC("DISPLAY", 0&, 0&, 0&)
        TempDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, BackGroundImageDimensions%(ImageIndex%, 1), BackGroundImageDimensions%(ImageIndex%, 2))
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new BackGroundDC Dimensions
        BackGroundDCDimensions%(1) = BackGroundImageDimensions%(ImageIndex%, 1)
        BackGroundDCDimensions%(2) = BackGroundImageDimensions%(ImageIndex%, 2)
        
        'Reverse the ImageIndex%() position to that of the index of the array
        'to remove again.
        If ImageIndex% = 1 Then
            tPos& = 0
        ElseIf ImageIndex% = 2 Then
            tPos& = (BackGroundImageDimensions%(1, 1) + 1)
        End If

        'Blit the remaining source image (NOT THE ONE BEING REMOVED), to the
        'temporary device context.
        lBitBlt& = BitBlt(TempDC&, 0, 0, BackGroundImageDimensions%(ImageIndex%, 1), BackGroundImageDimensions%(ImageIndex%, 2), BackGroundDC&, tPos&, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Make sure image copied.
        
        'Free resources for next use
        Call DeleteDC(BackGroundDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate main ImageDC device context compatible with current display,
        'create compatible bitmap to the same size as the one remaining image,
        'merge newly sized bitmap into the device context.
        BackGroundDC& = CreateCompatibleDC(CurrentDisp&)
        TempBMP& = CreateCompatibleBitmap(CurrentDisp&, BackGroundImageDimensions%(ImageIndex%, 1), BackGroundImageDimensions%(ImageIndex%, 2))
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        'Make sure nothing failed
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the image we placed in the TempDC back to the main ImageDC device
        'context we newly sized.
        lBitBlt& = BitBlt(BackGroundDC&, 0, 0, BackGroundImageDimensions%(ImageIndex%, 1), BackGroundImageDimensions%(ImageIndex%, 2), TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Make sure image copied OK
        
        'Set variable to hold remaining image dimensions.  It's an array, only
        'because if there's more than 1 remaining image, the array gets used.
        ReDim TempArray%(1 To 1, 1 To 2)
        TempArray%(1, 1) = BackGroundImageDimensions%(ImageIndex%, 1)
        TempArray%(1, 2) = BackGroundImageDimensions%(ImageIndex%, 2)
        
        'Redimension public BackGroundImageDimensions% variable to new size
        ReDim ImageDimensions(1 To 1, 1 To 2)
        BackGroundImageDimensions%(1, 1) = TempArray%(1, 1)
        BackGroundImageDimensions%(1, 2) = TempArray%(1, 2)
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteDC(CurrentDisp&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        Remove_ImageFromMemory = True 'Success
        Exit Function
    End If
    '--------------------------------------------------------------------------
    'If the process gets this far then:
    '--------------------------------------------------------------------------
    'There are 3 or more images in the index.  Figure out which one is being
    'removed, and copy the rest of them.
    '--------------------------------------------------------------------------
    If ImageIndex% = 1 Then '3 or more entries and index is the first image.
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 2 To UBound(BackGroundImageDimensions%())
            lCount& = (lCount& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
            If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(BackGroundDC&)
        TempBMP& = CreateCompatibleBitmap(BackGroundDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new BackGroundDC Dimensions
        BackGroundDCDimensions%(1) = lCount&
        BackGroundDCDimensions%(2) = lMaxHeight&
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, BackGroundDC&, BackGroundImageDimensions%(1, 1), 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied OK?
        
        'Free resources for use
        Call DeleteDC(BackGroundDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        BackGroundDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        'Make sure nothing failed
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(BackGroundDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(BackGroundImageDimensions%()) - 1), 1 To 2)
        
        'Transfer image dimensions from main array to properly sized temporary
        'array.
        For lPlaceKeeper& = 2 To UBound(BackGroundImageDimensions%())
            TempArray%((lPlaceKeeper& - 1), 1) = BackGroundImageDimensions%(lPlaceKeeper&, 1)
            TempArray%((lPlaceKeeper& - 1), 2) = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Redimension the main ImageDimensions() array to the proper size
        ReDim BackGroundImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
            BackGroundImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            BackGroundImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
        
    ElseIf ImageIndex% = UBound(BackGroundImageDimensions%()) Then 'Image is last in the index
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 1 To (UBound(BackGroundImageDimensions%()) - 1)
            lCount& = (lCount& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
            If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(BackGroundDC&)
        TempBMP& = CreateCompatibleBitmap(BackGroundDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new BackGroundDC Dimensions
        BackGroundDCDimensions%(1) = lCount&
        BackGroundDCDimensions%(2) = lMaxHeight&
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, BackGroundDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied OK?
        
        'Free resources for use
        Call DeleteDC(BackGroundDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        BackGroundDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        'Make sure nothing failed
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(BackGroundDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(BackGroundImageDimensions%()) - 1), 1 To 2)
        
        'Transfer image dimensions from main array to properly sized temporary
        'array.
        For lPlaceKeeper& = 1 To (UBound(BackGroundImageDimensions%()) - 1)
            TempArray%(lPlaceKeeper&, 1) = BackGroundImageDimensions%(lPlaceKeeper&, 1)
            TempArray%(lPlaceKeeper&, 2) = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Redimension the main ImageDimensions() array to the proper size
        ReDim BackGroundImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
            BackGroundImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            BackGroundImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
        
    Else 'Index is not 1 or the last (UBound())
    
        'Count new width and get largest height value of images remaining in
        'ImageDC device context.
        lCount& = 0
        lMaxHeight& = 0
        For lPlaceKeeper& = 1 To ImageIndex%
            If lPlaceKeeper& <> ImageIndex% Then
                lCount& = (lCount& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
                If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
            End If
        Next lPlaceKeeper&
        
        'Get width of second half of device context and store in tPos
        'also update height of an image larger then the biggest so far was
        'found
        For lPlaceKeeper& = (ImageIndex + 1) To UBound(BackGroundImageDimensions%())
            tPos& = (tPos& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
            If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Create a temporary DC and bitmap, compatible with the ImageDC device
        'context's settings, and merge the bitmap into the device context.
        TempDC& = CreateCompatibleDC(BackGroundDC&)
        TempBMP& = CreateCompatibleBitmap(BackGroundDC&, (lCount& + tPos&), lMaxHeight&)
        TempMerge& = SelectObject(TempDC&, TempBMP&)
        'Make sure nothing failed
        If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Set new BackGroundDC Dimensions
        BackGroundDCDimensions%(1) = (lCount& + tPos&)
        BackGroundDCDimensions%(2) = lMaxHeight&
        
        'Copy remaining images from ImageDC device context to the temporary
        'device context, TempDC
        lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, BackGroundDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'First half of DC copied OK?
        lBitBlt& = BitBlt(TempDC&, lCount&, 0, tPos&, lMaxHeight&, BackGroundDC&, (lCount& + BackGroundImageDimensions%(ImageIndex%, 1)), 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Second half of DC copied OK?
        
        'Free resources for use
        Call DeleteDC(BackGroundDC&) 'Delete main ImageDC device context!
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        'Recreate the ImageDC device context (based on the temporary device
        'contexts settings), and create a newly sized bitmap to be merged into
        'the new ImageDC device context.
        BackGroundDC& = CreateCompatibleDC(TempDC&)
        TempBMP& = CreateCompatibleBitmap(TempDC&, (lCount& + tPos&), lMaxHeight&)
        TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
        'Make sure nothing failed
        If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
        
        'Copy the images we put into the temporary device context back into
        'the newly sized main ImageDC device context.
        lBitBlt& = BitBlt(BackGroundDC&, 0, 0, (lCount& + tPos&), lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER 'Images copied back OK?
        
        'Set temporary array to new size
        ReDim TempArray%(1 To (UBound(BackGroundImageDimensions%()) - 1), 1 To 2)
        
        'Transfer first half of image dimensions up to ImageIndex (to remove)
        'from main array to properly sized temporary array.
        For lPlaceKeeper& = 1 To (ImageIndex% - 1)
            TempArray%(lPlaceKeeper&, 1) = BackGroundImageDimensions%(lPlaceKeeper&, 1)
            TempArray%(lPlaceKeeper&, 2) = BackGroundImageDimensions%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        'Copy second half of image dimensions after ImageIndex (to remove)
        'from main array to the temporary array.
        tPos& = lPlaceKeeper&
        For lPlaceKeeper& = (ImageIndex% + 1) To UBound(BackGroundImageDimensions%())
            TempArray%(tPos&, 1) = BackGroundImageDimensions%(tPos&, 1)
            TempArray%(tPos&, 2) = BackGroundImageDimensions%(tPos&, 2)
        Next lPlaceKeeper&
        
        'Redimension the main ImageDimensions() array to the proper size
        ReDim BackGroundImageDimensions%(1 To UBound(TempArray%), 1 To 2)
        
        'Transfer the contents we moved into the temporary array, back to
        'the newly sized main ImageDimensions() array.
        For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
            BackGroundImageDimensions%(lPlaceKeeper&, 1) = TempArray%(lPlaceKeeper&, 1)
            BackGroundImageDimensions%(lPlaceKeeper&, 2) = TempArray%(lPlaceKeeper&, 2)
        Next lPlaceKeeper&
        
        'Free resources
        Call DeleteDC(TempDC&)
        Call DeleteObject(TempBMP&)
        Call DeleteObject(TempMerge&)
        
        Remove_ImageFromMemory = True
        Exit Function
    End If
End If

ERROR_HANDLER:
'An error has occured.  The main device contexts, ImageDC, BackGroundDC and
'MaskDC (if it exists) may be unuseable at this point.  If this function
'returns a False value for failure, I would recommend making a call to
'Unload_DeviceContext and wiping all images from memory.
Call DeleteDC(TempDC&)
Call DeleteDC(CurrentDisp&)
Call DeleteObject(TempBMP&)
Call DeleteObject(TempMerge&)
Remove_ImageFromMemory = False
End Function
Public Function Render_BackGround(ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nBackGroundWidth As Long, ByVal nBackGroundHeight As Long, ImageIndex As Integer, ByVal nXSource As Long, ByVal nYSource As Long) As Boolean
'*******************************************************************************
'This function renders an image from the BackGroundDC device context onto a
'source destination.  Render width and height of the background image can be
'controlled with this function, unlike the Render_Image function.  Therefore,
'use this function when rendering static background images.
'
'Parameters:
'hdcDest = Destination to render on
'nXDest = X coordinate to render at in destination
'nYDest = Y coordinate to render at in destination
'nBackGroundWidth = New width of source image.  If 0 is specified, it will
'   use the default width of the source file
'nBackGroundHeight = New height of source image.  If 0 is specified it will
'   use the default height of the source file
'ImageIndex = Index of source image to render
'nXSource = X coordinate to start rendering from in source image
'nYSource = Y coordinate to start rendering from in source image
'
'Returns:
'True = successful
'False = failure to render
'
'Examples:
'
'1) Render the whole background with an index of 1 (or the first background that
'   was loaded into the BackGroundDC device context) with it's default width and
'   height (which is stored in the BackGroundImageDimensions%() array), to a
'   picturebox named Picture1, at the 0 X coordinate and 0 Y coordinate in the
'   picturebox.
'
'       bReturn = Render_BackGround(Picture1.hdc, 0, 0, 0, 0, 1, 0, 0)
'
'2) Render the second half of the background with an index of 1, to a picturebox
'   named Picture1, at the X:0, Y:0 coordinates in Picture1.  To render the 2nd
'   half of the background, we must access the BackGroundImageDimensions%()
'   array, to determine the new width of background to be rendered, and also the
'   coordinates to start rendering from.
'
'       NewWidth& = (BackGroundImageDimensions%(1, 1) / 2)
'       NewHeight& = BackGroundImageDimensions%(1, 2)
'       bReturn = Render_BackGround(Picture1.hdc, 0, 0, NewWidth&, NewHeight&, 1, NewWidth&, 0)
'
'*******************************************************************************
Dim lBitBlt As Long, lImageXPos As Long

'Make sure that the index specified exists and is valid.
If ImageIndex% > UBound(BackGroundImageDimensions%()) Then GoTo ERROR_HANDLER
'Get the image's x coordinate from the DC and add the new X Source coordinate
'to this position.
lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_BACKGROUND) 'Get the image's X position in the device context
'Make sure the X and Y coordinates to start rendering from aren't equal to or
'greater than the image's width or height.  If so, error, as no image will be
'rendered if we proceed.
If nXSource& >= BackGroundImageDimensions%(ImageIndex%, 1) Then GoTo ERROR_HANDLER
If nYSource& >= BackGroundImageDimensions%(ImageIndex%, 2) Then GoTo ERROR_HANDLER

'Check whether we're rendering a full image or a portion of it.  If rendering
'full image, specified by 0& values, set Width and Height to the full image
'dimensions.
If nBackGroundWidth& = 0 Then nBackGroundWidth& = BackGroundImageDimensions%(ImageIndex%, 1)
If nBackGroundHeight& = 0 Then nBackGroundHeight& = BackGroundImageDimensions%(ImageIndex%, 2)

'Blit the background from the BackGroundDC device context to the source
'destination.
lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, nBackGroundWidth&, nBackGroundHeight&, BackGroundDC&, lImageXPos&, nYSource&, SRCCOPY)
If lBitBlt& <> 0 Then
    Render_BackGround = True 'Success
    Exit Function
End If

ERROR_HANDLER:
Render_BackGround = False 'Failure
End Function
Public Function Render_Image(ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ImageIndex As Integer, Optional RenderWithMask As Boolean = False) As Boolean
'*******************************************************************************
'This function renders an image from the ImageDC ir BackGroundDC device
'context to a source destination.  It can also render the image with a mask,
'but ONLY if the mask was created when calling Load_ImageIntoMemory.  If no
'mask exists and you want to render an image with a mask, use the
'Render_MaskedImage function instead.
'
'Parameters:
'hdcDest = Source destination
'nXDest = X coordinate destination in source
'nYDest = Y coordinate destination in source
'ImageIndex = Index of image dimensions stored in ImageDimensions() array
'   This is used to properly find the image in the main ImageDC device context
'   and also render it to the accurate dimensions.
'RenderWithMask = If set to True, the mask for the image is drawn from MaskDC
'   and then the image from ImageDC is drawn on top of the mask.
'Returns:
'True = if successful
'False = if failed
'
'Example:
'
'1) Render an image with an index of 1 (first image loaded) to the X:0, Y:0
'   coordinates a destination named Picture1.  Render the image with a
'   mask that was automatically created by calling Load_ImageIntoMemory with
'   the CreateImageMask parameter set to True.
'
'       bReturn = Render_Image(Picture1.hdc, 0, 0, 1, True)
'
'*******************************************************************************
Dim lBitBlt As Long, lImageXPos As Long, LTColor As Long

'Make sure that the index specified exists and is valid.
If ImageIndex% > UBound(ImageDimensions%()) Then GoTo ERROR_HANDLER
lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_IMAGE) 'Get the image's X position in the device context

If RenderWithMask = False Then
    'Blit the image onto the source destination
    lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, lImageXPos&, 0, SRCCOPY)
    If lBitBlt& <> 0 Then
        Render_Image = True 'Success
        Exit Function
    End If
ElseIf RenderWithMask = True Then
    'Get color of pixel of ImageIndex at x:1,y:1 from ImageDC device context
    LTColor& = GetPixel(ImageDC&, lImageXPos&, 1)
    'Set the back color of the destination
    Call SetBkColor(hdcDest&, LTColor&)
    
    'Blit the source image inverted onto the destination
    lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, lImageXPos&, 0, SRCINVERT)
    If lBitBlt& <> 0 Then
        'Render the mask onto the destination using SRCAND
        lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), MaskDC&, lImageXPos&, 0, SRCAND)
        If lBitBlt& <> 0 Then
            'Render the source image inverted onto the destination again
            lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, lImageXPos&, 0, SRCINVERT)
            If lBitBlt& <> 0 Then 'If image rendered OK
                Render_Image = True 'Success
                Exit Function
            End If
        End If
    End If
End If

ERROR_HANDLER:
Render_Image = False 'Failure
End Function
Public Function Render_ImageFromFile(ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nImageWidth As Long, ByVal nImageHeight As Long, ImageFile As String, ByVal nXSource As Long, ByVal nYSource As Long) As Boolean
'*******************************************************************************
'This function loads and renders an image straight from it's file.
'It is slightly faster than Visual Basic's built-in LoadPicture() function.
'
'Parameters:
'hdcDest = Destination to render on
'nXDest = X coordinate to render at in destination
'nYDest = Y coordinate to render at in destination
'nImageWidth = New width of source image.  If 0 is specified, it will use the default width of the source file
'nImageHeight = New height of source image.  If 0 is specified it will use the default height of the source file
'ImageFile = path to the image to render
'nXSource = X coordinate to start rendering from in source image
'nYSource = Y coordinate to start rendering from in source image
'
'Returns:
'True = successful
'False = failure to render
'
'Example:
'
'1) Draw the first 20 pixels width-wise of a bitmap and it's full height into
'   a destination named Picture1, at X:0, Y:0 coordinates.  Putting the file
'   path into a string is not necessary.
'
'   IMG$ = App.Path & "\TestImage.bmp"
'   bReturn = Render_ImageFromFile(Picture1.hdc, 0, 0, 20, 0&, IMG$, 0, 0)
'*******************************************************************************
Dim TempDC As Long, TempBMP As Long, MergeObject As Long
Dim CurrentDisp As Long, BMPInfo As BITMAP, lBitBlt As Long

'Create a DC equal to that of the current display
CurrentDisp& = CreateDC("DISPLAY", 0&, 0&, 0&)
'Create a device context compatible with the current display
TempDC& = CreateCompatibleDC(CurrentDisp&)
'Load the bitmap
TempBMP& = LoadImage(0&, ImageFile$, IMAGE_BITMAP, 0&, 0&, IM_LOADFROMFILE + IM_DEFAULTSIZE + IM_DEFAULTCOLOR)
'Get the bitmap's information and store it in the BITMAP structure
Call GetObject(TempBMP&, Len(BMPInfo), BMPInfo)
'Merge the bitmap into the device context, TempDC
MergeObject& = SelectObject(TempDC&, TempBMP&)
'Make sure nothing failed
If TempDC& = 0 Or TempBMP& = 0 Or MergeObject& = 0 Then GoTo ERROR_HANDLER
If BMPInfo.bmWidth <= 0 Or BMPInfo.bmHeight <= 0 Then GoTo ERROR_HANDLER

'Make sure the X and Y coordinates to start rendering from aren't equal to or
'greater than the image's width or height.  If so, error, as no image will be
'rendered if we proceed.
If nXSource& >= BMPInfo.bmWidth Then GoTo ERROR_HANDLER
If nYSource& >= BMPInfo.bmHeight Then GoTo ERROR_HANDLER

'Check whether we're rendering a full image or a portion of it.  If rendering
'full image, specified by 0& values, set Width and Height to the full image
'dimensions.
If nImageWidth& = 0& Then nImageWidth& = BMPInfo.bmWidth
If nImageHeight& = 0& Then nImageHeight& = BMPInfo.bmHeight

'Blit the contents of the temporary device context, TempDC, to the
'destination.
lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, nImageWidth&, nImageHeight&, TempDC&, nXSource&, nYSource&, SRCCOPY)
If lBitBlt& = 0 Then GoTo ERROR_HANDLER

'Free resources
Call DeleteDC(CurrentDisp&)
Call DeleteDC(TempDC&)
Call DeleteObject(TempBMP&)
Call DeleteObject(MergeObject&)
Render_ImageFromFile = True 'Success
Exit Function

ERROR_HANDLER:
'Free resources
Call DeleteDC(CurrentDisp&)
Call DeleteDC(TempDC&)
Call DeleteObject(TempBMP&)
Call DeleteObject(MergeObject&)
Render_ImageFromFile = False 'Failure
End Function
Public Function Render_MaskedImage(ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ImageIndex As Integer) As Boolean
'*******************************************************************************
'This function renders the image and it's mask to the source destination from
'the ImageDC device context.  The mask is created on-the-fly.  It uses a
'temporary device context so as to avoid modifying the ImageDC device context.
'
'Parameters:
'hdcDest = Source destination
'nXDest = X coordinate destination in source
'nYDest = Y coordinate destination in source
'ImageIndex = Index of image dimensions stored in ImageDimensions() array
'   This is used to properly find the image in the main ImageDC device context
'   and also render it to the accurate dimensions.
'
'Returns:
'True = if successful
'False = if failed
'
'Example:
'
'1) Draw an image with an index of 1 in the ImageDC device context, with a mask
'   created on-the-fly at the X:0, Y:0 coordinates in Picture1.
'
'       bReturn = Render_MaskedImage(Picture1.hdc, 0, 0, 1)
'
'*******************************************************************************
Dim lBitBlt As Long
Dim lProcess As Long, lImageXPos As Long, LTColor As Long
Dim lMaskDC As Long, lMaskBMP As Long, lMaskBMPMerge As Long
Dim lTempDC As Long, lTempBMP As Long, lTempBMPMerge As Long
Dim lSource As Long, lSourceBMP As Long, lSourceBMPMerge As Long

'Determine image index position in ImageDC device context
lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_IMAGE)
'Get color of pixel of ImageIndex at x:1,y:1 from ImageDC device context
LTColor& = GetPixel(ImageDC&, lImageXPos&, 1)

'Create a source device context to avoid modifying ImageDC device context
'Create compatible bitmap of same size as image of ImageIndex in ImageDC,
'dimensions are stored in ImageDimensions%() array.  Merge the bitmap into
'the device context just created, lSource.
lSource& = CreateCompatibleDC(hdcDest&)
lSourceBMP& = CreateCompatibleBitmap(hdcDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2))
lSourceBMPMerge& = SelectObject(lSource&, lSourceBMP&)
lBitBlt& = BitBlt(lSource&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), ImageDC&, lImageXPos&, 0, SRCCOPY)
If lSource& = 0 Or lSourceBMP& = 0 Or lSourceBMPMerge& = 0 Or lBitBlt& = 0 Then
    GoTo ERROR_HANDLER
End If

'Set the backcolor of the lSource device context and the backcolor of the
'destination to render on to the color of the pixel gotten from above.
'If the color is not the same, the images may not always appear transparent.
lProcess& = SetBkColor(lSource&, LTColor&)
lProcess& = SetBkColor(hdcDest&, LTColor&)

'Create a device context to hold the mask, and a bitmap of the same size
'as the image that we're going to mask.  Merge the bitmap into the device context.
lMaskDC& = CreateCompatibleDC(hdcDest&)
lMaskBMP& = CreateCompatibleBitmap(hdcDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2))
lMaskBMPMerge& = SelectObject(lMaskDC&, lMaskBMP&)
If lMaskDC& = 0 Or lMaskBMP& = 0 Or lMaskBMPMerge& = 0 Then
    GoTo ERROR_HANDLER
End If

'Create a temporary device context to hold the monochrome mask we create
'using CreateBitmap.  The size is also the size of the image we're rendering.
'Merge the mono bitmap into the device context, making it monochrome.
lTempDC& = CreateCompatibleDC(hdcDest&)
lTempBMP& = CreateBitmap(ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), 1, 1, 0&)
lTempBMPMerge& = SelectObject(lTempDC&, lTempBMP&)
If lTempDC& = 0 Or lTempBMP& = 0 Or lTempBMPMerge& = 0 Then
    GoTo ERROR_HANDLER
End If

'Blit the source image (lSource&) into the monochrome device context, making
'it a mask.
lBitBlt& = BitBlt(lTempDC&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), lSource&, 0, 0, SRCCOPY)
'Blit the monochrome mask into the lMaskDC device context for later us.
lBitBlt& = BitBlt(lMaskDC&, 0, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), lTempDC&, 0, 0, SRCCOPY)

'Blit the source inverted into the destination
lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), lSource&, 0, 0, SRCINVERT)

If lBitBlt& <> 0 Then
    'Blit the mask using AND into the destination, which makes it transparent
    lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), lMaskDC&, 0, 0, SRCAND)
    If lBitBlt& <> 0 Then
        'Finally, blit another invert of the source into the destination, showing
        'the image with a transparent background.
        lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), lSource&, 0, 0, SRCINVERT)
        
        'Free resources
        Call DeleteDC(lSource&)
        Call DeleteObject(lSourceBMP&)
        Call DeleteObject(lSourceBMPMerge&)
        Call DeleteDC(lMaskDC&)
        Call DeleteObject(lMaskBMP&)
        Call DeleteObject(lMaskBMPMerge&)
        Call DeleteDC(lTempDC&)
        Call DeleteObject(lTempBMP&)
        Call DeleteObject(lTempBMPMerge&)
        Render_MaskedImage = True 'Success
        Exit Function
    End If
End If

ERROR_HANDLER:
'Free resources
Call DeleteDC(lSource&)
Call DeleteObject(lSourceBMP&)
Call DeleteObject(lSourceBMPMerge&)
Call DeleteDC(lMaskDC&)
Call DeleteObject(lMaskBMP&)
Call DeleteObject(lMaskBMPMerge&)
Call DeleteDC(lTempDC&)
Call DeleteObject(lTempBMP&)
Call DeleteObject(lTempBMPMerge&)
Render_MaskedImage = False 'Failure
End Function
Public Function Render_StretchedImage(ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, DeviceContext As m2DR_Type, ImageIndex As Integer, ByVal nWidth As Long, ByVal nHeight As Long, Optional StretchRenderType As m2DR_StretchType = 0) As Boolean
'*******************************************************************************
'This function renders an image from either the ImageDC device context or
'the BackGroundDC device context, with new dimensions specified in nWidth
'and nHeight.
'
'Parameters:
'hdcDest = Source destination
'nXDest = X coordinate destination in source
'nYDest = Y coordinate destination in source
'DeviceContext = Specify which device context the source image is in
'ImageIndex = Index of image dimensions stored in ImageDimensions() array
'   This is used to properly find the image in the main ImageDC device context
'   and also render it to the accurate dimensions.
'nWidth = The new width of the image to be rendered
'nHeight = The new height of the image to be rendered
'StretchRenderType = Only applies to ImageDC device context (L_IMAGE)
'   R_IMAGE = Renders the normal image, stretched
'   R_IMAGE_CREATE_MASK = Creates a mask on-the-fly and renders
'      the image masked
'   R_IMAGE_MASK = Renders the image masked using a mask that was created
'       when calling the Load_ImageIntoMemory function.

'Returns:
'True = if successful
'False = if failed
'
'Example:
'
'1) Draw an image with an index of 2 from the ImageDC device context.  Make the
'   image twice it's original size, using a mask that was created earlier when
'   Load_ImageIntoMemory() was called with the CreateImageMask parameter set to
'   true.  Draw the stretched, masked image to the destination Picture1, at the
'   coordinates of X:0, Y:0.
'
'       NewWidth& = (ImageDimensions%(2, 1) * 2)
'       NewHeight& = (ImageDimensions%(2, 2) * 2)
'       bReturn = Render_StretchedImage(Picture1.hdc, 0, 0, L_IMAGE, 2, NewWidth&, NewHeight&, R_IMAGE_MASK)
'
'*******************************************************************************
Dim lBitBlt As Long, sTime As Long
Dim lProcess As Long, lImageXPos As Long, LTColor As Long
Dim lMaskDC As Long, lMaskBMP As Long, lMaskBMPMerge As Long
Dim lTempDC As Long, lTempBMP As Long, lTempBMPMerge As Long
Dim lSource As Long, lSourceBMP As Long, lSourceBMPMerge As Long

If DeviceContext = L_IMAGE Then
    'Make sure that the index specified exists and is valid.
    If ImageIndex% > UBound(ImageDimensions%()) Then GoTo ERROR_HANDLER
    lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_IMAGE) 'Get the image's X position in the device context
    
    If StretchRenderType = R_IMAGE Then 'Render standard image
        'Stretch and blit the image onto the destination
        lBitBlt& = StretchBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, ImageDC&, lImageXPos&, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), SRCCOPY)
        If lBitBlt& <> 0 Then
            Render_StretchedImage = True 'Success
            Exit Function
        Else
            Render_StretchedImage = False 'Failure to render on destination
            Exit Function
        End If
    ElseIf StretchRenderType = R_IMAGE_MASK Then
        'Determine image index position in ImageDC device context
        lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_IMAGE)
        'Get color of pixel of ImageIndex at x:1,y:1 from ImageDC device context
        LTColor& = GetPixel(ImageDC&, lImageXPos&, 1)
        'Set the back color of the destination
        Call SetBkColor(hdcDest&, LTColor&)
        
        'Stretch and blit the source inverted onto the destination
        lBitBlt& = StretchBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, ImageDC&, lImageXPos&, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), SRCINVERT)
        If lBitBlt& <> 0 Then
            'Stretch and blit the mask using AND onto the destination
            lBitBlt& = StretchBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, MaskDC&, lImageXPos&, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), SRCAND)
            If lBitBlt& <> 0 Then
                'Render the source image inverted onto the destination again
                lBitBlt& = StretchBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, ImageDC&, lImageXPos&, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), SRCINVERT)
                If lBitBlt& <> 0 Then 'Image drawn on mask OK?
                    Render_StretchedImage = True 'Success
                    Exit Function
                Else
                    Render_StretchedImage = False 'Failure to render on destination
                    Exit Function
                End If
            Else
                Render_StretchedImage = False 'Failure to render on destination
                Exit Function
            End If
        Else
            Render_StretchedImage = False 'Failure to render on destination
            Exit Function
        End If
    ElseIf StretchRenderType = R_IMAGE_CREATE_MASK Then 'Create a stretched mask
        'Determine image index position in ImageDC device context
        lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_IMAGE)
        'Get color of pixel of ImageIndex at x:1,y:1 from ImageDC device context
        LTColor& = GetPixel(ImageDC&, lImageXPos&, 1)
        
        'Create a source device context to avoid modifying ImageDC device context
        'Create compatible bitmap of same size as image of ImageIndex in ImageDC,
        'dimensions are stored in ImageDimensions%() array.  Merge the bitmap into
        'the device context just created, lSource.
        lSource& = CreateCompatibleDC(hdcDest&)
        lSourceBMP& = CreateCompatibleBitmap(hdcDest&, nWidth&, nHeight&)
        lSourceBMPMerge& = SelectObject(lSource&, lSourceBMP&)
        lBitBlt& = StretchBlt(lSource&, 0, 0, nWidth&, nHeight&, ImageDC&, lImageXPos&, 0, ImageDimensions%(ImageIndex%, 1), ImageDimensions%(ImageIndex%, 2), SRCCOPY)
        If lSource& = 0 Or lSourceBMP& = 0 Or lSourceBMPMerge& = 0 Or lBitBlt& = 0 Then
            GoTo ERROR_HANDLER
        End If

        'Set the backcolor of the lSource device context and the backcolor of the
        'destination to render on to the color of the pixel gotten from above.
        'If the color is not the same, the images may not always appear transparent.
        lProcess& = SetBkColor(lSource&, LTColor&)
        lProcess& = SetBkColor(hdcDest&, LTColor&)
        
        'Create a device context to hold the mask, and a bitmap of the same size
        'as the image that we're going to mask.  Merge the bitmap into the device context.
        lMaskDC& = CreateCompatibleDC(hdcDest&)
        lMaskBMP& = CreateCompatibleBitmap(hdcDest&, nWidth&, nHeight&)
        lMaskBMPMerge& = SelectObject(lMaskDC&, lMaskBMP&)
        If lMaskDC& = 0 Or lMaskBMP& = 0 Or lMaskBMPMerge& = 0 Then
            GoTo ERROR_HANDLER
        End If
        
        'Create a temporary DC and a 1 pixel by one pixel monochrome bitmap.
        'Merge the two together.
        lTempDC& = CreateCompatibleDC(ImageDC&)
        lTempBMP& = CreateBitmap(1, 1, 1, 1, 0&)
        lTempBMPMerge& = SelectObject(lTempDC&, lTempBMP&)
        If lTempDC& = 0 Or lTempBMP& = 0 Or lTempBMPMerge& = 0 Then GoTo ERROR_HANDLER
        'Free resources for use
        Call DeleteObject(lTempBMP&)
        Call DeleteObject(lTempBMPMerge&)
        
        'Now create a bitmap compatible with the TempDC, which is monochrome.
        'Merge this PROPERLY SIZED bitmap into TempDC device context.
        '(note: for some reason I'm not sure of, CreateBitmap was refusing
        'to take rather small width and height dimensions, and I found this
        'way around the odd ?error?)
        lTempBMP& = CreateCompatibleBitmap(lTempDC&, nWidth&, nHeight&)
        lTempBMPMerge& = SelectObject(lTempDC&, lTempBMP&)
        If lTempDC& = 0 Or lTempBMP& = 0 Or lTempBMPMerge& = 0 Then GoTo ERROR_HANDLER

        'Blit the source image (lSource&) into the monochrome device context, making
        'it a mask.
        lBitBlt& = BitBlt(lTempDC&, 0, 0, nWidth&, nHeight&, lSource&, 0, 0, SRCCOPY)
        sTime& = GetTickCount()
        'Blit the monochrome mask into the lMaskDC device context for later us.
        lBitBlt& = BitBlt(lMaskDC&, 0, 0, nWidth&, nHeight&, lTempDC&, 0, 0, SRCCOPY)
        
        'Blit the source inverted into the destination
        lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, lSource&, 0, 0, SRCINVERT)
    
        If lBitBlt& <> 0 Then
            'Blit the mask using AND into the destination, which makes it transparent
            lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, lMaskDC&, 0, 0, SRCAND)
            If lBitBlt& <> 0 Then
                'Finally, blit another invert of the source into the destination, showing
                'the image with a transparent background.
                lBitBlt& = BitBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, lSource&, 0, 0, SRCINVERT)
                
                'Free resources
                Call DeleteDC(lSource&)
                Call DeleteObject(lSourceBMP&)
                Call DeleteObject(lSourceBMPMerge&)
                Call DeleteDC(lMaskDC&)
                Call DeleteObject(lMaskBMP&)
                Call DeleteObject(lMaskBMPMerge&)
                Call DeleteDC(lTempDC&)
                Call DeleteObject(lTempBMP&)
                Call DeleteObject(lTempBMPMerge&)
                Render_StretchedImage = True 'Success
                Exit Function
            End If
        End If
        
        GoTo ERROR_HANDLER
    End If
ElseIf DeviceContext = L_BACKGROUND Then
    'Make sure that the index specified exists and is valid.
    If ImageIndex% > UBound(BackGroundImageDimensions%()) Then GoTo ERROR_HANDLER
    lImageXPos& = Get_ImageXPosFromDC(ImageIndex%, L_BACKGROUND) 'Get the image's X position in the device context

    'Blit the image onto the source destination
    lBitBlt& = StretchBlt(hdcDest&, nXDest&, nYDest&, nWidth&, nHeight&, BackGroundDC&, lImageXPos&, 0, BackGroundImageDimensions%(ImageIndex%, 1), BackGroundImageDimensions%(ImageIndex%, 2), SRCCOPY)
    If lBitBlt& <> 0 Then
        Render_StretchedImage = True 'Success
        Exit Function
    Else
        GoTo ERROR_HANDLER
    End If
End If

ERROR_HANDLER:
'Free resources
Call DeleteDC(lSource&)
Call DeleteObject(lSourceBMP&)
Call DeleteObject(lSourceBMPMerge&)
Call DeleteDC(lMaskDC&)
Call DeleteObject(lMaskBMP&)
Call DeleteObject(lMaskBMPMerge&)
Call DeleteDC(lTempDC&)
Call DeleteObject(lTempBMP&)
Call DeleteObject(lTempBMPMerge&)
Render_StretchedImage = False 'Failure
End Function
Public Function Resize_DeviceContext(DeviceContext As m2DR_Type) As Boolean
'*******************************************************************************
'This function resizes the device context to the dimensions of the images stored
'inside it, thereby saving memory (if the device context's dimensions are greater
'than the images inside it)
'
'Parameters:
'DeviceContext = Which device context to resize:  ImageDC or BackGroundDC
'   If the DeviceContext is L_IMAGE (ImageDC device context) then the MaskDC
'   will be resized also and ALL images will have masks created for them, even
'   if you specified NOT to CreateImageMasks when the call was made to
'   Load_ImageIntoMemory.
'
'Returns:
'True = Device Context resized OK
'False = No images in device context OR error resizing
'
'Example:
'
'1) Resize the background device context (BackGroundDC)
'
'       bReturn = Resize_DeviceContext(L_BACKGROUND)
'
'*******************************************************************************
Dim lCount As Long, lMaxHeight As Long, lPlaceKeeper As Long
Dim TempMaskDC As Long, ImageBMP As Long, MergeObject As Long
Dim lBitBlt As Long, TempDC As Long, TempBMP As Long, TempMerge As Long

lCount& = 0
lMaxHeight& = 0
If DeviceContext = L_IMAGE Then 'Resize ImageDC device context
    
    'Make sure there's at least one image in the ImageDC device context
    If UBound(ImageDimensions%()) <= 0 Then
        Resize_DeviceContext = False
        Exit Function
    End If
    
    'Calculate proper total width and get tallest image dimensions
    For lPlaceKeeper& = 1 To UBound(ImageDimensions%())
        lCount& = (lCount& + ImageDimensions%(lPlaceKeeper&, 1))
        If ImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = ImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    'Create a temporary device context compatible with ImageDC's settings,
    'create a bitmap based on the width of all images in the ImageDC
    'device context, and with the height of the tallest image.  Merge the
    'bitmap into the temp device context.
    TempDC& = CreateCompatibleDC(ImageDC&)
    TempBMP& = CreateCompatibleBitmap(ImageDC&, lCount&, lMaxHeight&)
    TempMerge& = SelectObject(TempDC&, TempBMP&)
    If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
    
    'Record new ImageDC device context dimensions
    ImageDCDimensions%(1) = lCount&
    ImageDCDimensions%(2) = lMaxHeight&
    
    'Copy contents of ImageDC device context to TempDC
    lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, ImageDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    
    If MaskDC& <> 0 Then
        Call DeleteDC(MaskDC&) 'Delete MaskDC device context
            
        'Create a device context equal to ImageDC. Create a 1x1 pixel monochrome
        'bitmap and merge it into the TempMaskDC
        TempMaskDC& = CreateCompatibleDC&(ImageDC&)
        ImageBMP& = CreateBitmap(1, 1, 1, 1, 0&)
        MergeObject& = SelectObject(TempMaskDC&, ImageBMP&)
            
        'Free resources
        Call DeleteObject(ImageBMP&)
        Call DeleteObject(MergeObject&)
            
        'Create the MaskDC device context, compatible with the monochrome
        'TempMaskDC device context.
        MaskDC& = CreateCompatibleDC(TempMaskDC&)
        ImageBMP& = CreateCompatibleBitmap(MaskDC&, ImageDCDimensions%(1), ImageDCDimensions%(2))
        MergeObject = SelectObject(MaskDC&, ImageBMP&)
            
        'Free resources
        Call DeleteDC(TempMaskDC&)
        Call DeleteObject(ImageBMP&)
        Call DeleteObject(MergeObject&)
            
        'Copy the image mask that is not being removed to the MaskDC device
        'context
        lBitBlt& = BitBlt(MaskDC&, 0, 0, ImageDCDimensions%(1), ImageDCDimensions%(2), ImageDC&, 0, 0, SRCCOPY)
        If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    End If
    
    'Free resources
    Call DeleteDC(ImageDC&) 'Delete main ImageDC device context!
    Call DeleteObject(TempBMP&)
    Call DeleteObject(TempMerge&)
    
    'Recreate the main ImageDC device context compatible with TempDC settings,
    'create properly sized bitmap and merge bitmap into device context
    ImageDC& = CreateCompatibleDC(TempDC&)
    TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
    TempMerge& = SelectObject(ImageDC&, TempBMP&)
    If ImageDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
    
    'Copy contents of temporary device context back to newly sized ImageDC
    'device context.
    lBitBlt& = BitBlt(ImageDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    
    'Free resources
    Call DeleteDC(TempDC&)
    Call DeleteObject(TempBMP&)
    Call DeleteObject(TempMerge&)
    
    Resize_DeviceContext = True 'Success
    Exit Function
ElseIf DeviceContext = L_BACKGROUND Then 'Resize BackGroundDC device context
    'Make sure there's at least one image in the BackGroundDC device context
    If UBound(BackGroundImageDimensions%()) <= 0 Then
        Resize_DeviceContext = False
        Exit Function
    End If
    
    'Calculate proper total width and get tallest image dimensions
    For lPlaceKeeper& = 1 To UBound(BackGroundImageDimensions%())
        lCount& = (lCount& + BackGroundImageDimensions%(lPlaceKeeper&, 1))
        If BackGroundImageDimensions%(lPlaceKeeper&, 2) > lMaxHeight& Then lMaxHeight& = BackGroundImageDimensions%(lPlaceKeeper&, 2)
    Next lPlaceKeeper&
    
    'Create a temporary device context compatible with BackGroundDC's settings,
    'create a bitmap based on the width of all images in the BackGroundDC
    'device context, and with the height of the tallest image.  Merge the
    'bitmap into the temp device context.
    TempDC& = CreateCompatibleDC(BackGroundDC&)
    TempBMP& = CreateCompatibleBitmap(BackGroundDC&, lCount&, lMaxHeight&)
    TempMerge& = SelectObject(TempDC&, TempBMP&)
    If TempDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
    
    'Copy contents of BackGroundDC device context to TempDC
    lBitBlt& = BitBlt(TempDC&, 0, 0, lCount&, lMaxHeight&, BackGroundDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    
    'Record new ImageDC device context dimensions
    BackGroundDCDimensions%(1) = lCount&
    BackGroundDCDimensions%(2) = lMaxHeight&
    
    'Free resources
    Call DeleteDC(BackGroundDC&) 'Delete main BackGroundDC device context!
    Call DeleteObject(TempBMP&)
    Call DeleteObject(TempMerge&)
    
    'Recreate the main BackGroundDC device context compatible with TempDC
    'settings, create properly sized bitmap and merge bitmap into device context
    BackGroundDC& = CreateCompatibleDC(TempDC&)
    TempBMP& = CreateCompatibleBitmap(TempDC&, lCount&, lMaxHeight&)
    TempMerge& = SelectObject(BackGroundDC&, TempBMP&)
    If BackGroundDC& = 0 Or TempBMP& = 0 Or TempMerge& = 0 Then GoTo ERROR_HANDLER
    
    'Copy contents of temporary device context back to newly sized BackGroundDC
    'device context.
    lBitBlt& = BitBlt(BackGroundDC&, 0, 0, lCount&, lMaxHeight&, TempDC&, 0, 0, SRCCOPY)
    If lBitBlt& = 0 Then GoTo ERROR_HANDLER
    
    'Free resources
    Call DeleteDC(TempDC&)
    Call DeleteObject(TempBMP&)
    Call DeleteObject(TempMerge&)
    
    Resize_DeviceContext = True 'Success
    Exit Function
End If

ERROR_HANDLER:

'Free resources
Call DeleteDC(TempDC&)
Call DeleteObject(TempBMP&)
Call DeleteObject(TempMerge&)
Resize_DeviceContext = False 'Failure
End Function
Public Sub Unload_DeviceContext(DeviceContext As m2DR_Type)
'*******************************************************************************
'This function deletes the specified device context from memory, and also the
'associated public array holding the dimensions for the images that were in the
'device context.  The MaskDC device context is alway stays equal to the
'ImageDC device context.
'
'Parameters:
'DeviceContext = Specifies which device context to delete.
'
'Example:
'
'1) Unload the ImageDC device context
'
'       Call Unload_DeviceContext(L_IMAGE)
'
'*******************************************************************************
If DeviceContext = L_IMAGE Then
    ImageDCDimensions%(1) = 0
    ImageDCDimensions%(2) = 0
    ReDim ImageDimensions%(0, 1 To 2)
    Call DeleteDC(ImageDC&)
    ImageDC& = 0
    Call DeleteDC(MaskDC)
    MaskDC& = 0
    Exit Sub
ElseIf DeviceContext = L_BACKGROUND Then
    BackGroundDCDimensions%(1) = 0
    BackGroundDCDimensions%(2) = 0
    ReDim BackGroundImageDimensions%(0, 1 To 2)
    Call DeleteDC(BackGroundDC&)
    BackGroundDC& = 0
    Exit Sub
End If
End Sub
Public Sub Unload_m2DRender()
'*******************************************************************************
'This unloads the module's stored device contexts and erases all stored
'image dimensions.  If you plan to use the rendering and loading functions
'of the module again, you must re-initialize it.  This sub should always
'be called before ending your program if you have initialized the module.
'
'Example:
'
'1) Unload the module
'
'       Call Unload_m2DRender
'*******************************************************************************
DoEvents: Call DeleteDC(ImageDC&)
ImageDC& = 0
ImageDCDimensions%(1) = 0
ImageDCDimensions%(2) = 0
ReDim ImageDimensions%(0, 0)
DoEvents: Call DeleteDC(MaskDC&)
MaskDC& = 0
DoEvents: Call DeleteDC(BackGroundDC&)
BackGroundDC& = 0
BackGroundDCDimensions%(1) = 0
BackGroundDCDimensions%(2) = 0
ReDim BackGroundImageDimensions%(0, 0)
End Sub
