Attribute VB_Name = "PrintWall"
'submitted by Tom Pydeski
'Example for printing text onto the Desktop Wallpaper file.
'The name of the wallpaper file is retrieved from the registry and the picture is loaded
'into an image control to resize it to fit the screen size.  The image is then BitBlt'd
'into a picture control, where the desired message is printed to the image in the lower
'left hand corner.  The modified picture is then saved to a temp file (so as not to modify
'the original file).  The temp file is then set as the new wallpaper.
Option Explicit
Dim frmWall As VB.Form
Dim imgWall As VB.Image
Dim picWall As VB.PictureBox
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Dim fName As String
Dim fPath As String
Dim lRetVal As Long
Dim vValue As Variant
Const HKEY_CURRENT_USER = &H80000001
Const KEY_ALL_ACCESS = &H3F
Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4
Const ERROR_SUCCESS = 0&
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim WallName As String
' This API function allows us to change the parent of any component that has a hWnd
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Sub UpdateWall(Message As String)
Dim a$
Dim ThisKey As Long
Dim SubKey$
Dim hKey As Long
Dim sValueName As String
Dim ww As Integer
Dim hh As Integer
Dim pforeColor As Long
Dim pBackColor As Long
Dim picX As Integer
Dim picY As Integer
Dim ColorDiff As Long
Screen.MousePointer = 11
'get desktop wallpaper filename from the registry
ThisKey = HKEY_CURRENT_USER
SubKey$ = "Control Panel\desktop"
lRetVal = RegOpenKeyEx(ThisKey, SubKey$, 0, KEY_ALL_ACCESS, hKey) 'phkResult)
sValueName = "Wallpaper"
lRetVal = QueryValueEx(hKey, sValueName, vValue)
'vvalue contains the background picture's filename
WallName = Left(vValue, Len(vValue) - 1)
'
'create our form
Set frmWall = New Form1 'VB.Forms.Add("VB.Form", "frmWall")
'add the image control to the form
Set imgWall = frmWall.Controls.Add("VB.Image", "imgWall", frmWall)
'add the picture to the form
Set picWall = frmWall.Controls.Add("VB.PictureBox", "picWall", frmWall)
'set the properties of our picture for our desired font info
With picWall
    .AutoRedraw = -1         'True
    .BackColor = 0 'black
    .BorderStyle = 0        'None
    .FontName = "Tahoma"
    .FontSize = 11.25
    .Font.Charset = 0
    .Font.Weight = 400
    .Font.Underline = 0              'False
    .Font.Italic = 0                 'False
    .Font.Strikethrough = 0          'False
    .ForeColor = &H8000000E
    .Left = 0
    .Top = 0
    .Width = Screen.Width
    .Height = Screen.Height
End With
frmWall.WindowState = vbMaximized
'get the screen width and height
ww = Screen.Width / Screen.TwipsPerPixelX
hh = Screen.Height / Screen.TwipsPerPixelY
'
'set our image to resize the picture based on its size
'we need to do this because usually the desktop resizes the image when it is set
'to stretch the image to fill the desktop.
imgWall.Stretch = True
imgWall.Left = 0
imgWall.Top = 0
imgWall.Width = Screen.Width
imgWall.Height = Screen.Height
'load desktop image
imgWall.Picture = LoadPicture(WallName)
imgWall.ZOrder 0
'set the co-ordinates to write in the lower left corner of the screen
picX = 100
picY = Screen.Height - 400
'paint the stretched image to our picture box
picWall.PaintPicture imgWall.Picture, 0, 0, Screen.Width, Screen.Height
'get the backcolor of the pixel behind where we are going to write so we blend in
pBackColor = picWall.Point(picX, picY)
'set the backcolor of the picture box to match the pixel color we just retrieved
'we do this so when we use fonttransparent=false, the text background color will match
'the color of the picture we are writing on
picWall.BackColor = pBackColor
'set the forecolor to be the complement of the backcolor
'this works for most colors by painting black on white; yellow on blue; etc.
'some colors don't look right, but the alternative is to always have the forecolor
'be white, which won't look good on lighter desktops
pforeColor = vbWhite - pBackColor
ColorDiff = 3000000
'to account for some colors near the middle of the color spectrum, we will pick
'a constant difference and determine if the forcolor and backcolor are too close
'together.  if they are, we will print with the forecolor fixed to white
If Abs(pforeColor - pBackColor) < ColorDiff Then
    pforeColor = vbWhite
End If
picWall.ForeColor = pforeColor
'repaint the picture because changing the backcolor erases our image
picWall.PaintPicture imgWall.Picture, 0, 0, Screen.Width, Screen.Height
'use bitblt to copy the image from our image control to the picturebox device context
BitBlt picWall.hDC, 0, 0, ww, hh, imgWall, 0, 0, &HCC0020
picWall.ZOrder 0
'
'lets try to print the date on the desktop
WallName = "C:\windows\temp.bmp"
'set our print co-ordinates
picWall.CurrentX = picX
picWall.CurrentY = picY
'make the font transparent to write on the wallpaper
picWall.FontTransparent = True
'print our message
picWall.Print Message;
'the date will change each day, and won't look right as it overwrites itself
'therefore, set the fonttransparent to false so we write over the whole date
picWall.FontTransparent = False
picWall.Print Date
'now set the picture to the image.  If we don't do this, we won't capture our printing
picWall.Picture = picWall.Image
'save the picture's image to the temp file we will use as our new wallpaper file
'we do this so we don't modify the original picture file
SavePicture picWall.Image, WallName
'change the desktop wallpaper to our temp file
SetWall WallName
'this just shows us the colors
Debug.Print "BackColor="; pBackColor; " ("; Hex$(pBackColor); ")"; " - ";
Debug.Print "foreColor="; pforeColor; " ("; Hex$(pforeColor); ")", pforeColor - pBackColor
'that's it - we're done
Unload frmWall
Screen.MousePointer = 0
End Sub

Sub SetWall(Wname As String)
Dim lret As Long
Screen.MousePointer = 11
'set wallpaper to our new File
lret = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, Wname, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
Beep
Screen.MousePointer = 0
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
'function to get a value from the registry
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String
On Error GoTo QueryValueExError
' Determine the size and type of data to be read
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> 0 Then Error 5
Select Case lType
    ' For strings
Case REG_SZ:
    sValue = String(cch, 0)
    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
    If lrc = 0 Then
        vValue = Left$(sValue, cch)
    Else
        vValue = Empty
    End If
    ' For DWORDS
Case REG_DWORD:
    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
    If lrc = 0 Then vValue = lValue
Case Else
    'all other data types not supported
    lrc = -1
End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function

