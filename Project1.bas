Attribute VB_Name = "Miscellaneous"
Public MainHDC As Long
Public BufferHDC As Long
Public Const Pi As Double = 3.14159265358979
Public Const Radians As Double = (2 * Pi) / 360

'Constants for the GenerateDC function
'**LoadImage Constants**
Public Const IMAGE_BITMAP As Long = 0
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const LR_DEFAULTSIZE As Long = &H40
'ALL API CALLS:

Public NUM As Integer
Public Declare Function StretchBlt Lib "gdi32" _
    (ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTickCount Lib "Kernel32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function AlphaBlending Lib "Alphablending.dll" _
       (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
        ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public ResultRegion As Long

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public PolyPoints(4) As POINTAPI

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const HWND_TOPMOST = -1

Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXFULLSCREEN = 16

Public BackTile1 As Long
Public Const Tilewidth As Long = 3200
Public Const firstTileHeight As Long = 800

'Public Const SecondTileHeight As Long = 1600
'Public Const ThirdTileHeight As Long = 2400
'Public Const FourthTileHeight As Long = 3200
'General position variables
Public BackxPos As Long
Public OverlapBottom As Long
Public OverlapTop As Long

Public yrisoluz As Single
Public xrisoluz As Single

Public Const NumOfStars = 50
Public Const BulletSpeed = 100

Public Const HAcc As Single = 30
Public Const HDel As Single = 30
Public Const VAcc As Single = 30
Public Const VDel As Single = 30
Public Const BadBulletSpeed = 5
Public Const Damagelimit = 5
Public Const KeySpeed = 50
Public Const NumOfBullets = 10

Public Const ViewportHeight As Long = 600
Public Const ViewportWidth As Long = 800
Public Const BackTileHeight As Long = 600
Public Const BackTileWidth As Long = 800
'ALL TYPES:
Public Type Star
    X As Integer
    Y As Integer
    bright As Byte
    SPEED As Byte
End Type

Public Type planet
    X As Integer
    Y As Integer
End Type

Public Type upgr
    X As Integer
    Y As Integer
    mask As Object
    frame As Integer
End Type

Public Type bullet
    X As Integer
    Y As Integer
    Velocity As Integer
    activated As Byte
End Type

Public Type BadBullet
    X As Integer
    Y As Integer
    Velocity As Integer
    activated As Integer
    angle As Integer
End Type

Public Type Particles
    X As Integer
    Y As Integer
    Velocity As Integer
    activated As Integer
End Type

Public Type BadGuy
    PicT As Object
    mask As Object
    xsize As Integer
    ysize As Integer
    X As Integer
    Y As Integer
    oldX As Integer
    oldY As Integer
    Exploding As Integer
    ExplodingFrame As Double
    frame As Long
    indice As Integer
    activated As Integer
    DstX As Integer
    DstY As Integer
    Damage As Integer
    Velocity As Integer
    Bulletl(0 To 10) As BadBullet
    Bulletc(0 To 10) As BadBullet
    Bulletr(0 To 10) As BadBullet
    Bulletla(0 To 10) As Single
    Bulletra(0 To 10) As Single
    Bulletca(0 To 10) As Single
    BulletsActivated As Byte
    bulletlxpos As Integer
    bulletlypos As Integer
    bulletrxpos As Integer
    bulletrypos As Integer
    bulletcxpos As Integer
    bulletcypos As Integer
    particle(0 To 9) As Particles
    Firing As Byte
    boomed As Boolean
    shot As Boolean
    inverse As Boolean
End Type

Public Type Level
    NumOfBadGuys As Integer
    Damage As Integer
    Damagelimit As Integer
    Velocity As Integer
    OddsOfFiring As Integer
    BulletSpeed As Integer
End Type

Public Type PointXY
    X As Integer
    Y As Integer
End Type

Public Type RGB
    R As Integer
    G As Integer
    b As Integer
End Type

Public Type tStar
  nX As Long
  nY As Long
  nSpeed As Integer
  nColor As Long
End Type
Public G_dStar(599) As tStar

'ALL public VARIABLES:
Public Health As Single
Public BufferWidth As Long
Public BufferHeight As Long


Public Sub UpdateHealth()
If Health >= 0 Then Form1.bar(Health).Visible = False
End Sub

Public Sub DrawHealthBar()
For i = 0 To 18
Form1.bar(i).Visible = True
Next i
End Sub
