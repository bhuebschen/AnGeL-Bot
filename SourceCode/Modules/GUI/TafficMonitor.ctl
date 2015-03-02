VERSION 5.00
Begin VB.UserControl TafficMonitor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   Begin VB.PictureBox Traffic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   0
      Width           =   6285
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   2520
      Tag             =   "1"
      Top             =   840
   End
End
Attribute VB_Name = "TafficMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal NIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Boolean
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Boolean
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Dim Counted As Byte, CPUHeight As Integer, CPUWidth As Integer
Dim CPUHdc As Long, CPUProz As Single
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Dim Before As Byte, aPro As Byte, Want3d As Boolean
Private Type RGBPoint
  r As Byte
  g As Byte
  b As Byte
  Pos As Long
End Type
Dim Pt(30) As RGBPoint
Dim Range As Long, AnzPoints As Integer
Private Const MaxByteIn As Long = 32767
Private Const MaxByteOut As Long = 32767
Dim oldSessionBytesSent As Currency, oldSessionBytesReceived As Currency
Dim CurrSessionBytesSent As Currency, CurrSessionBytesReceived As Currency
Public Property Get TMode() As Integer
  TMode = CInt(Timer1.Tag)
End Property
Public Property Let TMode(NewMode As Integer)
  Timer1.Tag = CInt(NewMode)
End Property
Public Property Get Enabled() As Boolean
  Enabled = CBool(Timer1.Enabled)
End Property
Public Property Let Enabled(NewMode As Boolean)
  Timer1.Enabled = NewMode
End Property
Public Property Let SetRange(ByVal AnzPos As Long)
   Range = AnzPos
End Property
Sub Clear()
   AnzPoints = 0
End Sub
Sub AddPoint(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
   AnzPoints = AnzPoints + 1
   With Pt(AnzPoints)
      .r = r
      .g = g
      .b = b
   End With
End Sub
Sub InitMixer()
Dim u As Integer
   For u = 1 To AnzPoints
      Pt(u).Pos = (Range / (AnzPoints - 1) * (u - 1))
   Next u
End Sub
Function GetCol(ByVal posi As Long)
Dim FinalR As Long, FinalG As Long, FinalB As Long, u As Integer, far As Long
   far = Range / (AnzPoints - 1)
   For u = 1 To AnzPoints
      FinalR = FinalR + Pos(Pt(u).r / far * (far - Abs(posi - Pt(u).Pos)))
      FinalG = FinalG + Pos(Pt(u).g / far * (far - Abs(posi - Pt(u).Pos)))
      FinalB = FinalB + Pos(Pt(u).b / far * (far - Abs(posi - Pt(u).Pos)))
   Next u
   GetCol = RGB(FinalR, FinalG, FinalB)
End Function
Private Function Pos(ByVal a As Long) As Long
   If a > 0 Then Pos = a
End Function
Sub InitializeCPUGraph()
  Dim Text As String, Buffer As POINTAPI
  Traffic.Cls
  CPUHdc = Traffic.hdc
  CPUHeight = Traffic.ScaleHeight - 1
  CPUWidth = Traffic.ScaleWidth
  CPUProz = CPUHeight / 100
  Clear
  SetRange = CPUHeight
  AddPoint 200, 0, 0
  AddPoint 255, 0, 0
  AddPoint 255, 255, 0
  AddPoint 0, 255, 0
  InitMixer
  Call LineTo(CPUHdc, CPUWidth, 0)
  Call MoveToEx(CPUHdc, 0, CPUHeight, Buffer)  'Linie unten
  Call LineTo(CPUHdc, CPUWidth, CPUHeight)
  If Not SimplePaper Then ChangeColor CPUHdc, RGB(0, 100, 0)
  Call MoveToEx(CPUHdc, 0, CPUHeight / 2, Buffer) '50-Prozent-Linie
  Call LineTo(CPUHdc, CPUWidth, CPUHeight / 2)
  If Not SimplePaper Then ChangeColor CPUHdc, RGB(0, 80, 0)
  Call MoveToEx(CPUHdc, 0, CPUHeight - CPUProz * 75, Buffer) '75-Prozent-Linie
  Call LineTo(CPUHdc, CPUWidth, CPUHeight - CPUProz * 75)
  Call MoveToEx(CPUHdc, 0, CPUHeight - CPUProz * 25, Buffer) '25-Prozent-Linie
  Call LineTo(CPUHdc, CPUWidth, CPUHeight - CPUProz * 25)
  Dim u As Integer
  For u = 1 To CPUWidth Step CPUProz * 25 'Vertikale Linien in CPUProz*25 -er Abständen
     Call MoveToEx(CPUHdc, u + 2, 1, Buffer)
     Call LineTo(CPUHdc, u + 2, CPUHeight)
  Next u
  Counted = 9
  CPUWidth = CPUWidth - 2
End Sub
Sub ChangeColor(ByVal ContHdc As Long, ByVal Color As Long)
  DeleteObject (SelectObject(ContHdc, CreatePen(0&, 1&, Color)))
End Sub
Private Sub Timer1_Timer()
  If Invisible = True Then Exit Sub
  If Exitting Then Exit Sub
  Dim Buffer As POINTAPI, sPro As Byte
  BitBlt CPUHdc, 0&, 0&, CPUWidth, CPUHeight + 1, CPUHdc, 2&, 0&, vbSrcCopy
  If Counted > CPUProz * 12 - 1 Then
    Counted = 0
    ChangeColor CPUHdc, RGB(0, 80, 0)
    Call MoveToEx(CPUHdc, CPUWidth - 4, 1, Buffer)
    Call LineTo(CPUHdc, CPUWidth - 4, CPUHeight)
  Else
    Counted = Counted + 1
  End If
  sPro = aPro
  If TMode = 1 Then
    CurrSessionBytesSent = SessionBytesSent - oldSessionBytesSent
    aPro = (CurrSessionBytesSent / (MaxByteOut / 100)) 'Int(Rnd(100) * 100) + 1
    oldSessionBytesSent = SessionBytesSent
  Else
    CurrSessionBytesReceived = SessionBytesReceived - oldSessionBytesReceived
    aPro = (CurrSessionBytesReceived / (MaxByteOut / 100)) 'Int(Rnd(100) * 100) + 1
    oldSessionBytesReceived = SessionBytesReceived
  End If
    Linie CPUWidth - 5, CPUHeight - CPUProz * sPro, CPUWidth - 3, CPUHeight - CPUProz * aPro
  Traffic.Refresh
  lblCPU = CStr(aPro) + "%"
End Sub
Sub Linie(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
  Dim tmp As Integer, xincr As Integer, yincr As Integer, dx As Integer
  Dim dy As Integer, d As Integer, aincr As Integer, bincr As Integer
  Dim X As Long, Y As Long
  If Abs(x2 - x1) < Abs(y2 - y1) Then
    If y1 > y2 Then tmp = x1: x1 = x2: x2 = tmp: tmp = y1: y1 = y2: y2 = tmp
    If x2 > x1 Then xincr = 1 Else xincr = -1
    dx = Abs(x2 - x1)
    dy = y2 - y1
    d = 2 * dx - dy
    aincr = 2 * (dx - dy)
    bincr = 2 * dx
    X = x1
    Y = y1
    SetPixelV CPUHdc, X, Y, GetCol(Y)
    For Y = y1 + 1 To y2
      If d >= 0 Then X = X + xincr: d = d + aincr Else d = d + bincr
      SetPixelV CPUHdc, X, Y, GetCol(Y)
    Next Y
  Else
    If x1 > x2 Then tmp = x1: x1 = x2: x2 = tmp: tmp = y1: y1 = y2: y2 = tmp
    If y2 > y1 Then yincr = 1 Else yincr = -1
    dx = x2 - x1
    dy = Abs(y2 - y1)
    d = 2 * dy - dx
    aincr = 2 * (dy - dx)
    bincr = 2 * dy
    X = x1
    Y = y1
    SetPixelV CPUHdc, X, Y, GetCol(Y)
    For X = x1 + 1 To x2
      If d >= 0 Then Y = Y + yincr: d = d + aincr Else d = d + bincr
      SetPixelV CPUHdc, X, Y, GetCol(Y)
    Next X
  End If
End Sub
Private Sub UserControl_Initialize()
  CPUHdc = Traffic.hdc
  CPUWidth = Traffic.Width / 15
  CPUHeight = Traffic.Height / 15
  InitializeCPUGraph
End Sub
