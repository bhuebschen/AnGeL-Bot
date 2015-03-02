Attribute VB_Name = "GUI_TrayIcon"
',-======================- ==-- -  -
'|   AnGeL - GUI - TrayIcon
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit

'Tray icon states
Public Const SI_Offline As Byte = 1
Public Const SI_Connecting As Byte = 2
Public Const SI_Online As Byte = 3
Public Const SI_Hub As Byte = 4

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIF_MESSAGE = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private TrayIcon As NOTIFYICONDATA
Public TrayIconShown As Boolean, MSG_TaskbarCreated As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Adds the bot's tray icon
Public Sub AddTrayIcon() ' : AddStack "Windows_AddTrayIcon()"
  If Not Invisible Then
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = GUI_frmWinsock.hwnd
    TrayIcon.uId = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    If TrayIconShown = False Then
      If HubBot = True Then
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateHub(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateHub(0).Picture
        End If
        TrayIcon.szTip = IIf(BotNetNick <> "", BotNetNick, PrimaryNick) & " - Hub (no servers)" & Chr(0)
      Else
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateRed(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateRed(0).Picture
        End If
        TrayIcon.szTip = IIf(BotNetNick <> "", BotNetNick, PrimaryNick) & " - Not Connected" & Chr(0)
      End If
      Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
      PutLog "|  Added tray icon."
    Else
      Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
      PutLog "|  Added tray icon due to taskbar creation."
    End If
    TrayIconShown = True
  End If
End Sub

'Changes the tray icon
Public Sub SetTrayIcon(State As Byte) ' : AddStack "Windows_SetTrayIcon(" & State & ")"
Dim ShownBotNick As String
  If Not Invisible Then
    ShownBotNick = IIf(BotNetNick <> "", BotNetNick, PrimaryNick)
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP
    Select Case State
      Case SI_Offline
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateRed(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateRed(0).Picture
        End If
          TrayIcon.szTip = ShownBotNick & " - Not Connected" & Chr(0)
      Case SI_Connecting
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateYellow(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateYellow(0).Picture
        End If
          TrayIcon.szTip = ShownBotNick & " - Connecting..." & Chr(0)
      Case SI_Online
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateGreen(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateGreen(0).Picture
        End If
          TrayIcon.szTip = ShownBotNick & " - Connected" & Chr(0)
      Case SI_Hub
        If GUI_frmWinsock.HideThings.Visible = True Then
          TrayIcon.hIcon = GUI_frmWinsock.imgStateHub(1).Picture
        Else
          TrayIcon.hIcon = GUI_frmWinsock.imgStateHub(0).Picture
        End If
          TrayIcon.szTip = ShownBotNick & " - Hub (no servers)" & Chr(0)
    End Select
    Call Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
  End If
End Sub

Public Sub RemTrayIcon() ' : AddStack "Windows_RemTrayIcon()"
  If Not Invisible Then
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = GUI_frmWinsock.hwnd
    TrayIcon.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    PutLog "|  Removed tray icon."
    TrayIconShown = False
  End If
End Sub

