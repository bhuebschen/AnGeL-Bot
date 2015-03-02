VERSION 5.00
Begin VB.Form GUI_frmBotSetup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Setup - AnGeL"
   ClientHeight    =   5385
   ClientLeft      =   4875
   ClientTop       =   3585
   ClientWidth     =   7425
   Icon            =   "GUI_frmBotSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7425
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command2 
      Caption         =   "<-- Back"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4845
      TabIndex        =   6
      Top             =   4890
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next -->"
      Height          =   345
      Left            =   6120
      TabIndex        =   5
      Top             =   4890
      Width           =   1125
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   2460
      ScaleHeight     =   1095
      ScaleWidth      =   4965
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   90
      Width           =   4965
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Explanations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   195
         TabIndex        =   30
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Explanations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   210
         TabIndex        =   31
         Top             =   15
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008080&
         FillStyle       =   0  'Ausgefüllt
         Height          =   105
         Left            =   60
         Top             =   165
         Width           =   4815
      End
      Begin VB.Label Explanation 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   4605
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   765
         Left            =   60
         Top             =   255
         Width           =   4815
      End
   End
   Begin VB.PictureBox FirstPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3285
      Left            =   2460
      ScaleHeight     =   3285
      ScaleWidth      =   4965
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1400
      Width           =   4965
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IRC Servers"
         ForeColor       =   &H00000000&
         Height          =   1245
         Left            =   60
         TabIndex        =   19
         Top             =   1920
         Width           =   4815
         Begin VB.TextBox Text5 
            ForeColor       =   &H00000000&
            Height          =   795
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "GUI_frmBotSetup.frx":014A
            Top             =   270
            Width           =   4515
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Identification"
         Height          =   1815
         Left            =   60
         TabIndex        =   14
         Top             =   45
         Width           =   4815
         Begin VB.TextBox Text1 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   0
            Top             =   270
            Width           =   1965
         End
         Begin VB.TextBox Text2 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   1
            Top             =   630
            Width           =   1965
         End
         Begin VB.TextBox Text3 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   2
            Top             =   990
            Width           =   1965
         End
         Begin VB.TextBox Text4 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   3
            Text            =   "Frag doch :)"
            Top             =   1350
            Width           =   3165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Primary nick:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Secondary nick:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ident:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Real name:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   1410
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox SecondPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3315
      Left            =   2460
      ScaleHeight     =   3315
      ScaleWidth      =   4965
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1400
      Visible         =   0   'False
      Width           =   4965
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First time super owner identification keyword"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   90
         TabIndex        =   32
         Top             =   60
         Width           =   4815
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1500
            TabIndex        =   7
            Text            =   "hiya"
            Top             =   270
            Width           =   3165
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "/msg <your bot>"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Messages"
         ForeColor       =   &H00000000&
         Height          =   1125
         Left            =   90
         TabIndex        =   25
         Top             =   870
         Width           =   4815
         Begin VB.TextBox Text11 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   9
            Text            =   "Angel leaving... See ya!"
            Top             =   660
            Width           =   3165
         End
         Begin VB.TextBox Text10 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1500
            TabIndex        =   8
            Text            =   "YourNick <your@email.com>"
            Top             =   270
            Width           =   3165
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quit message:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bot admin info:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   330
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advanced settings"
         Height          =   1125
         Left            =   90
         TabIndex        =   21
         Top             =   2070
         Width           =   4815
         Begin VB.TextBox Text9 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1140
            TabIndex        =   12
            Top             =   660
            Width           =   3525
         End
         Begin VB.TextBox Text6 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1140
            TabIndex        =   10
            Text            =   "21"
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox Text7 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3540
            TabIndex        =   11
            Text            =   "3333"
            Top             =   270
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "LocalHost:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Telnet port:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   330
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Botnet port:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2520
            TabIndex        =   22
            Top             =   330
            Width           =   825
         End
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000F&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Ausgefüllt
      Height          =   645
      Left            =   0
      Top             =   4740
      Width           =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   7425
      Y1              =   4725
      Y2              =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      Index           =   0
      X1              =   0
      X2              =   7425
      Y1              =   4710
      Y2              =   4710
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   0
      Picture         =   "GUI_frmBotSetup.frx":0192
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "GUI_frmBotSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SettingFocus As Boolean, AlreadyActivated As Boolean

Private Sub Command1_Click() ' : AddStack "BotSetup_Command1_Click()"
Dim TheText As String, TheParam As String, u As Long
Dim TheMsg As String
  If FirstPage.Visible Then
    TheMsg = ""
    If TheMsg <> "" Then
      SettingFocus = True
      Text2.SetFocus
      MsgBox TheMsg, vbCritical, "Error"
      SettingFocus = False
      Exit Sub
    End If
    TheMsg = ""
    If Text3.Text = "" Then TheMsg = "Please enter an ident!"
    If TheMsg <> "" Then
      SettingFocus = True
      Text3.SetFocus
      MsgBox TheMsg, vbCritical, "Error"
      SettingFocus = False
      Exit Sub
    End If
    
    SecondPage.Visible = True
    FirstPage.Visible = False
    Command2.Enabled = True
    Text8.SetFocus
  Else
    WritePPString "Identification", "PrimaryNick", Text1.Text, AnGeL_INI
    WritePPString "Identification", "SecondaryNick", Text2.Text, AnGeL_INI
    WritePPString "Identification", "Identd", Text3.Text, AnGeL_INI
    WritePPString "Identification", "RealName", Text4.Text, AnGeL_INI
    WritePPString "Identification", "Admin", Text10.Text, AnGeL_INI
    If Text9.Text <> "" Then WritePPString "Identification", "LocalHost", Text9.Text, AnGeL_INI
    WritePPString "Identification", "FirstTimeKeyword", Text8.Text, AnGeL_INI
    
    WritePPString "Server", "QuitMessage", Text11.Text, AnGeL_INI
    TheText = Replace(Text5.Text, vbCrLf, "§")
    
    While InStr(TheText, "§§")
      TheText = Replace(TheText, "§§", "§")
    Wend
    
    If Left(TheText, 1) = "§" Then TheText = Mid(TheText, 2)
    If Right(TheText, 1) = "§" Then TheText = Left(TheText, Len(TheText) - 1)
    
    For u = 1 To ParamXCount(TheText, "§")
      TheParam = ParamX(TheText, "§", u)
      If u = 1 Then TheMsg = "Server" Else TheMsg = "Server" & Trim(Str(u))
      WritePPString "Server", TheMsg, TheParam, AnGeL_INI
    Next u
    
    WritePPString "TelNet", "UserPort", Text6.Text, AnGeL_INI
    WritePPString "TelNet", "Port", Text7.Text, AnGeL_INI
    
    If Text9.Text <> "" Then WritePPString "Others", "ResolveIP", "2", AnGeL_INI
    
    Clipboard.SetText "/msg " & Text1.Text & " " & Text8.Text
    MsgBox "You're done. After pressing OK, your bot will start and (hopefully) get on IRC." & vbCrLf & "Please send the identification message then (it has been copied to your clipboard):" & vbCrLf + vbCrLf & "/msg " & Text1.Text & " " & Text8.Text, vbInformation, "AnGeL Bot Setup complete!"
    Unload Me
  End If
End Sub

Private Sub Command2_Click() ' : AddStack "BotSetup_Command2_Click()"
  FirstPage.Visible = True
  SecondPage.Visible = False
  Command2.Enabled = False
End Sub

Private Sub Form_Activate() ' : AddStack "BotSetup_Form_Activate()"
  If AlreadyActivated = False Then
    SettingFocus = True
    Text1.SetFocus
    SettingFocus = False
    AlreadyActivated = True
  End If
End Sub

Private Sub Text1_GotFocus() ' : AddStack "BotSetup_Text1_GotFocus()"
  Explanation.Caption = "Please enter your desired primary bot nick (the nick that your bot will always try to get) into the first textbox below." & vbCrLf & "Press TAB to jump to the next field then."
End Sub

Private Sub Text1_LostFocus() ' : AddStack "BotSetup_Text1_LostFocus()"
Dim ErrMsg As String
  If SettingFocus Then Exit Sub
  'If Len(Text1.Text) > ServerNickLen Then ErrMsg = "A nickname can't be longer than " & Trim(Str(ServerNickLen)) & " characters!"
  If IsValidNick(Text1.Text) = False Then ErrMsg = "The nickname you entered is not valid! Try using characters only."
  If Text1.Text = "" Then ErrMsg = "Please enter a primary nick!"
  If ErrMsg <> "" Then
    SettingFocus = True
    Text1.SetFocus
    MsgBox ErrMsg, vbCritical, "Error"
    SettingFocus = False
  End If
End Sub

Private Sub Text2_GotFocus() ' : AddStack "BotSetup_Text2_GotFocus()"
  Explanation.Caption = "If the bot's primary nick is being used by somebody else, the bot will use the secondary nick specified here. For example, if your primary nick is 'Angie', you could enter 'Angie-' here."
End Sub

Private Sub Text2_LostFocus() ' : AddStack "BotSetup_Text2_LostFocus()"
Dim ErrMsg As String
  If SettingFocus Then Exit Sub
  'If Len(Text2.Text) > ServerNickLen Then ErrMsg = "A nickname can't be longer than " & Trim(Str(ServerNickLen)) & " characters!"
  If IsValidNick(Text2.Text) = False Then ErrMsg = "The nickname you entered is not valid! Try using characters only."
  If Text2.Text = "" Then ErrMsg = "Please enter a secondary nick!"
  If Text2.Text = Text1.Text Then ErrMsg = "The primary and the secondary nick have to be different!"
  If ErrMsg <> "" Then
    SettingFocus = True
    Text2.SetFocus
    MsgBox ErrMsg, vbCritical, "Error"
    SettingFocus = False
  End If
End Sub

Private Sub Text3_GotFocus() ' : AddStack "BotSetup_Text3_GotFocus()"
  Explanation.Caption = "An IRC hostmask looks like this: nick!ident@host.domain" & vbCrLf & "You can specify the 'ident' part here, the so-called 'username' (the part right after the bot nick)."
End Sub

Private Sub Text3_LostFocus() ' : AddStack "BotSetup_Text3_LostFocus()"
Dim ErrMsg As String
  If SettingFocus Then Exit Sub
  If Text3.Text = "" Then ErrMsg = "Please enter an ident!"
  If ErrMsg <> "" Then
    SettingFocus = True
    Text3.SetFocus
    MsgBox ErrMsg, vbCritical, "Error"
    SettingFocus = False
  End If
End Sub

Private Sub Text4_GotFocus() ' : AddStack "BotSetup_Text4_GotFocus()"
  Explanation.Caption = "When somebody does a /whois on your bot, he'll see the text you can specify here. It's supposed to be the real name, but mostly it's used for some cool message or whatever ;-)"
End Sub

Private Sub Text5_GotFocus() ' : AddStack "BotSetup_Text5_GotFocus()"
  Explanation.Caption = "Please enter a list of IRC servers here. The bot will try to connect to these servers in the order you specified them. Please ensure that all IRC servers are valid and that they allow bots!"
End Sub

Private Sub Text5_LostFocus() ' : AddStack "BotSetup_Text5_LostFocus()"
Dim TheText As String, TheParam As String, u As Long
  If SettingFocus Then Exit Sub
  TheText = Replace(Text5.Text, vbCrLf, "§")
  While InStr(TheText, "§§") > 0
    TheText = Replace(TheText, "§§", "§")
  Wend
  If Right(TheText, 1) = "§" Then TheText = Left(TheText, Len(TheText) - 1)
  For u = 1 To ParamXCount(TheText, "§")
    TheParam = ParamX(TheText, "§", u)
    If InStr(TheParam, ":") = 0 Then
      SettingFocus = True
      Text5.SetFocus
      MsgBox "Every IRC server entry has to look like this: irc.server.com:6667" & vbCrLf & "The following line is not correct:" & vbCrLf + vbCrLf + TheParam, vbCritical, "Error"
      SettingFocus = False
      Exit Sub
    Else
      If IsNumeric(ParamX(TheParam, ":", 2)) = False Then
        SettingFocus = True
        Text5.SetFocus
        MsgBox "The IRC server port has to be a number!" & vbCrLf & "The following line is not correct:" & vbCrLf + vbCrLf + TheParam, vbCritical, "Error"
        SettingFocus = False
        Exit Sub
      End If
      If InStr(ParamX(TheParam, ":", 1), " ") > 0 Then
        SettingFocus = True
        Text5.SetFocus
        MsgBox "The IRC server name can't contain spaces!" & vbCrLf & "The following line is not correct:" & vbCrLf + vbCrLf + TheParam, vbCritical, "Error"
        SettingFocus = False
        Exit Sub
      End If
    End If
  Next u
End Sub

Private Sub Text8_GotFocus() ' : AddStack "BotSetup_Text8_GotFocus()"
  Explanation.Caption = "The first time your bot is online, you'll have to send it a private message with a 'keyword' - you will be added as super owner then. This keyword can be specified here."
End Sub

Private Sub Text8_LostFocus() ' : AddStack "BotSetup_Text8_LostFocus()"
Dim ErrMsg As String
  If SettingFocus Then Exit Sub
  If Trim(Text8.Text) = "" Then ErrMsg = "You must enter an identification keyword!"
  If InStr(Text8.Text, " ") > 0 Then ErrMsg = "Please don't use spaces in your identification keyword!"
  If ErrMsg <> "" Then
    SettingFocus = True
    Text8.SetFocus
    MsgBox ErrMsg, vbCritical, "Error"
    SettingFocus = False
  End If
End Sub

Private Sub Text10_GotFocus() ' : AddStack "BotSetup_Text10_GotFocus()"
  Explanation.Caption = "This line will be shown if somebody does a '.who <botnick>' in the botnet. You should enter your real nick and e-mail address here so that you can be contacted when there are problems."
End Sub

Private Sub Text11_GotFocus() ' : AddStack "BotSetup_Text11_GotFocus()"
  Explanation.Caption = "This is the message that your bot will show on IRC when it is shutting down (because of the '.die' command or because the bot window is being closed)."
End Sub

Private Sub Text6_GotFocus() ' : AddStack "BotSetup_Text6_GotFocus()"
  Explanation.Caption = "If your bot is not on IRC, you can still connect to it via telnet. The telnet port your bot will use can be specified here. Leave this setting unchanged if you don't understand what it does."
End Sub

Private Sub Text7_GotFocus() ' : AddStack "BotSetup_Text7_GotFocus()"
  Explanation.Caption = "The botnet port is necessary to allow other bots to connect to your bot. It must be different from the telnet port. Leave this setting unchanged if you don't understand what it does."
End Sub

Private Sub Text9_GotFocus() ' : AddStack "BotSetup_Text9_GotFocus()"
  Explanation.Caption = "If your machine has more than one IP and different hosts, you can specify the local host you want to use here. You MUST leave this setting empty if you don't understand what it does!"
End Sub
