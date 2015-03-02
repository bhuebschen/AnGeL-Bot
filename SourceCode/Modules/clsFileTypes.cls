VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -——————————————————- -- -  -
'|   AnGeL FileType Class
'| © 2003 by Benedikt Hübschen
' -—————————- -- -  -
'|
' -——————-————- -——— -- -  -

Private Type PolicyFile
  PolName As String * 64
  PolReserved1 As String * 4
  PolReserved2 As String * 4
  PolValue As Byte
  PolReserved3 As String * 4
End Type

Private Policies() As PolicyFile

Private Sub Class_Initialize()
  ReDim Policies(0)
End Sub

' -——————————————————- -- -  -
'|Begin of ScriptPolicies
    Public Function CommandAllowed(Name As String) As Boolean
      If UsePolicies = False Then
        CommandAllowed = True
        Exit Function
      End If
      For i = 0 To UBound(Policies)
        If LCase(Name) = Trim(LCase(Policies(i).PolName)) Then
          CommandAllowed = IIf(Policies(i).PolValue = 1, True, False)
          Exit Function
        End If
      Next i
      CommandAllowed = False
    End Function
'|
    Public Sub AllowCommand(Name As String, setit As Boolean)
      For i = 0 To UBound(Policies)
        If LCase(Name) = Trim(LCase(Policies(i).PolName)) Then
          If setit = True Then
          Policies(i).PolValue = 1
          Else
          Policies(i).PolValue = 0
          End If
          Exit Sub
        End If
      Next i
    End Sub
'|
    Public Function UsePolicies() As Boolean
      If Dir(HomeDir + "policies.dat") <> "" Then
        UsePolicies = True
      Else
        UsePolicies = False
      End If
    End Function
'|
    Public Sub LoadPolicies()
      Dim FileNumber As Long
      
      Dim Header As String * 29
      Dim PolCount As Long
      Dim Footer As String * 53
      
      If UsePolicies Then
        FileNumber = FreeFile
        Open HomeDir + "policies.dat" For Binary As #FileNumber
          Get #FileNumber, , Header
          Get #FileNumber, , PolCount
          ReDim Policies(PolCount)
          Get #FileNumber, , Policies
          Get #FileNumber, , Footer
        Close #FileNumber
      End If
    End Sub
'|
    Sub WritePolicies()
      Dim FileNumber As Long
      
      Dim Header As String * 29
      Dim PolCount As Long
      Dim Footer As String * 53
      
      Header = DecryptString("D1CE1785EC80A0EFFCE99329658066691C55", "aNgEl")
      Footer = DecryptString("1C105B8FD6BF68AE3BD13F3D1F152A5FBA57B2B67230A21F32BC88D58B5BDF607D77E7BF2C17EC462950606F4293A8901BCE1080CC", "AnGeL")
      
        FileNumber = FreeFile
        Open HomeDir & "policies.dat" For Binary As #FileNumber
          Put #FileNumber, , Header
          Put #FileNumber, , UBound(Policies)
          Put #FileNumber, , Policies
          Put #FileNumber, , Footer
        Close #FileNumber
        If NTFS_IsNTFS(HomeDir & "policies.dat") Then
          'NTFS_Encrypt HomeDir & "policies.dat"
        End If
    End Sub
'|
    Sub CreatePolicies()
      ReDim Policies(15)
      Policies(0).PolName = "UserManagement": Policies(0).PolValue = 1: Policies(0).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(0).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 138 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(0).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(77) * 13 + 1))
      Policies(1).PolName = "Resolves": Policies(1).PolValue = 1: Policies(1).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(1).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 76 + 1)): Policies(1).PolReserved3 = Chr(Int(Rnd(110) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(2).PolName = "Database": Policies(2).PolValue = 1: Policies(2).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(2).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 130 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(2).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(3).PolName = "Objects": Policies(3).PolValue = 0: Policies(3).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(3).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(210) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(3).PolReserved3 = Chr(Int(Rnd(20) * 67 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(4).PolName = "FileOperations": Policies(4).PolValue = 1: Policies(4).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(4).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(4).PolReserved3 = Chr(Int(Rnd(90) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 44 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(5).PolName = "ExecuteCommand": Policies(5).PolValue = 1: Policies(5).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(5).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(150) * 13 + 1)) + Chr(Int(Rnd(20) * 66 + 1)): Policies(5).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(99) * 13 + 1))
      Policies(6).PolName = "TimeOperations": Policies(6).PolValue = 1: Policies(6).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(6).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(6).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 22 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(7).PolName = "SocketOperations": Policies(7).PolValue = 1: Policies(7).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(7).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(7).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(200) * 13 + 1))
      Policies(8).PolName = "INIoperations": Policies(8).PolValue = 1: Policies(8).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(8).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 23 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(8).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(9).PolName = "NotificationHooks": Policies(9).PolValue = 1: Policies(9).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 53 + 1)) + Chr(Int(Rnd(70) * 13 + 1)): Policies(9).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(10) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(9).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(10).PolName = "WMI": Policies(10).PolValue = 0: Policies(10).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(10).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(10).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1))
      Policies(11).PolName = "NativeFS": Policies(11).PolValue = 0: Policies(11).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(11).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 143 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(11).PolReserved3 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(210) * 13 + 1))
      Policies(12).PolName = "AddInModules": Policies(12).PolValue = 0: Policies(12).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(12).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(12).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      Policies(13).PolName = "BotInteractions": Policies(13).PolValue = 1: Policies(13).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(13).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(13).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      Policies(14).PolName = "ChannelInteractions": Policies(14).PolValue = 1: Policies(14).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(14).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(14).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      Policies(15).PolName = "SessionChange": Policies(15).PolValue = 1: Policies(15).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(15).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(15).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      'Policies(16).PolName = "AddInModules": Policies(16).PolValue = 0: Policies(16).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(16).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(16).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      'Policies(17).PolName = "AddInModules": Policies(17).PolValue = 0: Policies(17).PolReserved1 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 33 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(17).PolReserved2 = Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)): Policies(17).PolReserved3 = Chr(Int(Rnd(20) * 55 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(20) * 13 + 1)) + Chr(Int(Rnd(120) * 13 + 1))
      WritePolicies
    End Sub
'|End Of ScriptPolicies
' -——————————————————- -- -  -

