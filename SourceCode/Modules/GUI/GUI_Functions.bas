Attribute VB_Name = "GUI_Functions"
Option Explicit

Public Sub Output(ByVal What As String) 'NoTrap' : AddStack "mdlWinsock_Output(" & What & ")"
    If Invisible = True Then Exit Sub
    If Exitting Then Exit Sub
    If GUI_frmWinsock.WindowState <> 0 Then Exit Sub
    If ParamXCount(GUI_frmWinsock.txtWinsock.Text, vbCrLf) > 25 Then
      GUI_frmWinsock.txtWinsock.Text = GetRestX(GUI_frmWinsock.txtWinsock.Text, vbCrLf, 2) & What
    Else
      GUI_frmWinsock.txtWinsock.Text = GUI_frmWinsock.txtWinsock.Text & What
    End If
    GUI_frmWinsock.txtWinsock.SelStart = Len(GUI_frmWinsock.txtWinsock)
    GUI_frmWinsock.txtWinsock.SelLength = 0
End Sub

Public Sub Status(ByVal What As String) 'NoTrap' : AddStack "mdlWinsock_Status(" & What & ")"
    If Invisible = True Then Exit Sub
    If Exitting Then Exit Sub
    If ParamXCount(GUI_frmWinsock.txtStatus.Text, vbCrLf) > 10 Then
      GUI_frmWinsock.txtStatus.Text = GetRestX(GUI_frmWinsock.txtStatus.Text, vbCrLf, 2) & What
    Else
      GUI_frmWinsock.txtStatus.Text = GUI_frmWinsock.txtStatus.Text & What
    End If
    GUI_frmWinsock.txtStatus.SelStart = Len(GUI_frmWinsock.txtStatus.Text)
    GUI_frmWinsock.txtStatus.SelLength = 0
End Sub


