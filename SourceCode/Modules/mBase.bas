Attribute VB_Name = "mBase"
Option Explicit

Public oAddressList As cAddresslist

Sub fInitialize()
  Set oAddressList = New cAddresslist
End Sub

Sub fTerminate()
  Set oAddressList = Nothing
End Sub
