VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddresslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const myStep As Long = 10
Private myItem() As cAddresslist_Item
Private myCount As Long
Private myUsed  As Long


Private Sub Class_Initialize()
  ReDim myItem(myStep)
  myCount = 0
  myUsed = 0
End Sub


Private Sub Class_Terminate()
  Dim Index As Long
  For Index = 1 To myCount
    Set myItem(Index) = Nothing
  Next Index
End Sub


Public Property Get oItem(Index As Long) As cAddresslist_Item
  If Index > 0 Then
    If Not Index > myCount Then
      Set oItem = myItem(Index)
    End If
  End If
End Property


Public Property Get pCount() As Long
  pCount = myCount
End Property


Public Property Get pMemoryUsed() As Long
  Dim Index As Long
  Dim Dummy As Long
  Dummy = ((1 + UBound(myItem)) * 4) + 8
  For Index = 1 To myCount
    If Not myItem(Index) Is Nothing Then
      Dummy = Dummy + myItem(Index).pMemoryUsed
    End If
  Next Index
  pMemoryUsed = Dummy
End Property


Public Function fReset() As Long
  Dim Index As Long
  fReset = -1
  For Index = 1 To myCount
    Set myItem(Index) = Nothing
  Next Index
  ReDim myItem(myStep)
  myCount = 0
  myUsed = 0
  fReset = 0
End Function


Public Function fAdd() As Long
  Dim Index As Long
  fAdd = -1
  If myCount > myUsed Then
    For Index = 1 To myCount
      If myItem(Index) Is Nothing Then
        Set myItem(Index) = New cAddresslist_Item
        myUsed = myUsed + 1
        fAdd = Index
        Exit Function
      End If
    Next
  Else
    myCount = myCount + 1
    myUsed = myCount
    If myCount >= UBound(myItem) Then ReDim Preserve myItem((((myCount - (myCount Mod myStep)) / myStep) + 1) * myStep)
    Set myItem(myCount) = New cAddresslist_Item
    fAdd = myCount
  End If
End Function


Public Function fRemove(Index As Long) As Long
  fRemove = -1
  If Index > 0 Then
    If Not Index > myCount Then
      Set myItem(Index) = Nothing
      myUsed = myUsed - 1
      If Index = myCount Then
        Dim Index2 As Long
        For Index2 = myCount To 1 Step -1
          If myItem(Index2) Is Nothing Then
            myCount = myCount = -1
          Else
            Exit For
          End If
        Next
      End If
      If UBound(myItem) - myCount > myStep Then ReDim Preserve myItem((((myCount - (myCount Mod myStep)) / myStep) + 1) * myStep)
      fRemove = 0
    End If
  End If
End Function


Public Function fFind(Nick As String) As Long
  Dim Index As Long
  fFind = -1
  For Index = 1 To myCount
    If Not myItem(Index) Is Nothing Then
      If myItem(Index).pNick = Nick Then
        fFind = Index
        Exit Function
      End If
    End If
  Next
End Function


