Attribute VB_Name = "TimedEvents_Orders"
',-======================- ==-- -  -
'|   AnGeL - TimedEvents - Orders
'|   © 1998-2003 by the AnGeL-Team
'|-=============- -==- ==- -- -
'|
'|  Last Changed: 31.05.2003 - (SailorSat) Ariane Fugmann
'|
'`-=====================-- -===- ==- -- -
Option Explicit


Private Type TOrder
  DoThis As String
  AtTime As Currency
End Type

Public OrderCount As Long
Public Orders() As TOrder


Sub Orders_Load()
  ReDim Preserve Orders(5)
End Sub

Sub Orders_Unload()
'
End Sub


Public Sub Order(DoThis As String, HowLong As Long)
  OrderCount = OrderCount + 1
  If OrderCount > UBound(Orders()) Then ReDim Preserve Orders(UBound(Orders()) + 5)
  Orders(OrderCount).DoThis = DoThis
  Orders(OrderCount).AtTime = WinTickCount + HowLong * 1000
End Sub


Public Function IsOrdered(Name As String) As Boolean
  Dim Index As Long
  For Index = 1 To OrderCount
    If Index > UBound(Orders()) Then Exit For
    If LCase(Orders(Index).DoThis) = LCase(Name) Then IsOrdered = True: Exit Function
  Next Index
  IsOrdered = False
End Function


Public Sub RemOrder(Name As String)
  Dim Index As Long, RemovedOne As Boolean
  RemovedOne = True
  While RemovedOne = True
    RemovedOne = False
    For Index = 1 To OrderCount
      If Index > UBound(Orders()) Then Exit For
      If LCase(Orders(Index).DoThis) = LCase(Name) Then RemoveOrder Index: RemovedOne = True: Exit For
    Next Index
  Wend
End Sub


Public Sub RemoveOrder(Num As Long)
  Dim Index As Long
  For Index = Num To OrderCount - 1
    Orders(Index) = Orders(Index + 1)
  Next Index
  OrderCount = OrderCount - 1
  Index = ((OrderCount \ 50) + 1) * 50
  If Index < UBound(Orders()) Then ReDim Preserve Orders(Index)
End Sub

