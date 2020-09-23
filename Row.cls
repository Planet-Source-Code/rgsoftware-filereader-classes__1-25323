VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Row"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#package vb.com.rgsoftware.io

Option Explicit

Private Col() As Variant

Public Sub setColumnCount(Count As Integer)
    ReDim Col(1 To Count) As Variant
End Sub

Public Sub setData(Column As Integer, Value As Variant)
    On Error Resume Next
    Col(Column) = Value
End Sub

Public Function getData(Column As Integer) As Variant
    On Error Resume Next
    getData = Col(Column)
End Function
