VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeaderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#package vb.com.rgsoftware.io

Option Explicit

Private Header() As Variant

Public Sub setColumnCount(Count As Integer)
    ReDim Header(1 To Count, 1 To 2) As Variant
End Sub

Public Function setColumnName(Column As Integer, Value As Variant)
    Header(Column, 1) = Value
End Function

Public Function getColumnName(Column As Integer) As Variant
    getColumnName = Header(Column, 1)
End Function

Public Function setColumnInfo(Column As Integer, Value As Variant)
    On Error Resume Next
    Header(Column, 2) = Value
End Function

Public Function getColumnInfo(Column As Integer) As Variant
    getColumnInfo = Header(Column, 2)
End Function

Public Property Get ColumnCount() As Integer
    On Error Resume Next
    ColumnCount = UBound(Header)
End Property
