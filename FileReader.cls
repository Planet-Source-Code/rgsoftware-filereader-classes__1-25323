VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#package vb.com.rgsoftware.io

'Name: FileReader.cls (uses Row.cls)
'Purpose: Reads a delimited seperated value file into Row objects
'Author: Richard Gardner 7/21/01
'Copyleft 2001 by RG Software Corporation under GNU public license
''http://www.rgsoftware.com

Option Explicit
Option Compare Text 'For keys

Private Header As New HeaderInfo
Private Data() As New Row
Private m_RowCount As Long
Private m_ColumnCount As Integer

Public Function readFile(FileName As String, Delimiter As String, _
            FirstRowNames As Boolean, Optional Qualifier As String, _
            Optional StopRow As Long)
    Dim Column As Integer
    Dim Found As Long
    Dim nFile As Long
    Dim Count As Long
    Dim ASCII As String
    Dim Value As String
    Dim Del As String
    On Error GoTo FileError
    RowCount = 0
    ColumnCount = 0
    nFile = FreeFile
    'Get column count
    Open FileName For Input As #nFile
    Line Input #nFile, ASCII
    Close #nFile
    Found = -1
    If Found = 0 Then Exit Function
    Do While Found <> 0
        Found = InStr(ASCII, Delimiter)
        If Found <> 0 Then
            ASCII = Trim$(Mid$(ASCII, Found + 1))
        End If
        ColumnCount = ColumnCount + 1
    Loop
    Header.setColumnCount ColumnCount
    'Check for column names
    If FirstRowNames Then
        Open FileName For Input As #nFile
        Line Input #nFile, ASCII
        Close #nFile
        Column = 0
        Found = -1
        Do While Found <> 0
            Found = InStr(ASCII, Delimiter)
            If Found <> 0 Then
                Value = Mid$(ASCII, 1, Found - 1)
                ASCII = Trim$(Mid$(ASCII, Found + 1))
            Else
                Value = Trim$(ASCII)
            End If
            Column = Column + 1
            If Qualifier <> "" Then
                If InStr(Value, Qualifier) <> 0 Then
                    Value = Replace(Value, Qualifier, "")
                End If
            End If
            Header.setColumnName CInt(Column), Value
        Loop
    End If
    'Get row count
    Open FileName For Input As #nFile
    If FirstRowNames Then Line Input #nFile, ASCII
    Do While Not EOF(nFile)
        Line Input #nFile, ASCII
        RowCount = RowCount + 1
    Loop
    Close #nFile
    'Resize data array
    If StopRow <> 0 Then RowCount = StopRow
    ReDim Data(1 To RowCount) As New Row
    Count = 0
    'Read file into memory
    Open FileName For Input As #nFile
    If FirstRowNames Then Line Input #nFile, ASCII
    Do While Not EOF(nFile)
        Line Input #nFile, ASCII
        Count = Count + 1
        Data(Count).setColumnCount ColumnCount
        Column = 0
        Found = -1
        Do While Found <> 0
            Found = InStr(ASCII, Delimiter)
            If Found <> 0 Then
                Value = Mid$(ASCII, 1, Found - 1)
                ASCII = Trim$(Mid$(ASCII, Found + 1))
            Else
                Value = Trim$(ASCII)
            End If
            Column = Column + 1
            If Qualifier <> "" Then
                If InStr(Value, Qualifier) <> 0 Then
                    Value = Replace(Value, Qualifier, "")
                End If
            End If

            Data(Count).setData Column, Value
        Loop
        If StopRow <> 0 And Count = StopRow Then Exit Do
    Loop
    Close #nFile
    readFile = 1
    Exit Function
FileError:
    Err.Raise Err.Number, "readFile", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function getHeaderObject()
    Set getHeaderObject = Header
End Function

Public Sub setData(Key As Variant, Row As Long, Value As Variant)
    Data(Row).setData GetCol(Key), Value
End Sub

Public Function getData(Key As Variant, Row As Long) As Variant
    getData = Data(Row).getData(GetCol(Key))
End Function

Public Sub setColumnName(Col As Integer, Value As Variant)
    Header.setColumnName Col, Value
End Sub

Public Function getColumnName(Col As Integer) As Variant
    getColumnName = Header.getColumnName(Col)
End Function

Public Sub setColumnInfo(Col As Integer, Value As Variant)
    Header.setColumnInfo Col, Value
End Sub

Public Function getColumnInfo(Col As Integer) As Variant
    getColumnInfo = Header.getColumnInfo(Col)
End Function

Private Function GetCol(Key As Variant) As Integer
    Dim n As Integer
    'If the key is numeric the return the actual key,
    'otherwise lookup the column in ColumnNames.
    If IsNumeric(Key) Then
        GetCol = Key
    Else
        For n = ColumnCount To 0 Step -1 '0 if not found
            If Header.getColumnName(n) = Key Then
                GetCol = n
                Exit For
            End If
        Next n
    End If
End Function

Private Property Let ColumnCount(Value As Integer)
    m_ColumnCount = Value
End Property

Public Property Get ColumnCount() As Integer
    ColumnCount = m_ColumnCount
End Property

Private Property Let RowCount(Value As Long)
    m_RowCount = Value
End Property

Public Property Get RowCount() As Long
    RowCount = m_RowCount
End Property
