VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl DataSelector 
   AccessKeys      =   "Ctl+FR"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ScaleHeight     =   4260
   ScaleWidth      =   6855
   ToolboxBitmap   =   "ctlFRSelectData.ctx":0000
   Begin VB.CheckBox chkFirstRow 
      Caption         =   "&First Row Contains Column Names"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   2895
   End
   Begin VB.Frame fraDelimiter 
      Caption         =   "Choose the delimiter that seperates your columns:"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6615
      Begin VB.OptionButton optDelimiter 
         Caption         =   "&Tab"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   350
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optDelimiter 
         Caption         =   "S&emicolon"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   350
         Width           =   1095
      End
      Begin VB.OptionButton optDelimiter 
         Caption         =   "&Comma"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   350
         Width           =   855
      End
      Begin VB.OptionButton optDelimiter 
         Caption         =   "&Space"
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   9
         Top             =   350
         Width           =   855
      End
      Begin VB.OptionButton optDelimiter 
         Caption         =   "&Other"
         Height          =   195
         Index           =   4
         Left            =   5160
         TabIndex        =   8
         Top             =   350
         Width           =   735
      End
      Begin VB.TextBox txtOther 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6000
         TabIndex        =   7
         Top             =   320
         Width           =   330
      End
   End
   Begin VB.ComboBox cboQualifier 
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Text            =   "cboQualifier"
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Field Options"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
      Begin VB.TextBox txtColName 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   280
         Width           =   1530
      End
      Begin VB.ComboBox cboOptions 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Column Name:"
         Height          =   195
         Index           =   0
         Left            =   350
         TabIndex        =   4
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Options:"
         Height          =   195
         Index           =   1
         Left            =   3320
         TabIndex        =   3
         Top             =   315
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxPreview 
      Height          =   1845
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3254
      _Version        =   393216
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   -2147483644
      GridColor       =   -2147483640
      GridColorFixed  =   -2147483640
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "ctlFRSelectData.ctx":0312
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Text &Qualifier:"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   855
      Width           =   975
   End
End
Attribute VB_Name = "DataSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#package vb.com.rgsoftware.io
'#package vb.com.rgsoftware.gui

Option Explicit
Option Compare Text

Private FR As New FileReader
Private Delimiter As String
Private m_FileName As String

Public Event Change(Header As HeaderInfo)

Private Sub cboOptions_Click()
    FR.setColumnInfo flxPreview.Col + 1, cboOptions.Text
End Sub

Private Sub txtOther_Change()
    optDelimiter_Click 4
End Sub

Private Sub cboQualifier_Change()
    Display
End Sub

Private Sub chkFirstRow_Click()
    Display
End Sub

Private Sub flxPreview_Click()
    Dim n As Integer
    Dim Info As String
    If flxPreview.Rows = 2 Then Exit Sub
    If flxPreview.Cols = 2 Then Exit Sub
    flxPreview.Row = 0
    flxPreview.RowSel = flxPreview.Rows - 1
    If flxPreview.Col = -1 Then Exit Sub
    txtColName = FR.getColumnName(flxPreview.Col + 1)
    txtColName.SetFocus
    txtColName.SelStart = 0
    txtColName.SelLength = Len(txtColName)
    'Display column options
    Info = FR.getColumnInfo(flxPreview.Col + 1)
    cboOptions.ListIndex = 0
    For n = 0 To cboOptions.ListCount - 1
        If cboOptions.List(n) = Info Then
            cboOptions.ListIndex = n
            Exit For
        End If
    Next n
End Sub

Public Sub AddOptions(ByVal OptionList As String)
    Dim Found As Integer
    Dim Temp As String
    Dim n As Integer
    Do
        Found = InStr(OptionList, ",")
        If Found <> 0 Then
            Temp = Mid$(OptionList, 1, Found - 1)
            cboOptions.AddItem Temp
            OptionList = Trim$(Mid$(OptionList, Found + 1))
        Else
            OptionList = Trim$(OptionList)
            cboOptions.AddItem OptionList
            Exit Do
        End If
    Loop
End Sub

Public Sub RemoveOptions(ByVal OptionList As String)
    Dim Found As Integer
    Dim Temp As String
    Dim n As Integer
    Do
        Found = InStr(OptionList, ",")
        If Found <> 0 Then
            Temp = Mid$(OptionList, 1, Found - 1)
            For n = 0 To cboOptions.ListCount - 1
                If cboOptions.List(n) = Temp Then
                    If OptionList <> "NONE" Then cboOptions.RemoveItem (n)
                    Exit For
                End If
            Next n
            OptionList = Trim$(Mid$(OptionList, Found + 1))
        Else
            OptionList = Trim$(OptionList)
            For n = 0 To cboOptions.ListCount - 1
                If cboOptions.List(n) = OptionList Then
                    If OptionList <> "NONE" Then cboOptions.RemoveItem (n)
                    Exit For
                End If
            Next n
            Exit Do
        End If
    Loop
End Sub

Public Function getUserSettings() As HeaderInfo
    Set getUserSettings = FR.getHeaderObject
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set FR = Nothing
End Sub

Private Sub optDelimiter_Click(Index As Integer)
    txtOther.BackColor = vbButtonFace
    Select Case Index
        Case 0
            Delimiter = vbTab
        Case 1
            Delimiter = ";"
        Case 2
            Delimiter = ","
        Case 3
            Delimiter = " "
        Case 4
            txtOther.BackColor = vbWindowBackground
            Delimiter = txtOther
    End Select
    Display
End Sub

Private Sub Display()

    Dim intLoop As Integer
    Dim n As Integer
    Dim Row As Long

    On Error Resume Next

    If FileName = "" Then Exit Sub

    FR.readFile FileName, Delimiter, CBool(chkFirstRow.Value), cboQualifier.Text, 10

    txtColName = ""
    flxPreview.Clear
    flxPreview.Cols = 0
    flxPreview.Rows = 1

    'Add the columns
    If chkFirstRow.Value = 0 Then 'Make sure column names are available
        For intLoop = 1 To FR.ColumnCount
            flxPreview.Cols = flxPreview.Cols + 1
            flxPreview.Col = intLoop - 1
            flxPreview.Row = 0
            flxPreview.Text = "Column" & intLoop
            FR.setColumnName intLoop, "Column" & intLoop
        Next intLoop
    Else
        For intLoop = 1 To FR.ColumnCount
            flxPreview.Cols = flxPreview.Cols + 1
            flxPreview.Col = intLoop - 1
            flxPreview.Row = 0
            flxPreview.Text = FR.getColumnName(intLoop)
        Next intLoop
    End If

    'Set default option to NONE
    For intLoop = 1 To FR.ColumnCount
        FR.setColumnInfo intLoop, "NONE"
    Next intLoop

    'Set column captions
    flxPreview.Row = 0
    For intLoop = 1 To FR.ColumnCount
        flxPreview.Col = intLoop - 1
        flxPreview.Text = FR.getColumnName(intLoop)
    Next intLoop

    'Fill the grid with the 10 rows of preview data
    flxPreview.Rows = 2
    For intLoop = 1 To FR.RowCount
        flxPreview.Row = flxPreview.Rows - 1
        For n = 1 To FR.ColumnCount
            flxPreview.Col = n - 1
            flxPreview.Text = FR.getData(n, flxPreview.Rows - 1)
        Next n
        flxPreview.Rows = flxPreview.Rows + 1
    Next intLoop
    flxPreview.Rows = flxPreview.Rows - 1
    
    RaiseEvent Change(getUserSettings)

End Sub

Private Sub txtColName_Change()
    Dim n As Integer
    'Validate
    If flxPreview.Col = 1 Then Exit Sub
    If txtColName = "" Then Exit Sub
    'Look for duplicate column names
    flxPreview.Row = 0
    For n = 1 To FR.ColumnCount
        If n <> flxPreview.Col + 1 Then
            If FR.getColumnName(n) = txtColName Then
                'Duplicate column name
                Exit Sub
            End If
        End If
    Next n
    FR.setColumnName flxPreview.Col + 1, txtColName
    flxPreview.Text = txtColName
    flxPreview.RowSel = flxPreview.Rows - 1
    RaiseEvent Change(getUserSettings)
End Sub

Private Sub UserControl_Initialize()
    Delimiter = vbTab
    cboQualifier.AddItem Chr$(34)
    cboQualifier.AddItem "'"
    cboQualifier.AddItem "{none}"
    cboQualifier.Text = "{none}"
    cboOptions.AddItem "NONE"
End Sub

Public Property Let FileName(Value As String)
    m_FileName = Value
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property
