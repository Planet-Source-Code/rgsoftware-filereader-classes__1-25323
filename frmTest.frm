VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin FileReaderExample.DataSelector DataSelector1 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7646
   End
   Begin VB.PictureBox picHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Your ASCII Data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   2310
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   6000
         Picture         =   "frmTest.frx":1CCA
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "What delimiter seperates your columns?  Select the appropriate delimiter to see how your data is affected in the preview below."
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   400
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FileReader, Row and HeaderInfo classes are the ones that load and manage the data.
'DataSelector control is seperate from these classes and it is only meant to
'show how you might use the classes together.

'You can use the DataSelctor control to import data into SQL server, Access, etc.

'Copyright 2001 by RG Software Corporation http://www.rgsoftware.com
'Author: Richard Gardner rgardner@rgsoftware.com

Private Sub DataSelector1_Change(Header As HeaderInfo)
    'You can capture the events as they are changed
    Dim Info As String
    Dim n As Integer
    For n = 1 To Header.ColumnCount
        Info = Info & Header.getColumnName(n) & "  "
    Next n
    Caption = Info
    Set Header = Nothing
End Sub

Private Sub Form_Load()
    DataSelector1.AddOptions "Input Column,  Output Column,Target Column"
    DataSelector1.FileName = App.Path & "\Win32API.csv"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim n As Integer
    Dim Header As HeaderInfo
    Dim Info As String
    Set Header = DataSelector1.getUserSettings
    For n = 1 To Header.ColumnCount
    Info = Info & "Name: " & Header.getColumnName(n) & " Options: " & Header.getColumnInfo(n) & vbCrLf
    Next n
    Set Header = Nothing
    MsgBox Info
End Sub
