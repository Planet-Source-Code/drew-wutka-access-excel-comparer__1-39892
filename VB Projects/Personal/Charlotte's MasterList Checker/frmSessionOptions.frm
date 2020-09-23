VERSION 5.00
Begin VB.Form frmSessionOptions 
   BackColor       =   &H80000008&
   Caption         =   "Session Options"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmSessionOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeleteMatchedFields 
      Caption         =   "Delete Selected Matched Fields"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdAddMatchedFields 
      Caption         =   "Add Matched Fields From Field Lists"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   3015
   End
   Begin VB.ListBox lstMatchedFields 
      Height          =   2985
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   10
      Top             =   3960
      Width           =   6255
   End
   Begin VB.ListBox lstExcelFields 
      Height          =   1815
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ListBox lstAccessFields 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox cmbExcelSheet 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox cmbAccessTable 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "Matched Fields List (Access Field on Left, Excel Field on right)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label lblExcelKey 
      BackColor       =   &H80000008&
      Caption         =   "Key Field: (Double Click above to select)"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lblAccessKey 
      BackColor       =   &H80000008&
      Caption         =   "Key Field: (Double Click above to select)"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "Excel Sheet Fields:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "Access Table Fields:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Excel File Sheets:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Access Database Table:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSessionOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExlCnn As ADODB.Connection
Dim AccCnn As ADODB.Connection
Private Sub cmbAccessTable_Click()
On Error GoTo ErrorHandler
Dim Resp
Me.lstAccessFields.Clear
If "" & Me.cmbAccessTable <> aecDefaultSettings.AccessTable And aecDefaultSettings.MatchListCount > 0 Then
    Resp = MsgBox("Selecting a new table will clear the selected Matched Fields list, " & _
    "and may remove your Key Field Entry.  Do you wish to continue?", vbYesNo + vbExclamation, "Warning")
    If Resp = vbYes Then
        aecDefaultSettings.DeleteAllMatchedFields
    Else
        Me.cmbAccessTable = aecDefaultSettings.AccessTable
        Exit Sub
    End If
End If
Dim rs As ADODB.Recordset
Dim i As Long
Dim strTemp As String
Dim blKeyValid As Boolean
blKeyValid = aecDefaultSettings.AccessKeyField = ""
aecDefaultSettings.AccessTable = Me.cmbAccessTable
Set rs = New ADODB.Recordset
strTemp = Me.cmbAccessTable
rs.Open strTemp, AccCnn, adOpenKeyset, adLockReadOnly, adCmdTableDirect
For i = 0 To rs.Fields.Count - 1
    Me.lstAccessFields.AddItem rs.Fields(i).Name
    If Not blKeyValid Then
        If aecDefaultSettings.AccessKeyField = rs.Fields(i).Name Then blKeyValid = True
    End If
Next i
rs.Close
If Not blKeyValid Then
    aecDefaultSettings.AccessKeyField = ""
End If
SetDefaults "NoTables"
Exit Sub

ErrorHandler:
Me.lstAccessFields.Clear
SetDefaults "NoTables"
Set rs = Nothing
End Sub
Private Sub cmbExcelSheet_Click()
On Error GoTo ErrorHandler
Dim Resp
Me.lstExcelFields.Clear
If "" & Me.cmbExcelSheet <> aecDefaultSettings.ExcelSheet And aecDefaultSettings.MatchListCount > 0 Then
    Resp = MsgBox("Selecting a new table will clear the selected Matched Fields list, " & _
    "and may remove your Key Field Entry.  Do you wish to continue?", vbYesNo + vbExclamation, "Warning")
    If Resp = vbYes Then
        aecDefaultSettings.DeleteAllMatchedFields
    Else
        Me.cmbExcelSheet = aecDefaultSettings.ExcelSheet
        Exit Sub
    End If
End If
Dim rs As ADODB.Recordset
Dim i As Long
Dim strTemp As String
Dim blKeyValid As Boolean
blKeyValid = aecDefaultSettings.ExcelKeyField = ""
aecDefaultSettings.ExcelSheet = Me.cmbExcelSheet
Set rs = New ADODB.Recordset
strTemp = Me.cmbExcelSheet
rs.Open strTemp, ExlCnn, adOpenKeyset, adLockReadOnly, adCmdTableDirect
For i = 0 To rs.Fields.Count - 1
    Me.lstExcelFields.AddItem rs.Fields(i).Name
    If Not blKeyValid Then
        If aecDefaultSettings.ExcelKeyField = rs.Fields(i).Name Then blKeyValid = True
    End If
Next i
rs.Close
If Not blKeyValid Then
    aecDefaultSettings.ExcelKeyField = ""
End If
SetDefaults "NoTables"
Exit Sub

ErrorHandler:

Me.lstExcelFields.Clear
SetDefaults "NoTables"
Set rs = Nothing
End Sub
Private Sub cmdAddMatchedFields_Click()
Dim Resp As String
Resp = aecDefaultSettings.AddMatchListItem("" & Me.lstAccessFields, "" & Me.lstExcelFields)
SetDefaults
If Resp <> "Done" Then MsgBox Resp, vbOKOnly + vbExclamation, "Add Failure"
End Sub
Private Sub cmdDeleteMatchedFields_Click()
Dim i As Long
Dim Resp As String
For i = 0 To Me.lstMatchedFields.ListCount - 1
    If Me.lstMatchedFields.Selected(i) = True Then
        Resp = aecDefaultSettings.DeleteMatchListItem(Me.lstMatchedFields.ItemData(i))
        If Resp <> "Done" Then MsgBox Resp, vbOKOnly + vbExclamation, "Deletion Failure"
    End If
Next i
SetDefaults
End Sub
Private Sub Form_Load()
Dim rs As ADODB.Recordset
If Dir(aecDefaultSettings.AccessPath) = "" Or aecDefaultSettings.AccessPath = "" Then
    MsgBox "Unable to find selected Access Database.", vbCritical + vbOKOnly, "Missing Database"
    Unload Me
    Exit Sub
End If
If Dir(aecDefaultSettings.ExcelPath) = "" Or aecDefaultSettings.ExcelPath = "" Then
    MsgBox "Unable to find selected Excel File.", vbCritical + vbOKOnly, "Missing Excel File"
    Unload Me
    Exit Sub
End If
OpenExcelConnection ExlCnn
OpenAccessConnection AccCnn
Set rs = New ADODB.Recordset
Set rs = ExlCnn.OpenSchema(adSchemaTables)
If rs.EOF = False Then rs.MoveFirst
Do Until rs.EOF = True
    If rs.Fields("TABLE_TYPE") = "TABLE" Then
        Me.cmbExcelSheet.AddItem rs.Fields("TABLE_NAME").Value
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset
Set rs = AccCnn.OpenSchema(adSchemaTables)
If rs.EOF = False Then rs.MoveFirst
Do Until rs.EOF = True
    If rs.Fields("TABLE_TYPE") = "TABLE" Then
        Me.cmbAccessTable.AddItem rs.Fields("TABLE_NAME").Value
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
On Error Resume Next
SetDefaults
Me.Show
End Sub
Private Sub SetDefaults(Optional NoTables As String)
Dim mli As MatchListInfo
Dim i As Long
If NoTables <> "NoTables" Then
    Me.cmbAccessTable = aecDefaultSettings.AccessTable
    Me.cmbExcelSheet = aecDefaultSettings.ExcelSheet
End If
Me.lblAccessKey.Caption = "Key Field: (Double Click above to select)" & vbCrLf & aecDefaultSettings.AccessKeyField
Me.lblExcelKey.Caption = "Key Field: (Double Click above to select)" & vbCrLf & aecDefaultSettings.ExcelKeyField
Me.lstMatchedFields.Clear
For i = 1 To aecDefaultSettings.MatchListCount
    Set mli = aecDefaultSettings.MatchListItem(i)
    Me.lstMatchedFields.AddItem mli.AccessField & " <==> " & mli.ExcelField
    Me.lstMatchedFields.ItemData(Me.lstMatchedFields.NewIndex) = mli.ID
Next i
Set mli = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
AccCnn.Close
ExlCnn.Close
Set AccCnn = Nothing
Set ExlCnn = Nothing
frmMain.mnuSession.Enabled = True
frmMain.mnuFile.Enabled = True
Exit Sub

ErrorHandler:
'frmMain.Show
Resume Next
End Sub
Private Sub lstAccessFields_DblClick()
aecDefaultSettings.AccessKeyField = Me.lstAccessFields
SetDefaults
End Sub
Private Sub lstExcelFields_DblClick()
aecDefaultSettings.ExcelKeyField = Me.lstExcelFields
SetDefaults
End Sub
