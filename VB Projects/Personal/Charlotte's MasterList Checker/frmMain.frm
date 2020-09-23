VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   Caption         =   "Access / Excel Comparer"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRunComparison 
      Caption         =   "Run Comparison"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2055
      Left            =   8280
      TabIndex        =   11
      Top             =   960
      Width           =   2295
      Begin VB.CheckBox chkOptionItem 
         BackColor       =   &H80000012&
         Caption         =   "Clear Results Window"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   16
         ToolTipText     =   "Clears results in the window to the left, when a comparison is run."
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkOptionItem 
         BackColor       =   &H80000012&
         Caption         =   "Update Different Records"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   15
         ToolTipText     =   "Update Records in the Servant Source to match Records in the Master Source."
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox chkOptionItem 
         BackColor       =   &H80000012&
         Caption         =   "List Different Records"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   14
         ToolTipText     =   "List Records in the Servant Source which are different from the Master Source."
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkOptionItem 
         BackColor       =   &H80000012&
         Caption         =   "Add Missing Records"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   13
         ToolTipText     =   "Add records missing from the Servant Source."
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkOptionItem 
         BackColor       =   &H80000012&
         Caption         =   "List Missing Records"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "List Records missing from the Servant Source, that are in the Master Source."
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Switch"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "Switch Master and Servant Source."
      Top             =   720
      Width           =   855
   End
   Begin MSComctlLib.TreeView tvwResults 
      Height          =   4935
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8705
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H80000008&
      Caption         =   "Total Errors: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   25
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label lblUpdated 
      BackColor       =   &H80000008&
      Caption         =   "Total Updated: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   24
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label lblDifferent 
      BackColor       =   &H80000008&
      Caption         =   "Total Different: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblMultiples 
      BackColor       =   &H80000008&
      Caption         =   "Total Multiples: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lblAdded 
      BackColor       =   &H80000008&
      Caption         =   "Total Added: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   21
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblMissing 
      BackColor       =   &H80000008&
      Caption         =   "Total Missing: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblMasterCurrent 
      BackColor       =   &H80000008&
      Caption         =   "Current Master Source Record:"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8280
      TabIndex        =   19
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMasterCount 
      BackColor       =   &H80000008&
      Caption         =   "Master Source # of Records:"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   690
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "(Master Source)."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "(Servant Source)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3135
      TabIndex        =   7
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "to the"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4400
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblMaster 
      BackColor       =   &H80000008&
      Caption         =   "Excel Sheet"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4890
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblServant 
      BackColor       =   &H80000008&
      Caption         =   "Access Table"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2100
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Compare Records from the"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblCurrentExcel 
      BackColor       =   &H80000007&
      Caption         =   "Current Excel File:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10455
   End
   Begin VB.Label lblCurrentDB 
      BackColor       =   &H80000007&
      Caption         =   "Current Database:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
   Begin VB.Menu mnuSession 
      Caption         =   "&Session"
      Begin VB.Menu mnuCreateNewSession 
         Caption         =   "Create New Session"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOpenExistingSession 
         Caption         =   "Open Existing Session"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSessionOptions 
         Caption         =   "Edit Session Options"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCloseSession 
         Caption         =   "Close Session"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuRunComparison 
         Caption         =   "Run Comparison"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSessionSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Enabled         =   0   'False
      Begin VB.Menu mnuSelectAccessDatabase 
         Caption         =   "Select Access Database"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectExcelFile 
         Caption         =   "Select Excel File"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuExportSelection 
         Caption         =   "Export Options"
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Missing"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Added"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Different"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Updated"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Multiples"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mnuExportSelections 
            Caption         =   "Errors"
            Checked         =   -1  'True
            Index           =   5
         End
      End
      Begin VB.Menu mnuExportResults 
         Caption         =   "Export Results"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionItem 
         Caption         =   "List Missing"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptionItem 
         Caption         =   "Add Missing"
         Index           =   1
      End
      Begin VB.Menu mnuOptionItem 
         Caption         =   "List Different"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuOptionItem 
         Caption         =   "Update Different"
         Index           =   3
      End
      Begin VB.Menu mnuOptionItem 
         Caption         =   "Clear Results Window"
         Checked         =   -1  'True
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents CurrentSettings As Defaults
Attribute CurrentSettings.VB_VarHelpID = -1
Dim blCancelComparison As Boolean
Dim intNodeCount As Long
Function RunComparison()
On Error GoTo ErrorHandler
Dim MCnn As ADODB.Connection
Dim SCnn As ADODB.Connection
Dim mRS As ADODB.Recordset
Dim sRS As ADODB.Recordset
Dim intMasterCount As Long
Dim strTemp As String
Dim mli As MatchListInfo
Dim strSQL As String
Dim strCoreSQL As String
Dim strWhereWrapper As String
Dim strCurrentKey As String
Dim i As Long
Dim j As Long
Dim blRecordIsDifferent As Boolean
Dim blNodeCreated As Boolean
Dim strProcess As String
strProcess = "Initializing"
If Me.chkOptionItem(4) = 1 Then
    'Me.tvwResults(1).Nodes.Clear
    Unload Me.tvwResults(1)
    Load Me.tvwResults(1)
    Me.tvwResults(1).Visible = True
    intNodeCount = 1
End If
If Me.lblServant.Caption = "Access Table" Then
    OpenExcelConnection MCnn
    OpenAccessConnection SCnn
    strSQL = "SELECT `" & CurrentSettings.ExcelSheet & "`.* "
    strSQL = strSQL & "From `" & CurrentSettings.ExcelSheet & "`;"
Else
    OpenExcelConnection SCnn
    OpenAccessConnection MCnn
    strSQL = "SELECT " & CurrentSettings.AccessTable & ".* "
    strSQL = strSQL & "From " & CurrentSettings.AccessTable & ";"
End If
Set mRS = New ADODB.Recordset
mRS.Open strSQL, MCnn, adOpenKeyset, adLockReadOnly
If mRS.EOF = False Then
    mRS.MoveLast
    intMasterCount = mRS.RecordCount
    mRS.MoveFirst
Else
    intMasterCount = 0
End If
Me.lblMasterCount.Caption = "Master Source # of Records:" & vbCrLf & intMasterCount
Me.lblMasterCurrent.Caption = "Current Master Source Record:" & vbCrLf & "0"
If intMasterCount > 0 Then
    If Me.lblServant.Caption <> "Access Table" Then
        strCoreSQL = "SELECT `" & CurrentSettings.ExcelSheet & "`.* "
        strCoreSQL = strCoreSQL & "From `" & CurrentSettings.ExcelSheet & "` "
        strTemp = CurrentSettings.ExcelKeyField
    Else
        strCoreSQL = "SELECT " & CurrentSettings.AccessTable & ".* "
        strCoreSQL = strCoreSQL & "From " & CurrentSettings.AccessTable & " "
        strTemp = CurrentSettings.AccessKeyField
    End If
    Set sRS = New ADODB.Recordset
    strSQL = Left(strCoreSQL, Len(strCoreSQL) - 1) & ";"
    sRS.Open strSQL, SCnn, adOpenKeyset, adLockReadOnly
    Select Case sRS.Fields(strTemp).Type
        Case adBigInt, adBinary, adBoolean, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVariant
            strWhereWrapper = ""
        Case adBSTR, adChar, adWChar, adVarWChar, adLongVarWChar
            strWhereWrapper = "'"
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            strWhereWrapper = "#"
    End Select
    sRS.Close
    Set sRS = Nothing
    
    For i = 1 To intMasterCount
        strProcess = "Looping"
        Me.lblMasterCurrent.Caption = "Current Master Source Record:" & vbCrLf & i
        DoEvents
        If blCancelComparison = True Then GoTo StopComparisonLine:
        If Me.lblServant.Caption <> "Access Table" Then
            strSQL = strCoreSQL & "WHERE (((`" & CurrentSettings.ExcelSheet & "`.[" & _
            CurrentSettings.ExcelKeyField & "])=" & strWhereWrapper & _
            mRS.Fields(CurrentSettings.AccessKeyField).Value & strWhereWrapper & _
            "));"
            strCurrentKey = mRS.Fields(CurrentSettings.AccessKeyField).Value
        Else
            strSQL = strCoreSQL & "WHERE (((" & CurrentSettings.AccessTable & "." & _
            CurrentSettings.AccessKeyField & ")=" & strWhereWrapper & _
            mRS.Fields(CurrentSettings.ExcelKeyField).Value & strWhereWrapper & _
            "));"
            strCurrentKey = mRS.Fields(CurrentSettings.ExcelKeyField).Value
        End If
        Set sRS = New ADODB.Recordset
        sRS.Open strSQL, SCnn, adOpenKeyset, adLockOptimistic
        blNodeCreated = False
        If sRS.EOF = True Then
            'We have no matching records
            If Me.chkOptionItem(0).Value = 1 Then
                Me.tvwResults(1).Nodes.Add , , "Missing" & intNodeCount, "Record #" & i & _
                " Master Key Field: " & strCurrentKey & _
                "     Result: No Matching Record"
                blNodeCreated = True
            End If
            If Me.chkOptionItem(1).Value = 1 Then
                strProcess = "Adding"
                sRS.AddNew
                If Me.lblServant.Caption = "Access Table" Then
                    sRS.Fields(CurrentSettings.AccessKeyField).Value = mRS.Fields(CurrentSettings.ExcelKeyField).Value
                    For j = 1 To CurrentSettings.MatchListCount
                        Set mli = CurrentSettings.MatchListItem(j)
                        sRS.Fields(mli.AccessField).Value = mRS.Fields(mli.ExcelField).Value
                    Next j
                    Set mli = Nothing
                Else
                    sRS.Fields(CurrentSettings.ExcelKeyField).Value = mRS.Fields(CurrentSettings.AccessKeyField).Value
                    For j = 1 To CurrentSettings.MatchListCount
                        Set mli = CurrentSettings.MatchListItem(j)
                        sRS.Fields(mli.ExcelField).Value = mRS.Fields(mli.AccessField).Value
                    Next j
                    Set mli = Nothing
                End If
                sRS.Update
                If strProcess <> "Adding" Then
                    Me.tvwResults(1).Nodes.Add , , "Error" & intNodeCount, "Record #" & _
                    i & " Master Key Field: " & _
                    strCurrentKey & "     Result: Failed to Add New Record to Servant Source"
                    Me.tvwResults(1).Nodes.Add "Error" & intNodeCount, tvwChild, "ErrMessage" & intNodeCount, strProcess
                    sRS.Cancel
                Else
                    Me.tvwResults(1).Nodes.Add , , "Adding" & intNodeCount, "Record #" & _
                    i & " Master Key Field: " & _
                    strCurrentKey & "     Result: Added New Record to Servant Source"
                End If
                blNodeCreated = True
            End If
        Else
            sRS.MoveLast
            sRS.MoveFirst
            If sRS.RecordCount > 1 Then
                If Me.chkOptionItem(2).Value = 1 Or Me.chkOptionItem(3).Value = 1 Then
                    Me.tvwResults(1).Nodes.Add , , "Multiple" & intNodeCount, "Record #" & i & _
                    " Master Key Field: " & strCurrentKey & " Result: Multiple Servant " & _
                    "Source Records."
                    blNodeCreated = True
                End If
            Else
                strProcess = "Comparing"
                If Me.chkOptionItem(2).Value = 1 Then
                    blRecordIsDifferent = False
                    If Me.lblServant = "Access Table" Then
                        For j = 1 To CurrentSettings.MatchListCount
                            Set mli = CurrentSettings.MatchListItem(j)
                            If sRS.Fields(mli.AccessField).Value <> mRS.Fields(mli.ExcelField).Value Then blRecordIsDifferent = True
                        Next j
                    Else
                        For j = 1 To CurrentSettings.MatchListCount
                            Set mli = CurrentSettings.MatchListItem(j)
                            If sRS.Fields(mli.ExcelField).Value <> mRS.Fields(mli.AccessField).Value Then blRecordIsDifferent = True
                        Next j
                    End If
                    If blRecordIsDifferent Then
                        Me.tvwResults(1).Nodes.Add , , "Diff" & intNodeCount, "Record #" & i & _
                        " Master Key Field: " & strCurrentKey & " Result: Field Valu" & _
                        "es are Different."
                        If Me.lblServant = "Access Table" Then
                            For j = 1 To CurrentSettings.MatchListCount
                                Set mli = CurrentSettings.MatchListItem(j)
                                If sRS.Fields(mli.AccessField).Value <> mRS.Fields(mli.ExcelField).Value Then
                                    Me.tvwResults(1).Nodes.Add "Diff" & intNodeCount, tvwChild, intNodeCount & _
                                    "FieldComparison" & j, mli.AccessField & ": " & _
                                    sRS.Fields(mli.AccessField).Value & " " & mli.ExcelField & ": " & _
                                    mRS.Fields(mli.ExcelField).Value
                                End If
                            Next j
                        Else
                            For j = 1 To CurrentSettings.MatchListCount
                                Set mli = CurrentSettings.MatchListItem(j)
                                If sRS.Fields(mli.ExcelField).Value <> mRS.Fields(mli.AccessField).Value Then
                                    Me.tvwResults(1).Nodes.Add "Diff" & intNodeCount, tvwChild, intNodeCount & _
                                    "FieldComparison" & j, mli.ExcelField & ": " & _
                                    sRS.Fields(mli.ExcelField).Value & " " & mli.AccessField & ": " & _
                                    mRS.Fields(mli.AccessField).Value
                                End If
                            Next j
                        End If
                        blNodeCreated = True
                     End If
                End If
                strProcess = "Updating"
                If Me.chkOptionItem(3).Value = 1 Then
                    blRecordIsDifferent = False
                    If Me.lblServant = "Access Table" Then
                        For j = 1 To CurrentSettings.MatchListCount
                            Set mli = CurrentSettings.MatchListItem(j)
                            If sRS.Fields(mli.AccessField).Value <> mRS.Fields(mli.ExcelField).Value Then
                                blRecordIsDifferent = True
                                sRS.Fields(mli.AccessField).Value = mRS.Fields(mli.ExcelField).Value
                            End If
                        Next j
                    Else
                        For j = 1 To CurrentSettings.MatchListCount
                            Set mli = CurrentSettings.MatchListItem(j)
                            If sRS.Fields(mli.ExcelField).Value <> mRS.Fields(mli.AccessField).Value Then
                                blRecordIsDifferent = True
                                sRS.Fields(mli.ExcelField).Value = mRS.Fields(mli.AccessField).Value
                            End If
                        Next j
                    End If
                    If blRecordIsDifferent Then sRS.Update
                    If blRecordIsDifferent And strProcess = "Updating" Then
                        Me.tvwResults(1).Nodes.Add , , "Updated" & intNodeCount, "Record #" & _
                        i & " Master Key Field: " & _
                        strCurrentKey & "     Result: Update Servant Source Record to match Master Source."
                        blNodeCreated = True
                    End If
                    If strProcess <> "Updating" Then
                        Me.tvwResults(1).Nodes.Add , , "Error" & intNodeCount, "Record #" & _
                        i & " Master Key Field: " & _
                        strCurrentKey & "     Result: Failure to Update Record in Servant Source."
                        Me.tvwResults(1).Nodes.Add "Error" & intNodeCount, tvwChild, "ErrMessage" & intNodeCount, strProcess
                        sRS.Cancel
                        blNodeCreated = True
                    End If
                End If
            End If
        End If
        
EndOfLoop:
        
        strProcess = "ClosingRecordset"
        sRS.Close
        Set sRS = Nothing
        If blNodeCreated Then intNodeCount = intNodeCount + 1
        mRS.MoveNext
    Next i
End If
strProcess = "ClosingUp"
StopComparisonLine:

mRS.Close
Set mRS = Nothing
MCnn.Close
Set MCnn = Nothing
SCnn.Close
Set SCnn = Nothing
StopComparison
Exit Function

ErrorHandler:
Select Case strProcess
    Case "Initializing"
        MsgBox "The following Error occurred during Initialization: " & vbCrLf & vbCrLf & _
        Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
        Err.Clear
        GoTo StopComparisonLine
    Case "Adding", "Updating"
        strProcess = Err.Number & " " & Err.Description
        Err.Clear
        Resume Next
    Case "ClosingRecordset"
        Err.Clear
        Resume Next
    Case "Looping", "Comparing"
        Me.tvwResults(1).Nodes.Add , , "Error" & intNodeCount, "Record #" & _
        i & " Master Key Field: " & _
        strCurrentKey & "     Result: Unexpected Error"
        Me.tvwResults(1).Nodes.Add "Error" & intNodeCount, tvwChild, "ErrMessage" & intNodeCount, strProcess
        blNodeCreated = True
        Err.Clear
        GoTo EndOfLoop
    Case Else
        Me.tvwResults(1).Nodes.Add , , "Error" & intNodeCount, "Record #" & _
        i & " Master Key Field: " & _
        strCurrentKey & "     Result: Unexpected Error"
        Me.tvwResults(1).Nodes.Add "Adding" & intNodeCount, tvwChild, "ErrMessage" & intNodeCount, strProcess
        blNodeCreated = True
        Err.Clear
        GoTo StopComparisonLine
End Select
End Function
Function StopComparison()
Me.cmdRunComparison.Caption = "Run Comparison"
Me.mnuRunComparison.Enabled = True
GetTreeViewTotals
End Function
Function GetTreeViewTotals()
Dim intMissing As Long
Dim intAdding As Long
Dim intErrors As Long
Dim intMultiples As Long
Dim intDiffs As Long
Dim intUpdated As Long
Dim nd As Node
intMissing = 0
intAdding = 0
intErrors = 0
intMultiples = 0
intDiffs = 0
intUpdated = 0
For Each nd In Me.tvwResults(1).Nodes
    If InStr(1, nd.Key, "Missing", vbTextCompare) = 1 Then intMissing = intMissing + 1
    If InStr(1, nd.Key, "Adding", vbTextCompare) = 1 Then intAdding = intAdding + 1
    If InStr(1, nd.Key, "Error", vbTextCompare) = 1 Then intErrors = intErrors + 1
    If InStr(1, nd.Key, "Multiple", vbTextCompare) = 1 Then intMultiples = intMultiples + 1
    If InStr(1, nd.Key, "Diff", vbTextCompare) = 1 Then intDiffs = intDiffs + 1
    If InStr(1, nd.Key, "Updated", vbTextCompare) = 1 Then intUpdated = intUpdated + 1
Next
Me.lblMissing.Caption = "Total Missing: " & intMissing
Me.lblAdded.Caption = "Total Added: " & intAdding
Me.lblErrors.Caption = "Total Errors: " & intErrors
Me.lblDifferent.Caption = "Total Different: " & intDiffs
Me.lblUpdated.Caption = "Total Updated: " & intUpdated
Me.lblMultiples.Caption = "Total Multiples: " & intMultiples
End Function
Private Sub chkOptionItem_Click(Index As Integer)
Me.mnuOptionItem(Index).Checked = Me.chkOptionItem(Index) = 1
End Sub
Private Sub cmdRunComparison_Click()
If Me.cmdRunComparison.Caption = "Run Comparison" Then
    blCancelComparison = False
    Me.cmdRunComparison.Caption = "Stop Comparison"
    Me.mnuRunComparison.Enabled = False
    RunComparison
Else
    blCancelComparison = True
End If
End Sub
Private Sub cmdSwitch_Click()
Dim strTemp As String
strTemp = Me.lblMaster.Caption
Me.lblMaster.Caption = Me.lblServant.Caption
Me.lblServant.Caption = strTemp
End Sub
Private Sub CurrentSettings_DefaultHasChanged(ByVal strMsg As String, blChangeSetting As Boolean)
Dim Resp
Resp = MsgBox(strMsg, vbYesNo + vbQuestion, "Default Change")
blChangeSetting = Resp = vbYes
ShowCurrentSettings
End Sub
Private Function ShowCurrentSettings()
Me.lblCurrentDB.Caption = "Current Database: " & CurrentSettings.AccessPath
Me.lblCurrentExcel.Caption = "Current Excel File: " & CurrentSettings.ExcelPath
Me.mnuFile.Enabled = CurrentSettings.SessionID > 0
Me.mnuRunComparison.Enabled = CurrentSettings.SessionID > 0
Me.cmdRunComparison.Enabled = CurrentSettings.SessionID > 0
Me.mnuEditSessionOptions.Enabled = CurrentSettings.SessionID > 0
End Function
Private Sub CurrentSettings_SessionChanged()
ShowCurrentSettings
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
intNodeCount = 1
Load Me.tvwResults(1)
Me.tvwResults(1).Visible = True
If aecDefaultSettings.SessionID > 0 Then
    Set CurrentSettings = aecDefaultSettings
Else

NoObject:

    Set CurrentSettings = New Defaults
    Set aecDefaultSettings = CurrentSettings
End If
Me.Show
ShowCurrentSettings
Exit Sub

ErrorHandler:
Err.Clear
GoTo NoObject
End Sub
Private Sub mnuCloseSession_Click()
CurrentSettings.SessionID = 0
End Sub
Private Sub mnuCreateNewSession_Click()
Dim Resp
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Resp = InputBox("What would you like to call this Session?", "New Session Name")
If Resp <> "" Then
    SettingsConnection cnn
    Set rs = New ADODB.Recordset
    rs.Open "tblSessions", cnn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rs.AddNew
    rs.Fields(1).Value = Resp
    rs.Update
    CurrentSettings.SessionID = rs.Fields(0).Value
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End If
End Sub
Private Sub mnuEditSessionOptions_Click()
On Error Resume Next
Me.mnuSession.Enabled = False
Me.mnuFile.Enabled = False
Load frmSessionOptions
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuExportResults_Click()
Dim nd As Node
Dim et As ExportTypes
Dim i As Long
Dim f As Long
Dim strPath As String
Dim strTemp As String
Dim Resp
Dim chNd As Node
Dim intChildren As Long
Me.cdMain.Filter = "Text File|*.txt|Any File|*.*"
Me.cdMain.ShowSave
strPath = Me.cdMain.FileName
If strPath = "" Then Exit Sub
If Dir(strPath) <> "" Then
    Resp = MsgBox("Would you like to write over the existing file?", vbYesNo + vbQuestion, "File Exists")
    If Resp = vbYes Then
        Kill strPath
    Else
        Exit Sub
    End If
End If
f = FreeFile
Open strPath For Binary Access Write As f
Set et = New ExportTypes
For i = 1 To et.ExportTypeCount
    If Me.mnuExportSelections(i - 1).Checked Then
        strTemp = et.ExportTypeDescription(i) & ":" & vbCrLf & vbCrLf
        Put f, , strTemp
        For Each nd In Me.tvwResults(1).Nodes
            If InStr(1, nd.Key, et.ExportTypeWord(i), vbTextCompare) = 1 Then
                strTemp = nd.Text & vbCrLf
                Put f, , strTemp
                If nd.Children > 0 Then
                    intChildren = nd.Children
                    Set chNd = nd.Child
                    strTemp = Chr(9) & chNd.Text & vbCrLf
                    Put f, , strTemp
                    intChildren = intChildren - 1
                    Do Until intChildren = 0
                        Set chNd = chNd.Next
                        strTemp = Chr(9) & chNd.Text & vbCrLf
                        Put f, , strTemp
                        intChildren = intChildren - 1
                    Loop
                    Set chNd = Nothing
                End If
            End If
        Next
        strTemp = vbCrLf & vbCrLf
        Put f, , strTemp
    End If
Next i
Close f
Set et = Nothing
MsgBox "Export Complete!", vbOKOnly + vbInformation, "Done"
End Sub
Private Sub mnuExportSelections_Click(Index As Integer)
Me.mnuExportSelections(Index).Checked = Not Me.mnuExportSelections(Index).Checked
End Sub
Private Sub mnuOpenExistingSession_Click()
On Error Resume Next
frmSessions.Show
End Sub
Public Function OpenSession(intSessionID As Long)
CurrentSettings.SessionID = intSessionID
End Function
Private Sub mnuOptionItem_Click(Index As Integer)
Me.mnuOptionItem(Index).Checked = Not Me.mnuOptionItem(Index).Checked
If Me.mnuOptionItem(Index).Checked Then
    Me.chkOptionItem(Index).Value = 1
Else
    Me.chkOptionItem(Index) = 0
End If
End Sub
Private Sub mnuRunComparison_Click()
blCancelComparison = False
Me.cmdRunComparison.Caption = "Stop Comparison"
Me.mnuRunComparison.Enabled = False
RunComparison
End Sub
Private Sub mnuSelectAccessDatabase_Click()
Dim strTemp As String
Me.cdMain.Filter = "Microsoft Access Databases|*.mdb;*.mde"
Me.cdMain.DialogTitle = "Select Microsoft Access Database to Compare"
Me.cdMain.ShowOpen
strTemp = Me.cdMain.FileName
If Dir(strTemp) <> "" Then
    CurrentSettings.DeleteAllMatchedFields
    CurrentSettings.DeleteNonFileSettings
    CurrentSettings.AccessPath = strTemp
End If
End Sub
Private Sub mnuSelectExcelFile_Click()
Dim strTemp As String
Me.cdMain.Filter = "Microsoft Excel Files|*.xls"
Me.cdMain.DialogTitle = "Select Microsoft Excel File to Compare"
Me.cdMain.ShowOpen
strTemp = Me.cdMain.FileName
If Dir(strTemp) <> "" Then
    CurrentSettings.DeleteAllMatchedFields
    CurrentSettings.DeleteNonFileSettings
    CurrentSettings.ExcelPath = strTemp
End If
End Sub
