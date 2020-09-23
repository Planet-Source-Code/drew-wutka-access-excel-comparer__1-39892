VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Defaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim intSessionID As Long
Dim MatchList As Collection
Dim strAccessPath As String
Dim strExcelPath As String
Dim strDefaultAccessTable As String
Dim strDefaultExcelSheet As String
Dim strAccessKey As String
Dim strExcelKey As String
Event SessionChanged()
Event DefaultHasChanged(ByVal strMsg As String, ByRef blChangeSetting As Boolean)
Function DeleteNonFileSettings()
Dim cnn As ADODB.Connection
Dim strSQL As String
strSQL = "UPDATE tblDefaultSettings SET tblDefaultSettings.DefaultAccessTable " & _
"= '', tblDefaultSettings.DefaultExcelSheet = '', tblDefaultSettings.AccessKey" & _
"Field = '', tblDefaultSettings.ExcelKeyField = '' " & _
"WHERE (((tblDefaultSettings.SessionID)=" & intSessionID & "));"
SettingsConnection cnn
cnn.Execute strSQL
cnn.Close
Set cnn = Nothing
GetSessionData
End Function
Function DeleteAllMatchedFields()
Dim cnn As ADODB.Connection
Dim strSQL As String
strSQL = "DELETE tblFieldMatchList.*, tblFieldMatchList.SessionID " & _
"From tblFieldMatchList " & _
"WHERE (((tblFieldMatchList.SessionID)=" & intSessionID & "));"
SettingsConnection cnn
cnn.Execute strSQL
Set MatchList = New Collection
cnn.Close
Set cnn = Nothing
End Function
Property Get MatchListCount() As Long
MatchListCount = MatchList.Count
End Property
Property Get MatchListItem(intPos As Long) As MatchListInfo
Set MatchListItem = MatchList(intPos)
End Property
Function AddMatchListItem(strAccess As String, strExcel As String) As String
On Error GoTo ErrorHandler
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mli As MatchListInfo
If strAccess = "" Or strExcel = "" Then
    AddMatchListItem = "Blank Field Names are not allowed!"
    Exit Function
End If
SettingsConnection cnn
Set rs = New ADODB.Recordset
rs.Open "tblFieldMatchList", cnn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rs.AddNew
rs.Fields(1).Value = intSessionID
rs.Fields(2).Value = strAccess
rs.Fields(3).Value = strExcel
rs.Update
Set mli = New MatchListInfo
mli.ID = rs.Fields(0).Value
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing
mli.AccessField = strAccess
mli.ExcelField = strExcel
MatchList.Add mli, "ID:" & mli.ID
Set mli = Nothing
AddMatchListItem = "Done"
Exit Function

ErrorHandler:
AddMatchListItem = Err.Number & ":" & vbCrLf & vbCrLf & Err.Description
End Function
Function DeleteMatchListItem(intID As Long) As String
On Error GoTo ErrorHandler
Dim cnn As ADODB.Connection
Dim strSQL As String
strSQL = "DELETE tblFieldMatchList.*, tblFieldMatchList.MatchListID " & _
"From tblFieldMatchList " & _
"WHERE (((tblFieldMatchList.MatchListID)=" & intID & "));"
SettingsConnection cnn
cnn.Execute strSQL
cnn.Close
Set cnn = Nothing
MatchList.Remove "ID:" & intID
DeleteMatchListItem = "Done"
Exit Function

ErrorHandler:
DeleteMatchListItem = Err.Number & ":" & vbCrLf & vbCrLf & Err.Description
End Function
Property Let SessionID(intEnter As Long)
intSessionID = intEnter
GetSessionData
RaiseEvent SessionChanged
End Property
Property Get ExcelKeyField() As String
ExcelKeyField = strExcelKey
End Property
Property Get AccessKeyField() As String
AccessKeyField = strAccessKey
End Property
Property Let ExcelKeyField(strEnter As String)
If strExcelKey <> strEnter Then
    ChangeSetting "ExcelKeyField", strEnter
End If
strExcelKey = strEnter
End Property
Property Let AccessKeyField(strEnter As String)
If strAccessKey <> strEnter Then
    ChangeSetting "AccessKeyField", strEnter
End If
strAccessKey = strEnter
End Property
Private Function GetSessionData()
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim varTemp As Variant
Dim i As Long
Dim mli As MatchListInfo
SettingsConnection cnn
Set rs = New ADODB.Recordset
strSQL = "SELECT tblDefaultSettings.DefaultAccessPath, tblDefaultSettings.DefaultExcel" & _
"File, tblDefaultSettings.DefaultAccessTable, tblDefaultSettings.DefaultExcelSheet, " & _
"tblDefaultSettings.AccessKeyField, tblDefaultSettings.ExcelKeyField " & _
"From tblDefaultSettings " & _
"WHERE (((tblDefaultSettings.SessionID)=" & intSessionID & "));"
rs.Open strSQL, cnn, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
    rs.MoveFirst
    For i = 0 To 5
        varTemp = rs.Fields(i).Value
        If IsNull(varTemp) Then varTemp = ""
        Select Case i
            Case 0
                strAccessPath = varTemp
            Case 1
                strExcelPath = varTemp
            Case 2
                strDefaultAccessTable = varTemp
            Case 3
                strDefaultExcelSheet = varTemp
            Case 4
                strAccessKey = varTemp
            Case 5
                strExcelKey = varTemp
        End Select
    Next i
Else
    strAccessPath = ""
    strExcelPath = ""
    strDefaultAccessTable = ""
    strDefaultExcelSheet = ""
End If
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset
Set MatchList = New Collection
strSQL = "SELECT tblFieldMatchList.MatchListID, tblFieldMatchList.AccessF" & _
"ield, tblFieldMatchList.ExcelField " & _
"From tblFieldMatchList " & _
"WHERE (((tblFieldMatchList.SessionID)=" & intSessionID & "));"
rs.Open strSQL, cnn, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then rs.MoveFirst
Do Until rs.EOF = True
    Set mli = New MatchListInfo
    mli.ID = rs.Fields(0).Value
    mli.AccessField = rs.Fields(1).Value
    mli.ExcelField = rs.Fields(2).Value
    MatchList.Add mli, "ID:" & mli.ID
    Set mli = Nothing
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing
End Function
Property Get SessionID() As Long
SessionID = intSessionID
End Property
Private Function ChangeSetting(strField As String, strValue As String)
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
SettingsConnection cnn
Set rs = New ADODB.Recordset
strSQL = "SELECT tblDefaultSettings.*" & _
"From tblDefaultSettings " & _
"WHERE (((tblDefaultSettings.SessionID)=" & intSessionID & "));"
rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
If rs.EOF = False Then
    rs.MoveFirst
    rs.Fields(strField).Value = strValue
    rs.Update
Else
    rs.AddNew
    rs.Fields("SessionID").Value = intSessionID
    rs.Fields(strField).Value = strValue
    rs.Update
End If
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing
End Function
Property Let AccessPath(strEnter As String)
Dim blResp As Boolean
If strAccessPath <> strEnter Then
    strAccessPath = strEnter
    blResp = False
    RaiseEvent DefaultHasChanged("Would you like to set this database to be the " & _
    "default Access Database?", blResp)
    If blResp Then ChangeSetting "DefaultAccessPath", strEnter
End If
End Property
Property Let ExcelSheet(strEnter As String)
If strDefaultExcelSheet <> strEnter Then
    ChangeSetting "DefaultExcelSheet", strEnter
End If
strDefaultExcelSheet = strEnter
End Property
Property Let AccessTable(strEnter As String)
If strDefaultAccessTable <> strEnter Then
    ChangeSetting "DefaultAccessTable", strEnter
End If
strDefaultAccessTable = strEnter
End Property
Property Let ExcelPath(strEnter As String)
Dim blResp As Boolean
If strExcelPath <> strEnter Then
    strExcelPath = strEnter
    blResp = False
    RaiseEvent DefaultHasChanged("Would you like to set this Excel File to be the " & _
    "default Excel File Database?", blResp)
    If blResp Then ChangeSetting "DefaultExcelFile", strEnter
End If
End Property
Property Get AccessPath() As String
AccessPath = strAccessPath
End Property
Property Get ExcelPath() As String
ExcelPath = strExcelPath
End Property
Property Get AccessTable() As String
AccessTable = strDefaultAccessTable
End Property
Property Get ExcelSheet() As String
ExcelSheet = strDefaultExcelSheet
End Property
