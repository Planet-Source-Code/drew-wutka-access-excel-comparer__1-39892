VERSION 5.00
Begin VB.Form frmSessions 
   BackColor       =   &H80000008&
   Caption         =   "Existing Sessions"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   2970
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstSessions 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
SettingsConnection cnn
strSQL = "SELECT tblSessions.SessionID, tblSessions.SessionName " & _
"From tblSessions " & _
"ORDER BY tblSessions.SessionName;"
Set rs = New ADODB.Recordset
rs.Open strSQL, cnn, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
    rs.MoveFirst
    Do Until rs.EOF = True
        Me.lstSessions.AddItem rs.Fields(1).Value
        Me.lstSessions.ItemData(Me.lstSessions.NewIndex) = rs.Fields(0).Value
        rs.MoveNext
    Loop
Else
    MsgBox "There are no existing Sessions.", vbOKOnly + vbInformation, "No Sessions"
End If
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing
If Me.lstSessions.ListCount = 1 Then
    frmMain.OpenSession Me.lstSessions.ItemData(0)
    Unload Me
End If
End Sub
Private Sub lstSessions_DblClick()
frmMain.OpenSession Me.lstSessions.ItemData(Me.lstSessions.ListIndex)
Unload Me
End Sub
