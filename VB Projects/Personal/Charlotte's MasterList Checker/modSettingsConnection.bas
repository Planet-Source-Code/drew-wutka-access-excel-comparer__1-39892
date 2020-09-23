Attribute VB_Name = "modSettingsConnection"
Option Explicit
Function SettingsConnection(ByRef cnn As ADODB.Connection)
Set cnn = New ADODB.Connection
With cnn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Open App.Path & "\AECSettings.mdb"
End With
End Function
