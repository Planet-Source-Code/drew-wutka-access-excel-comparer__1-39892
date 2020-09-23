Attribute VB_Name = "modCompareConnections"
Option Explicit

Function OpenExcelConnection(ByRef cnn As ADODB.Connection)
Set cnn = New ADODB.Connection
With cnn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Extended Properties") = "Excel 8.0"
    .Open aecDefaultSettings.ExcelPath
End With
End Function
Function OpenAccessConnection(ByRef cnn As ADODB.Connection)
Set cnn = New ADODB.Connection
With cnn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open aecDefaultSettings.AccessPath
End With
End Function
