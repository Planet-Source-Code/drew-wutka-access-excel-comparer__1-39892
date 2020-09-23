VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Property Get ExportTypeCount() As Long
ExportTypeCount = 6
End Property
Property Get ExportTypeWord(intPos As Long) As String
Dim strTemp As String
Select Case intPos
    Case 1
        strTemp = "Missing"
    Case 2
        strTemp = "Adding"
    Case 3
        strTemp = "Diff"
    Case 4
        strTemp = "Updated"
    Case 5
        strTemp = "Multiple"
    Case 6
        strTemp = "Error"
End Select
ExportTypeWord = strTemp
End Property
Property Get ExportTypeDescription(intPos As Long) As String
Dim strTemp As String
Select Case intPos
    Case 1
        strTemp = "Missing Records"
    Case 2
        strTemp = "Added Records"
    Case 3
        strTemp = "Different Records"
    Case 4
        strTemp = "Updated Records"
    Case 5
        strTemp = "Multiple Records"
    Case 6
        strTemp = "Errors"
End Select
ExportTypeDescription = strTemp
End Property
