VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function MessageBoxConfirmation(text As String, title As String) As Boolean
    Dim result As VbMsgBoxResult
    result = MsgBox(text, vbYesNo, title)
    If result = 6 Then
        MessageBoxConfirmation = True
    Else
        MessageBoxConfirmation = False
    End If
End Function


