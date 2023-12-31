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

Public Function OnlyNumbers(e As Integer) As Boolean
If (e < 48 Or e > 57) And e <> 8 Then
    OnlyNumbers = False
Else
     OnlyNumbers = True
End If
End Function

Public Function NumberAndLetters(e As Integer)
 If e >= 48 And e <= 57 Then '0-9
     NumberAndLetters = True
 ElseIf e >= 65 And e <= 90 Then 'a-z
     NumberAndLetters = True
 ElseIf e >= 97 And e <= 122 Then 'A-Z
     NumberAndLetters = True
 ElseIf e = 8 Then 'back space
     NumberAndLetters = True
 Else
     NumberAndLetters = False
 End If
End Function

