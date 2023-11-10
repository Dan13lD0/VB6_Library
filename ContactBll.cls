VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New ContactDal

Public Function GetContact(code As String) As ADODB.Recordset
If code = "" Then
code = "0"
End If
Set GetContact = dal.GetContact(code)
End Function


Public Sub CreateUpdateContact(code As String, typeContact As String, contact As String, person As String)

If code = "0" Then
    dal.CreateContact typeContact, contact, person
Else
    dal.UpdateContact code, typeContact, contact, person
End If
End Sub

Public Sub DeleteContact(code As String)
If code = "" Then
code = "0"
End If
dal.RemoveContact (code)
End Sub

Public Function GetLastId() As Integer
Dim record As New ADODB.Recordset
Set record = dal.GetLastId()
With record.Fields
GetLastId = .Item(0).Value
End With
End Function
