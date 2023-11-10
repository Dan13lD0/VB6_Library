VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New AddressDal

Public Function GetAddress(code As String) As ADODB.Recordset
If code = "" Then
code = "0"
End If

Set GetAddress = dal.GetAddress(code)
End Function

Public Sub CreateUpdateAddress(code As String, zipCode As String, street As String, complement As String, city As String, state As String)
If code = "0" Then
    dal.CreateAddress zipCode, street, complement, city, state
Else
    dal.UpdateAddress code, zipCode, street, complement, city, state
End If
End Sub

Public Sub DeleteAddress(code As String)
If code = "" Then
code = "0"
End If

dal.DeleteAddress (code)
End Sub

Public Function GetLastId() As Integer
Dim record As New ADODB.Recordset
Set record = dal.GetLastId()
With record.Fields
GetLastId = .Item(0).Value
End With
End Function
