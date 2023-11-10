VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New ClientDal

Public Function GetClients(code As String, name As String, rg As String, cpf As String, status As String) As ADODB.Recordset
If code = "" Then
    code = "0"
End If

If status <> "" And status <> "All" Then
    If status = "Active" Then
        status = "1"
    Else
        status = "0"
    End If
Else
status = ""
End If

cpf = Replace(cpf, ".", "")
cpf = Replace(cpf, "-", "")
cpf = Replace(cpf, ",", "")
rg = Replace(rg, ".", "")
rg = Replace(rg, "-", "")
rg = Replace(rg, ",", "")

Set GetClients = dal.GetClients(code, name, rg, cpf, status)

End Function

 Public Function GetClient(code As String) As ADODB.Recordset
    If code = "" Then
        code = "0"
    End If
    
    Set GetClient = dal.GetClient(code)
 End Function

Public Sub CreateUpdateClient(code As String, name As String, birthday As String, rg As String, cpf As String, contactId As String, addressId As String, status As String)
    If code = "" Then
        dal.CreateClient name, birthday, "", rg, cpf, addressId, contactId
    Else
        dal.UpdateClient code, name, birthday, "", rg, cpf, addressId, contactId, IIf(status = "Active", 1, 0)
    End If
End Sub

Public Sub DeleteClient(code As String)
    If code = "" Then
        code = "0"
    End If
    dal.EnableDisableClient (code)
End Sub


Public Function GetClientByName(name As String) As ADODB.Recordset
Set GetClientByName = dal.GetClientByName(name)
End Function
