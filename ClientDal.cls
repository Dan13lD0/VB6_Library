VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetClients(id As String, name As String, rg As String, cpf As String, status As String) As ADODB.Recordset
sql = "Select  id,Name,RG,CPF,Status = (case when Status =1 then 'Active' else 'Inactive' end) from Client where Id = (case when '" & id & "' ='0' then Id else '' end) and Name like '%" & name & "%' and RG like '%" & rg & "%' and CPF like '%" & cpf & "%' and Status = (case when '" & status & "' ='' then status else '" & status & "' end)"
Set GetClients = conn.GetDataTable(sql)
End Function

Public Function GetClient(id As String) As ADODB.Recordset
sql = "Select * from Client where Id =" & id
Set GetClient = conn.GetDataTable(sql)
End Function


Public Function CreateClient(name As String, birthday As String, age As String, rg As String, cpf As String, address As String, contact As String)
sql = "Insert Into Client(Name,BirthDay,Age,RG,CPF,AddressId,ContactId,Status) values('" & name & "','" & birthday & "','" & age & "','" & rg & "','" & cpf & "'," & address & "," & contact & ",1)"
conn.ExcCommand (sql)
End Function


Public Function UpdateClient(id As String, name As String, birthday As String, age As String, rg As String, cpf As String, address As String, contact As String, status As String)
sql = "Update Client set Name='" & name & "',BirthDay='" & birthday & "',Age='" & age & "',RG='" & rg & "',CPF='" & cpf & "',AddressId=" & address & ",ContactId=" & contact & ",Status=" & status & " where id=" & id
conn.ExcCommand (sql)
End Function


Public Function EnableDisableClient(id As String)
sql = "Update Client set Status=(case when status =0 then 1 else 0 end) where id =" & id
conn.ExcCommand (sql)
End Function

Public Function ExistsClient(id As String, name As String)
sql = "Select * from Client where id <> " & id & " and Name = '" & name & "'"
conn.ExcCommand (sql)
End Function

Public Function GetId(name As String) As Integer
    sql = "Select * from Client where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    GetId = result.Fields.Item(0).Value
End Function

Public Function GetClientByName(name As String) As ADODB.Recordset
    sql = "Select * from Client where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    Set GetClientByName = result
End Function


