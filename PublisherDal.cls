VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublisherDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetPublishers(code As String, name As String, status As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Publisher where Name like '%" & name & "%' " & IIf(code <> "0", " and Id = " & code & " ", "") & IIf(status <> "", " and Status = " & status, "")
    Set GetPublishers = conn.GetDataTable(sql)
End Function

Public Function GetPublisher(id As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Publisher where Id =" & id
    Set GetPublisher = conn.GetDataTable(sql)
End Function

Public Function CreatePublisher(name As String)
    sql = "Insert into Publisher(Name,Status) values ('" & name & "',1)"
    conn.ExcCommand (sql)
End Function

Public Function UpdatePublisher(id As String, name As String, status As String)
    sql = "Update Publisher set Name = '" & name & "', Status = " & status & " where Id = " & id
    conn.ExcCommand (sql)
End Function

Public Function EnabledDisabledPublisher(id As String)
    sql = "Update Publisher set Name = 'daneil', Status = (case when Status = 1 then 0 else 1 end) where Id = " & id
    conn.ExcCommand (sql)
End Function


Public Function ExistsPublisher(id As Integer, name As String) As ADODB.Recordset
    sql = "Select * from Publisher where id <> " & id & " and Name = '" & name & "'"
    Set ExistsPublisher = conn.GetDataTable(sql)
End Function

Public Function GetId(name As String) As Integer
    sql = "Select * from Publisher where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    GetId = result.Fields.Item(0).Value
End Function

