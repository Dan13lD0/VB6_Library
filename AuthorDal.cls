VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AuthorDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetAuthors(code As String, name As String, status As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Author where Name like '%" & name & "%' " & IIf(code <> "0", " and Id = " & code & " ", "") & IIf(status <> "", " and Status = " & status, "")
    Set GetAuthors = conn.GetDataTable(sql)
End Function

Public Function GetAuthor(id As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Author where Id =" & id
    Set GetAuthor = conn.GetDataTable(sql)
End Function

Public Function CreateAuthor(name As String)
    sql = "Insert into Author(Name,Status) values ('" & name & "',1)"
    conn.ExcCommand (sql)
End Function

Public Function UpdateAuthor(id As String, name As String, status As String)
    sql = "Update Author set Name = '" & name & "', Status = " & status & " where Id = " & id
    conn.ExcCommand (sql)
End Function

Public Function EnabledDisabledAuthor(id As String)
    sql = "Update Author set Name = 'daneil', Status = (case when Status = 1 then 0 else 1 end) where Id = " & id
    conn.ExcCommand (sql)
End Function


Public Function ExistsAuthor(id As Integer, name As String) As ADODB.Recordset
    sql = "Select * from Author where id <> " & id & " and Name = '" & name & "'"
    Set ExistsAuthor = conn.GetDataTable(sql)
End Function

Public Function GetId(name As String) As Integer
    sql = "Select * from Author where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    GetId = result.Fields.Item(0).Value
End Function

