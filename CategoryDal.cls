VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetCategories(code As String, name As String, status As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Category where Name like '%" & name & "%' " & IIf(code <> "0", " and Id = " & code & " ", "") & IIf(status <> "", " and Status = " & status, "")
    Set GetCategories = conn.GetDataTable(sql)
End Function

Public Function GetCategory(id As String) As ADODB.Recordset
    sql = "Select Id,Name,Status =case when Status = 1 then 'Active' else 'Inactive' end from Category where Id =" & id
    Set GetCategory = conn.GetDataTable(sql)
End Function

Public Sub CreateCategory(name As String)
    sql = "Insert into Category(Name,Status) values ('" & name & "',1)"
    conn.ExcCommand (sql)
End Sub

Public Sub UpdateCategory(id As String, name As String, status As String)
    sql = "Update Category set Name = '" & name & "', Status = " & status & " where Id = " & id
    conn.ExcCommand (sql)
End Sub

Public Sub EnabledDisabledCategory(id As String)
    sql = "Update Category set Name = 'daneil', Status = (case when Status = 1 then 0 else 1 end) where Id = " & id
    conn.ExcCommand (sql)
End Sub


Public Function ExistsCategory(id As Integer, name As String) As ADODB.Recordset
    sql = "Select * from Category where id <> " & id & " and Name = '" & name & "'"
    Set ExistsCategory = conn.GetDataTable(sql)
End Function


Public Function GetId(name As String) As Integer
    sql = "Select * from Category where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    GetId = result.Fields.Item(0).Value
End Function
