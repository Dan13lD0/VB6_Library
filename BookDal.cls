VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BookDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetBooks(name As String, category As String, author As String, publisher As String, status As String) As ADODB.Recordset
sql = "Select Id,Name,Category =(select top 1 name from Category where Id = CategoryId), Author=(select top 1 name from Author where Id = AuthorId ) ,Publisher = (select top 1 name from Publisher  where Id = PublisherId),Status=(case when Status = 1 then 'Active' else 'Inactive' end) from Book where Name like '%" & name & "%'and CategoryId in (select id from Category c where c.Name like '%" & category & "%') and AuthorId in (select id from Author a where a.Name like '%" & author & "%')  and PublisherId in (select id from Publisher p where p.Name like '%" & publisher & "%')  and Status = (case when '" & status & "' ='' then Status else '" & status & "' end)"
Set GetBooks = conn.GetDataTable(sql)
End Function

Public Function GetBook(id As String) As ADODB.Recordset
sql = "Select Id,Name,Category =(select top 1 name from Category where Id = CategoryId), Author=(select top 1 name from Author where Id = AuthorId ) ,Publisher = (select top 1 name from Publisher  where Id = PublisherId),Status=(case when Status = 1 then 'Active' else 'Inactive' end)  from Book where  id = " & id
Set GetBook = conn.GetDataTable(sql)
End Function

Public Function CreateBook(name As String, description As String, author As String, category As String, publisher As String)
sql = "Insert Into book(Name,Description,AuthorId,CategoryId,PublisherId,Status) values('" & name & "','" & description & "'," & author & "," & category & "," & publisher & ",1)"
conn.ExcCommand (sql)
End Function

Public Function UpdateBook(id As String, name As String, description As String, author As String, category As String, publisher As String, status As String)
sql = "Update book set Name = '" & name & "',Description = '" & description & "',AuthorId =" & author & ",CategoryId=" & category & ",PublisherId=" & publisher & ",Status=" & status & " where id =" & id
conn.ExcCommand (sql)
End Function

Public Function EnableDisableBook(id As String)
sql = "Update book set Status=(case when status =0 then 1 else 0 end) where id =" & id
conn.ExcCommand (sql)
End Function


Public Function GetId(name As String) As Integer
    sql = "Select * from book where Name = '" & name & "'"
    Dim result As New ADODB.Recordset
    
    Set result = conn.GetDataTable(sql)
    
    GetId = result.Fields.Item(0).Value
End Function
