VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BorrowDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String



Public Function GetBorrows(client As String, book As String) As ADODB.Recordset
sql = "select Code = b.Id,Book = (select top 1 name from Book where id = b.BookId),Client = (Select Name from client where id =b.ClientId),b.DateBorrow ,b.DateReturn, Status = case when Status = 1 then 'Active' else 'Inactive' end from Borrow b where ClientId in (select id from Client where Name like '%" & client & "%') and BookId in (select id from Book where Name like '%" & book & "%') and Status = 1 "
Set GetBorrows = conn.GetDataTable(sql)
End Function

Public Function GetBooksForBorrow(name As String, category As String, author As String, publisher As String) As ADODB.Recordset
sql = "Select Code = b.id,Name,Category = (Select Name from Category where id = b.CategoryId),Author = (Select Name from Author where id = b.AuthorId),Publisher = (Select Name from Publisher where id = b.PublisherId),Status = case when (Select count(*) from Borrow where BookId = b.id and  Status = 1) = 0 then 'Available' else 'Unvaliable' end from Book b where Name like '%" & name & "%' and CategoryId in (Select id from Category where Name like '%" & category & "%') and AuthorId in (Select id from Author where Name like '%" & author & "%') and PublisherId in (Select id from Publisher where Name like '%" & publisher & "%')"
Set GetBooksForBorrow = conn.GetDataTable(sql)
End Function


Public Function GetBorrow(id As String)
sql = "Select * from Borrow where id = " & id
conn.ExcCommand (sql)
End Function


Public Function GetBorrowByBookId(bookId As String)
sql = "Select * from Borrow where BookId  = " & bookId
conn.ExcCommand (sql)
End Function


Public Function CreateBorrow(book As String, client As String, dateborrow As String, datereturn As String)
sql = "Insert Into Borrow(BookId,ClientId,DateBorrow,DateReturn,Status) values (" & book & "," & client & ",'" & dateborrow & "','" & datereturn & "',1)"
conn.ExcCommand (sql)
End Function

Public Function UpdateBorrow(id As String, book As String, client As String, dateborrow As String, datereturn As String, status As String)
sql = "Update Borrow set BookId=" & book & ",ClientId=" & client & ",DateBorrow='" & dateborrow & "',DateReturn='" & datereturn & "',Status=" & status & " where id =" & id
conn.ExcCommand (sql)
End Function

Public Function EnableDisableBorrow(id As String)
sql = "Delete from Borrow where BookId =" & id
conn.ExcCommand (sql)
End Function

Public Function BorrowChangeStatus(id As String)
sql = "Update Borrow set Status = (case when Status = 0 then 1 else 0 end) where id = " & id
conn.ExcCommand (sql)
End Function
