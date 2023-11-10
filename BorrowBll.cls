VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BorrowBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New BorrowDal
Dim dalcategory As New CategoryDal
Dim dalauthor As New AuthorDal
Dim dalpublisher As New PublisherDal
Dim dalclient As New ClientDal
Dim dalbook As New BookDal

Public Function GetBorrows(book As String, category As String, author As String, publisher As String) As ADODB.Recordset
category = Replace(category, "All", "")
author = Replace(author, "All", "")
publisher = Replace(publisher, "All", "")
Set GetBorrows = dal.GetBooksForBorrow(book, category, author, publisher)
End Function

Public Function GetBorrow(code As String) As ADODB.Recordset
If code = "" Then
code = "0"
End If
Set GetBorrow = dal.GetBorrow(code)
End Function

Public Function GetBorrows2(client As String, book As String) As ADODB.Recordset
   Set GetBorrows2 = dal.GetBorrows(client, book)
End Function

Public Sub CreateUpdateBorrow(book As String, client As String, dateborrow As String, datereturn As String)
dal.CreateBorrow book, client, dateborrow, datereturn
End Sub

Public Sub DeleteBorrow(code As String)
If code = "" Then
code = "0"
End If

dal.EnableDisableBorrow (code)
End Sub

Public Sub BorrowChangeStatus(id As String)

If id = "" Then
id = "0"
End If

dal.BorrowChangeStatus (id)

End Sub
