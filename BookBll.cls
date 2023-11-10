VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BookBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New BookDal
Dim dalAutor As New AuthorDal
Dim dalCategory As New CategoryDal
Dim dalPublisher As New PublisherDal

Public Function GetBooks(name As String, category As String, author As String, publish As String, status As String) As ADODB.Recordset
 If code = "" Then
 code = "0"
 End If
 
 category = Replace(category, "All", "")
 author = Replace(author, "All", "")
 publish = Replace(publish, "All", "")
 
 If status <> "All" And status <> "" Then
    If status = "Active" Then
    status = "1"
    Else
    status = "0"
    End If
 Else
 status = ""
 End If

 Set GetBooks = dal.GetBooks(name, category, author, publish, status)
End Function

 Public Function GetBook(code As String) As ADODB.Recordset
 If code = "" Then
 code = "0"
 End If
 Set GetBook = dal.GetBook(code)
 End Function

 Public Sub CreateUpdateBook(code As String, name As String, category As String, author As String, publisher As String, description As String, status As String)
 
 category = dalCategory.GetId(category)
 author = dalAutor.GetId(author)
 publisher = dalPublisher.GetId(publisher)
 
 If code = "" Then
    dal.CreateBook name, description, author, category, publisher
 Else
    dal.UpdateBook code, name, description, author, category, publisher, IIf(status = "Active", "1", "0")
 End If
 End Sub

 Public Sub DeleteBook(code As String)
  If code = "" Then
 code = "0"
 End If
 dal.EnableDisableBook (code)
 End Sub
