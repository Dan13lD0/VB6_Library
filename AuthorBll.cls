VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AuthorBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal As New AuthorDal

Public Function GetAuthors(code As String, name As String, status As String) As ADODB.Recordset
    Dim vStatus As String
    
    If Trim(code) = "" Then
    code = "0"
    End If
    
    If Trim(status) = "" Or Trim(status) = "All" Then
    vStatus = ""
    ElseIf Trim(status) = "Active" Then
    vStatus = "1"
    ElseIf Trim(status) = "Inactive" Then
    vStatus = "0"
    End If
    
    Set GetAuthors = dal.GetAuthors(code, name, vStatus)
End Function

Public Function GetAuthor(code As String) As ADODB.Recordset
    If Trim(code) = "" Then
    code = "0"
    End If
    Set GetAuthor = dal.GetAuthor(code)
End Function

Public Sub CreateUpdateAuthor(code As String, name As String, status As String)
    If Trim(code) = "" Then
        dal.CreateAuthor (name)
    Else
        dal.UpdateAuthor code, name, IIf(status = "Active", "1", "0")
    End If
End Sub

Public Sub DeleteAuthor(code As String)
    If Trim(code) = "" Then
    code = "0"
    End If
    dal.EnabledDisabledAuthor (code)
End Sub

