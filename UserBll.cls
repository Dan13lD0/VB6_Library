VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserBll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dal  As New UserDal
Dim dalLogin As New LoginDal
Public Function ValidateUser(login As String, password As String) As Boolean
ValidateUser = dalLogin.ValidateUser(login, password)
End Function

Public Function GetUsers(code As String, name As String, login As String, email As String, status As String) As ADODB.Recordset

If code = "" Then
code = "0"
End If

Dim vStatus As String

If status = "All" Or status = "" Then
vStatus = ""
ElseIf status = "Active" Then
vStatus = "1"
ElseIf status = "Inactive" Then
vStatus = "0"
End If

Set GetUsers = dal.GetUsers(code, name, login, email, vStatus)

End Function

Public Function GetUser(code As String) As ADODB.Recordset
If code = "" Then
code = "0"
End If

Set GetUser = dal.GetUser(code)

End Function


Public Sub CreateUpdateUser(code As String, name As String, login As String, password As String, email As String, status As String)
    If code = "" Then
     dal.CreateUser name, login, password, email
    Else
     dal.UpdateUser code, name, login, password, email, IIf(status = "Active", 1, 0)
    End If
End Sub

Public Sub DeleteUser(code As String)
If code = "" Then
code = "0"
End If

dal.EnableDisableUser (code)
End Sub


Public Sub InitializeUser()
Dim exist As New ADODB.Recordset

Set exist = dal.GetUsers("0", "Administrator", "admin", "", "1")
If exist.RecordCount = 0 Then
dal.CreateUser "Administrator", "admin", "admin", "admin@gmail.com"
End If

End Sub

