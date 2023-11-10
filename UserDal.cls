VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String



Public Function GetUsers(id As String, name As String, login As String, email As String, status As String) As ADODB.Recordset
sql = "Select Id,Name,Login,Email,Status = (case when Status = 1 then 'Active' else 'Inactive' end) from [LibraryDev].[dbo].[User]  where id=(case when '" & id & "' ='0' then id else '" & id & "' end) and Name like '%" & name & "%' and Login like '%" & login & "%'and Email like '%" & email & "%' and Status = (case when '" & status & "' ='' then status else '" & status & "' end)"
Set GetUsers = conn.GetDataTable(sql)
End Function

Public Function GetUser(id As String) As ADODB.Recordset
sql = "Select * from [LibraryDev].[dbo].[User]  where id=" & id
Set GetUser = conn.GetDataTable(sql)
End Function


Public Sub CreateUser(name As String, login As String, password As String, email As String)
sql = "Insert Into [LibraryDev].[dbo].[User](Name,Login,Password,Email,Status) values ('" & name & "','" & login & "','" & password & "','" & email & "',1)"
conn.ExcCommand (sql)
End Sub


Public Sub UpdateUser(id As String, name As String, login As String, password As String, email As String, status As String)
sql = "Update [LibraryDev].[dbo].[User] set Name='" & name & "',Login='" & login & "',Password='" & password & "',Email='" & email & "',Status=" & status & " where id =" & id
conn.ExcCommand (sql)
End Sub


Public Sub EnableDisableUser(id As String)
sql = "Update [LibraryDev].[dbo].[User] set Status=(case when status =0 then 1 else 0 end) where id =" & id
conn.ExcCommand (sql)
End Sub

Public Sub ExistsUser(id As String, login As String)
sql = "Select * from [LibraryDev].[dbo].[User] where  id <> " & id & " and Login = '" & login & "'"
conn.ExcCommand (sql)
End Sub
