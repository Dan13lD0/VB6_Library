VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String


Public Function GetContact(id As String) As ADODB.Recordset
sql = "Select * from Contact where id = " & id
Set GetContact = conn.GetDataTable(sql)
End Function

Public Sub CreateContact(typeContact As String, contact As String, person As String)
sql = "Insert Into Contact (Type,Contact,Person) values ('" & typeContact & "','" & contact & "','" & person & "')"
conn.ExcCommand (sql)
End Sub


Public Sub UpdateContact(id As String, typeContact As String, contact As String, person As String)
sql = "Update Contact set Type = '" & typeContact & "',Contact='" & contact & "',Person='" & person & "' where id = " & id
conn.ExcCommand (sql)
End Sub


Public Sub RemoveContact(id As String)
sql = "Delete from Contact where id =" & id
conn.ExcCommand (sql)
End Sub

Public Function GetLastId() As ADODB.Recordset
sql = "SELECT TOP (1) id FROM Contact order by Id  desc"
Set GetLastId = conn.GetDataTable(sql)
End Function

