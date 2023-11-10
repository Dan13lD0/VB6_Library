VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function GetAddress(id As String) As Recordset
sql = "Select * from Address  where id = " & id
Set GetAddress = conn.GetDataTable(sql)
End Function

Public Function CreateAddress(zipCode As String, street As String, complement As String, city As String, state As String)
sql = "Insert into Address(ZipCode,Street,Complement,City,State) values ('" & zipCode & "','" & street & "','" & complement & "','" & city & "','" & state & "')"
conn.ExcCommand (sql)
End Function

Public Function UpdateAddress(id As String, zipCode As String, street As String, complement As String, city As String, state As String)
sql = "Update Address set ZipCode = '" & zipCode & "',Street = '" & street & "',Complement = '" & complement & "', City = '" & city & "',State = '" & state & "' where id = " & id
conn.ExcCommand (sql)
End Function

Public Function DeleteAddress(id As String)
sql = "Delete from Address where id = " & id
conn.ExcCommand (sql)
End Function

Public Function GetLastId() As ADODB.Recordset
sql = "SELECT TOP (1) id FROM Address order by Id  desc"
Set GetLastId = conn.GetDataTable(sql)
End Function

