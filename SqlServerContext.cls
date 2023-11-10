VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlServerContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetDataTable(sql As String) As Recordset
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

con.Open "Provider=MSOLEDBSQL;Server=(local);Database=LibraryDev;UID=sa;PWD=123456;"
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
 rs.Open sql, con

Set GetDataTable = rs

 
End Function

Public Function ExcCommand(sql As String)
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

con.Open "Provider=MSOLEDBSQL;Server=(local);Database=LibraryDev;UID=sa;PWD=123456;"
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
rs.Open sql, con

 
End Function
