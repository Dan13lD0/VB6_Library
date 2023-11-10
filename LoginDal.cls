VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginDal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New SqlServerContext
Dim sql As String

Public Function ValidateUser(user As String, password As String) As Boolean
    sql = "Select * from [LibraryDev].[dbo].[User] where Login = '" & user & "' and Password = '" & password & "' and Status=1"
    Dim vali As New ADODB.Recordset
    Set vali = conn.GetDataTable(sql)
    
    ValidateUser = IIf(vali.RecordCount > 0, True, False)

End Function
