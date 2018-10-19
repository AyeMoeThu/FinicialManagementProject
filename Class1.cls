VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function openConnection() As Boolean

 Dim l_ConnString As String
 Dim l_Sqlcmd As String
 
 Set l_ConnString = "Provider = MSDASQL; driver ={SQL Server}; database=db; server=localhost\sqlexpress;uid=sa;pwd=123;"
 

End Function
