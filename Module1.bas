Attribute VB_Name = "Module1"
'Declartion as use in the project for connection

Public mCon As New ADODB.Connection
Public mRsFind As New ADODB.Recordset
Public mRsUser As New ADODB.Recordset
Public mRsUpdate As New ADODB.Recordset
Public rsdelete As New ADODB.Recordset


Public Function OpenConn()
    Set mCon = New ADODB.Connection
    mCon.ConnectionString = ""
    
End Function



