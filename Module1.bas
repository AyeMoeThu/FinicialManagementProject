Attribute VB_Name = "Common"
'Declartion as use in the project for connection

Public mCon As New ADODB.Connection
Public mRsFind As New ADODB.Recordset
Public mRsUser As New ADODB.Recordset
Public mRsUpdate As New ADODB.Recordset
Public mRsDelete As New ADODB.Recordset
Public mADOConString As String



Public Function OpenConn()
    Set mCon = New ADODB.Connection
    mCon.ConnectionString = "Connection2"
    
End Function



