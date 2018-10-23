VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAOGLAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GLAccount As New InfoGLAccount
Private SqlQry As String
Private Rff As Integer   'Row Affected

Function SaveRecord() As Boolean

    GetConnection
    OpenConn mCon, 1000, mADOConString
    
    SqlQry = "Insert Into GLAccount(AccountNumber,AccountName,AccountType,Remarks)" _
           & " Values ('" & GLAccount.AccountNumber & "', '" & GLAccount.AccountName & "', '" & GLAccount.AccountType & "', '" + GLAccount.Remarks & "')"
    
    mCon.Execute SqlQry, Rff
    
    If Rff > 0 Then
        SaveRecord = True
    Else
        SaveRecord = False
    End If
    
    CutConnection mCon, True
    
End Function

