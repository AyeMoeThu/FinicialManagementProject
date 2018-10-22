Attribute VB_Name = "General"
'Declartion as use in the project for connection
Option Explicit

Public mCon As New ADODB.Connection
Public mRsFind As New ADODB.Recordset
Public mRsUser As New ADODB.Recordset
Public mRsUpdate As New ADODB.Recordset
Public mRsDelete As New ADODB.Recordset
Public mADOConString As String

Function GetConnection()

     mADOConString = "ODBC;DRIVER={SQL Server};SERVER=" + gsLSN + _
                ";UID=" + gsLID + ";PWD=" + gsLPW + _
                ";APP=Visual Basics;Database=" + gsLDN + _
                ";QueryLog_On=Yes;Time=900000"

End Function

Function OpenConn(Conn As ADODB.Connection, Tm As Double, ConStr As String)
    If Conn.State = 0 Then
        Conn.Open ConStr
        Conn.CommandTimeout = Tm
        Conn.Execute "Set DateFormat dmy"
    End If
    
    If Conn.State <> adStateClosed Then
        Conn.Execute "Set DateFormat dmy"
    End If
End Function

Function CutConnection(Conn As ADODB.Connection, Optional CutConn As Boolean)
    If CutConn Then
       If Conn.State = 1 Then
        Conn.Close
       End If
    End If
End Function

Function InformMessage(MsgPrompt As String, MsgTitle As String, Optional Typ As Integer) As Boolean
    If Len(MsgTitle) = 0 Then
        MsgTitle = "Error Message"
    End If
    
    If Len(MsgPrompt) = 0 Then
        MsgPrompt = " & Err.Number & " + " ----> " + Err.Description
        MsgBox MsgPrompt, vbOKOnly, MsgTitle
    End If
    
    InformMessage = True
    
    Select Case Typ
    Case 1
        MsgBox MsgPrompt, vbInformation, MsgTitle
    Case 2
        MsgBox MsgPrompt, vbCritical, MsgTitle
    Case 3
        MsgBox MsgPrompt, vbQuestion, MsgTitle
    Case 4
        MsgBox MsgPrompt, vbExclamation, MsgTitle
    Case 5
        If MsgBox(MsgPrompt, vbInformation + vbYesNo, MsgTitle) = vbNo Then
            InformMessage = False
        End If
    Case 6
        If MsgBox(MsgPrompt, vbCritical + vbYesNo, MsgTitle) = vbNo Then
            InformMessage = False
        End If
    Case 7
        If MsgBox(MsgPrompt, vbQuestion + vbYesNo, MsgTitle) = vbNo Then
            InformMessage = False
        End If
    Case 8
        If MsgBox(MsgPrompt, vbExclamation + vbYesNo, MsgTitle) = vbNo Then
            Information = False
        End If
    Case Else
        MsgBox MsgPrompt, vbInformation, MsgTitle
    End Select
End Function

Sub Wait()
    Screen.MousePointer = 0
End Sub

Sub EndWait()
    Screen.MousePointer = 0 'comment
End Sub







