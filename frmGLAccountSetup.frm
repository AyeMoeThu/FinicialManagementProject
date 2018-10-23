VERSION 5.00
Begin VB.Form frmGLAccountSetup 
   Caption         =   "GL Account Setup"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtRemark 
         Height          =   735
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtAccountName 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox cboAccountType 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtAccountNumber 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblAccName 
         Caption         =   "Account Name"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblAccountType 
         Caption         =   "Account Type"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblAccountNumber 
         Caption         =   "Account Number"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGLAccountSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSave_Click()
    Dim mDAOGlAccount As New DAOGLAccount
    
    With mDAOGlAccount.GLAccount
         .ID = 0
         .AccountNumber = Trim(frmGLAccountSetup.txtAccountNumber.Text)
         .AccountName = Trim(frmGLAccountSetup.txtAccountName.Text)
         .AccountType = 0 'frmGLAccountSetup.cboAccountType.ListIndex
         .Remarks = Trim(frmGLAccountSetup.txtRemark.Text)
    End With
    
        mDAOGlAccount.SaveRecord
        
End Sub

