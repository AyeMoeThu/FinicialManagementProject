VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGL 
      Caption         =   "GL Account"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGL_Click()
    frmGLAccountSetup.Show 1
End Sub
