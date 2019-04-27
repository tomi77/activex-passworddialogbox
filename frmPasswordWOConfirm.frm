VERSION 5.00
Begin VB.Form frmPasswordWOConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter password :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmPasswordWOConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public booCancel As Boolean
Public strPassword As String

Private Sub cmdCancel_Click()

  booCancel = True
  strPassword = ""
  frmPasswordWOConfirm.Hide

End Sub

Private Sub cmdOK_Click()

  booCancel = False
  strPassword = txtPassword(0).Text
  frmPasswordWOConfirm.Hide

End Sub

