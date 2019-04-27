VERSION 5.00
Begin VB.Form frmPasswordWConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1815
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
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
   Begin VB.Label lblPassword 
      Caption         =   "Confirm password :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
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
Attribute VB_Name = "frmPasswordWConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public booCancel As Boolean
Public strPassword As String
Public strErrorConfirmMessage As String
Public strErrorConfirmCaption As String

Private Sub cmdCancel_Click()

  booCancel = True
  strPassword = ""
  frmPasswordWConfirm.Hide

End Sub

Private Sub cmdOK_Click()

  If txtPassword(0).Text = txtPassword(1).Text Then
    booCancel = False
    strPassword = txtPassword(0).Text
    frmPasswordWConfirm.Hide
  Else
    MsgBox strErrorConfirmMessage, vbOKOnly + vbExclamation, strErrorConfirmCaption
    txtPassword(0).Text = ""
    txtPassword(1).Text = ""
    txtPassword(0).SetFocus
  End If

End Sub

