VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PasswordDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strPassword As String
Private booCancel As Boolean
Private strLabelPassword As String
Private strLabelConfirmPassword As String
Private strErrorConfirmMessage As String
Private strErrorConfirmCaption As String

Public Sub ShowDialog(Caption As String, Optional WithConfirm As Boolean = False)

  Dim frmForm As Object

  If WithConfirm Then
    Set frmForm = frmPasswordWConfirm
  Else
    Set frmForm = frmPasswordWOConfirm
  End If
  With frmForm
    .Caption = Caption
    .lblPassword(0).Caption = strLabelPassword
    If WithConfirm Then
      .lblPassword(1).Caption = strLabelConfirmPassword
      .strErrorConfirmMessage = strErrorConfirmMessage
      .strErrorConfirmCaption = strErrorConfirmCaption
    End If
    .Show vbModal
    strPassword = .strPassword
    booCancel = .booCancel
  End With
  Unload frmForm
  Set frmForm = Nothing

End Sub

Public Property Get Password() As String

  Password = strPassword

End Property

Public Property Get Cancel() As Boolean

  Cancel = booCancel

End Property


Public Property Get LabelPassword() As String

  LabelPassword = strLabelPassword

End Property

Public Property Let LabelPassword(ByVal NewValue As String)

  strLabelPassword = NewValue

End Property

Public Property Get LabelConfirmPassword() As String

  LabelConfirmPassword = strLabelConfirmPassword

End Property

Public Property Let LabelConfirmPassword(ByVal NewValue As String)

  strLabelConfirmPassword = NewValue

End Property

Private Sub Class_Initialize()

  strLabelPassword = "Enter password :"
  strLabelConfirmPassword = "Confirm password :"
  strErrorConfirmMessage = "Password do not match !"
  strErrorConfirmCaption = "Error"

End Sub



Public Property Get ErrorConfirmMessage() As String

  ErrorConfirmMessage = strErrorConfirmMessage

End Property

Public Property Let ErrorConfirmMessage(ByVal NewValue As String)

  strErrorConfirmMessage = NewValue

End Property

Public Property Get ErrorConfirmCaption() As String

  ErrorConfirmCaption = strErrorConfirmCaption

End Property

Public Property Let ErrorConfirmCaption(ByVal NewValue As String)

  strErrorConfirmCaption = NewValue

End Property
