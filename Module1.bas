Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
  Dim dlg As New PasswordDialogBox.PasswordDlg
  dlg.LabelPassword = "Podaj has�o :"
  dlg.LabelConfirmPassword = "Powt�rz has�o :"
  dlg.ErrorConfirmMessage = "Has�a si� nie zgadzaj� !"
  dlg.ErrorConfirmCaption = "B��d"
  dlg.ShowDialog "W/O Confirm"
  If Not dlg.Cancel Then _
    MsgBox dlg.Password
  dlg.ShowDialog "W Confirm", True
  If Not dlg.Cancel Then _
    MsgBox dlg.Password
End Sub
