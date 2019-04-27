Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
  Dim dlg As New PasswordDialogBox.PasswordDlg
  dlg.LabelPassword = "Podaj has³o :"
  dlg.LabelConfirmPassword = "Powtórz has³o :"
  dlg.ErrorConfirmMessage = "Has³a siê nie zgadzaj¹ !"
  dlg.ErrorConfirmCaption = "B³¹d"
  dlg.ShowDialog "W/O Confirm"
  If Not dlg.Cancel Then _
    MsgBox dlg.Password
  dlg.ShowDialog "W Confirm", True
  If Not dlg.Cancel Then _
    MsgBox dlg.Password
End Sub
