Option Explicit

Dim dlg
set dlg = CreateObject("PasswordDialogBox.PasswordDlg")
With dlg
  .LabelPassword = "Podaj has�o :"
  .LabelConfirmPassword = "Powt�rz has�o :"
  .ErrorConfirmMessage = "Has�a si� nie zgadzaj� !"
  .ErrorConfirmCaption = "B��d"
  .ShowDialog "W/O Confirm"
  If Not .Cancel Then _
    MsgBox .Password
  .ShowDialog "W Confirm", True
  If Not .Cancel Then _
    MsgBox .Password
End With
Set dlg = Nothing
