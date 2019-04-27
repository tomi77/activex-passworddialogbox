Option Explicit

Dim dlg
set dlg = CreateObject("PasswordDialogBox.PasswordDlg")
With dlg
  .LabelPassword = "Podaj has³o :"
  .LabelConfirmPassword = "Powtórz has³o :"
  .ErrorConfirmMessage = "Has³a siê nie zgadzaj¹ !"
  .ErrorConfirmCaption = "B³¹d"
  .ShowDialog "W/O Confirm"
  If Not .Cancel Then _
    MsgBox .Password
  .ShowDialog "W Confirm", True
  If Not .Cancel Then _
    MsgBox .Password
End With
Set dlg = Nothing
