```
Private Sub BTLogin_Click()
Dim Username As String, Password As String
Dim MyUsername As String, MyPassword As String

Username = TBUsername.Text
Password = TBPassword.Text
MyUsername = "admin"
MyPassword = "123"

If Username = Empty Then
    MsgBox "Silahkan Isi Username", vbInformation, "Peringatan"
    TBUsername.SetFocus
ElseIf Password = Empty Then
    MsgBox "Silahkan Isi Password", vbInformation, "Peringatan"
    TBPassword.SetFocus
    
ElseIf Username = MyUsername And MyPassword = Password Then
    MsgBox "Login Berhasil", vbInformation, "Lgin"
Else
    MsgBox "Username atau Password Salah", vbCritical, "Peringatan"
End If



End Sub

Private Sub CKPassword_Click()
If CKPassword Then
    TBPassword.PasswordChar = ""
Else
    TBPassword.PasswordChar = "0"
End If
    
End Sub


Private Sub UserForm_Click()

End Sub
```
