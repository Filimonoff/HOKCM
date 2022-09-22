Private Sub CommandButton1_Click()
SetChainNameForm.Hide
If SetChainNameForm.ComboBox1.Value = "АТБ" Then
Call ATBfinesComb
ElseIf SetChainNameForm.ComboBox1.Value = "Fozzy" Then
Call FOZZYfinesComb
ElseIf SetChainNameForm.ComboBox1.Value = "Сільпо" Then
Call SILPOfinesComb
ElseIf SetChainNameForm.ComboBox1.Value = "Novus" Then
Call NOVUSfinesComb
End If
End Sub
