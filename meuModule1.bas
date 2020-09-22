Attribute VB_Name = "Module1"
Private Sub Main()
Dim x() As Byte
x = LoadResData(101, "CUSTOM")
Open App.Path & "\Teste.jpg" For Binary As 1
Put #1, 1, x
Close 1

Form1.Show
End Sub
