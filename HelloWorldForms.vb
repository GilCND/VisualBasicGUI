'Program name: Hello World Forms
'Objective: This program is the first one using forms and better GUI in VB
'Programmer: Felipe SG

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnClickThis.Click
        lblHelloWorld.Text = "Hello World and Dave!"
    End Sub
End Class
