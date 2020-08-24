Program Name: Average With IF
Objective: Create an user interface to calculate average using if statements.
Programmer: Felipe SG


Public Class Form1

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub btn_calculate_Click(sender As Object, e As EventArgs) Handles btn_calculate.Click
        lb_result.Text = ""
        lb_Fail.Text = ""
        lb_warning.Text = ""
        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim num3 As Decimal
        Dim result As Decimal
        
'Check if the boixes are empty, If so, show an label informing the error.
        If (txt_num1.Text = "" Or txt_num2.Text = "" Or txt_num3.Text = "") Then
            lb_warning.Text = "ERROR, Empty values"
        ElseIf (IsNumeric(txt_num1.Text) And IsNumeric(txt_num2.Text) And IsNumeric(txt_num3.Text)) Then
            If (txt_num1.Text < 0 Or txt_num1.Text > 100 Or txt_num2.Text < 0 Or txt_num2.Text > 100 Or txt_num3.Text < 0 Or txt_num3.Text > 100) Then
                lb_warning.Text = "Only Values between 0 and 100"
                
 'If there is no error, assign the values and calulate the result.
           Else
                num1 = txt_num1.Text
                num2 = txt_num2.Text
                num3 = txt_num3.Text
                result = (num1 + num2 + num3) / 3
'If the result is less or equal to 60 Use label in red
                If (result >= 60) Then
                    lb_result.Text = ($" The average is {result.ToString("n2")}")
'If the average is greater then 60 show in a green label
'Here I have 2 labels, one red and one green.
                Else
                    lb_Fail.Text = ($" The average is {result.ToString("n2")}")

                End If

            End If
        Else
            lb_warning.Text = "ERROR, Invalid Characters"

        End If

    End Sub
    
'To clear the labels I have a button clear. 
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_clear.Click
        txt_num1.Text = ""
        txt_num2.Text = ""
        txt_num3.Text = ""
        lb_result.Text = ""
        lb_Fail.Text = ""
        lb_warning.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class
