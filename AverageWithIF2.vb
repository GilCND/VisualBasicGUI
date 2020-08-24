Program Name: Average with If 2
Objective: Calculate the average using if statements exercvise 2. 
Programmer: Felipe SG

Public Class Form1

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub btn_calculate_Click(sender As Object, e As EventArgs) Handles btn_calculate.Click
        'Clear all variables from possible previously interactions. 
        lb_result.Text = ""
        lb_Fail.Text = ""
        lb_warning.Text = ""
        lb_rewrite.Text = ""
        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim num3 As Decimal
        Dim result As Decimal

'Check for empty values. 
        If (txt_num1.Text = "" Or txt_num2.Text = "" Or txt_num3.Text = "") Then
            lb_warning.Text = "ERROR, Empty values"
        ElseIf Not (IsNumeric(txt_num1.Text)) Then
            lb_warning.Text = "ERROR, Invalid Characters"
        ElseIf Not (IsNumeric(txt_num2.Text)) Then
            lb_warning.Text = "ERROR, Invalid Characters"
        ElseIf Not (IsNumeric(txt_num3.Text)) Then
            lb_warning.Text = "ERROR, Invalid Characters"

        Else
'If is not empty check if it is not less than 0 and greater than 100
            If (txt_num1.Text < 0 Or txt_num1.Text > 100 Or txt_num2.Text < 0 Or txt_num2.Text > 100 Or txt_num3.Text < 0 Or txt_num3.Text > 100) Then
                lb_warning.Text = "Only Values between 0 and 100"
           
'All good? assign those values to variables and calculate           
           Else
                num1 = txt_num1.Text
                num2 = txt_num2.Text
                num3 = txt_num3.Text
                result = (num1 + num2 + num3) / 3
                
'If the result is less or equal to 59, but onde of the values are greater or equal to 95 the student is eligible to rewrite.               
               If (result <= 59) Then
                    lb_Fail.Text = ($" The average is {result.ToString("n2")}")
                    If (num2 >= 95 Or num3 >= 95) Then
                        lb_rewrite.Text = ("Your are eligible to re-write Grade 1")
                    ElseIf (num1 >= 95 Or num3 >= 95) Then
                        lb_rewrite.Text = ("Your are eligible to re-write Grade 2")
                    ElseIf (num1 >= 95 And num2 >= 95) Then
                        lb_rewrite.Text = ("Your are eligible to re-write Grade 3")
                    End If
                Else

'If none of the grades are greater or equal to 95, too bad, print the result =(

                    lb_result.Text = ($" The average is {result.ToString("n2")}")
                End If
            End If
        End If

    End Sub

'Button to clear all values
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_clear.Click
        txt_num1.Text = ""
        txt_num2.Text = ""
        txt_num3.Text = ""
        lb_result.Text = ""
        lb_Fail.Text = ""
        lb_warning.Text = ""
        lb_rewrite.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
