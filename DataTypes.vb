'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Data Types
'Date: 30/09/2019
'Autor: Felipe Souza Gil
'Operation: This program will wait the user to imput data in order to calculate average
'If the user input ivalid characters, the program should inform it 
' Ivalid characters are: 
' Any letter
' Negative numbers
' Number bigger tham 100 
' Check the eligibility based on the rule: If one of the numbers are 95 or highter and the average
' are below or equal to 60, you are eligible to retest 1 score. 

'The program should print on the label:       
'           Any error 
'           Average
'           Eligibility based in one rule

'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'09/10/2019         Felipe Souza Gil        Initial Version
'16/10/2019         Felipe Souza Gil        Code commented and implemented better if statement.




Public Class Form1

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub btn_calculate_Click(sender As Object, e As EventArgs) Handles btn_calculate.Click
        lb_result.Text = ""
        lb_Fail.Text = ""
        lb_warning.Text = ""
        lb_rewrite.Text = ""
        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim num3 As Decimal
        Dim result As Decimal

        ' This first part validate the data entry.
        If (txt_num1.Text = "" Or txt_num2.Text = "" Or txt_num3.Text = "") Then
            lb_warning.Text = "ERROR, Empty values"
        ElseIf Not (IsNumeric(txt_num1.Text) And (IsNumeric(txt_num2.Text) And (IsNumeric(txt_num3.Text)))) Then
            lb_warning.Text = "ERROR, Invalid Characters"
        ElseIf (txt_num1.Text < 0 Or txt_num1.Text > 100 Or txt_num2.Text < 0 Or txt_num2.Text > 100 Or txt_num3.Text < 0 Or txt_num3.Text > 100) Then
            lb_warning.Text = "Only Values between 0 and 100"
            '------------------------------------------------------------------------------------------------------------------------------------------

            ' if all values are validated and can be saved, the program will assign them to a variable
        Else
            num1 = txt_num1.Text
            num2 = txt_num2.Text
            num3 = txt_num3.Text
            result = (num1 + num2 + num3) / 3
            '------------------------------------------------------------------------------------------------------------------------------------------


            If (result <= 59.99 And (num1 >= 95 Or num2 >= 95 Or num3 >= 95)) Then
                lb_Fail.Text = ($" The average is {result.ToString("n2")}")
                lb_rewrite.Text = ("Your are eligible to re-write ")
            ElseIf (result >= 60) Then
                lb_result.Text = ($" The average is {result.ToString("n2")}")
            Else
                lb_Fail.Text = ($" The average is {result.ToString("n2")}")
            End If
        End If

    End Sub

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
