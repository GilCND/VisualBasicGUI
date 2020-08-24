'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Loan Calculator WEith Exceptions
'Date: 18/11/2019
'Autor: Felipe Souza Gil
' This program should calculate the loan using 3 fields
'Loan Amount
'Interest Rate
'Loan period

'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'25/09/2019         Felipe Souza Gil        Initial Version
'18/11/2019         Felipe Souza Gil        Using Exceptions



Public Class Form1
    'R = Interest rate
    Dim R As Decimal
    Dim R1 As Decimal
    'P = Principal Amount of the Loan
    Dim P As Decimal
    'N = Total months for the Loan
    Dim N As Decimal
    'result1 = fisrt part of the formula (r*p)
    Dim Result1 As Decimal
    'result2 = Seccond part of the formula (1-(1+R)^N
    Dim Result2 As Decimal
    'result3 = result 1 divided by result 2
    Dim Result3 As Decimal
    'Final = result 3 multiplied by N
    Dim Final As Decimal
    'Its the difference paid on the end of contract
    Dim Dif As Decimal
    Const Negative_Error = "ERROR - Please Inform a positive number"
    Dim Errormsg As String
    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub btn_calculate_Click(sender As Object, e As EventArgs) Handles btn_calculate.Click
        'Exception code starts here
        Try
            If Txt_Loan.Text < 0 Then
                Txt_Loan.Focus()
                Errormsg = "Field Loan"
                Throw New Exception
            ElseIf Txt_Period.Text < 0 Then
                Txt_Period.Focus()
                Errormsg = "Field Period"
                Throw New Exception
            ElseIf Txt_Rate.Text < 0 Then
                Txt_Rate.Focus()
                Errormsg = "Field Rate"
                Throw New Exception
            End If

            P = Txt_Loan.Text
            R1 = Txt_Rate.Text
            N = Txt_Period.Text
            R = ((R1 / 100) / 12)
            Result1 = (R * P)
            Result2 = 1 - (1 + R) ^ (-N)
            Result3 = Result1 / Result2
            Final = Result3 * N
            Dif = (Final - P)

            lb_InterestRate.Text = ($"{R1.ToString("n2")} %")
            lb_LoanAmount.Text = ($"${P}")
            lb_LoanPMonths.Text = ($"{N} Months")
            lb_result.Text = ($"Total cost Of the Loan  {Final.ToString("n2")}")
            Lb_Warning.Text = ""


        Catch exDivbyZero As DivideByZeroException
            MsgBox($"A divide by zero Exception was thrown. This code will only execute for Divide by Zero Exceptions. The Error was:    {exDivbyZero.Message}")
        Catch exInvalidCast As InvalidCastException
            MsgBox($"An invalid Cast Exception was thrown. This code will only execute for invalid Cast Exceptions. The Error was:     {exInvalidCast.Message}")
        Catch ex As Exception

            MsgBox($"{Negative_Error} - {Errormsg}")
        End Try


    End Sub

    Private Sub Btn_Clear_Click(sender As Object, e As EventArgs) Handles Btn_Clear.Click
        lb_InterestRate.Text = ""
        lb_LoanAmount.Text = ""
        lb_LoanPMonths.Text = ""
        lb_result.Text = ""
        Lb_Warning.Text = ""
        Txt_Loan.Text = ""
        Txt_Rate.Text = ""
        Txt_Period.Text = ""

    End Sub

    Private Sub Txt_Loan_TextChanged(sender As Object, e As EventArgs) Handles Txt_Loan.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
