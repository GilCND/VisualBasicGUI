'Program Name: Loan Calculator
'Objective: Create a GUI and code a Loan calculator. 
'Programmer: Felipe SG


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
    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub btn_calculate_Click(sender As Object, e As EventArgs) Handles btn_calculate.Click
        If (Txt_Loan.Text = "" Or Txt_Period.Text = "" Or Txt_Rate.Text = "") Then
            Lb_Warning.Text = "ERROR, Empty values not acceptable"
        Else
            P = Txt_Loan.Text
            R1 = Txt_Rate.Text
            N = Txt_Period.Text
            R = ((R1 / 100) / 12)
            Result1 = (R * P)
            Result2 = 1 - (1 + R) ^ (-N)
            Result3 = Result1 / Result2
            Final = Result3 * N
            Dif = (Final - P)

            '________________________________________________________________________

            lb_InterestRate.Text = ($"{R1.ToString("n2")} %")
            lb_LoanAmount.Text = ($"${P}")
            lb_LoanPMonths.Text = ($"{N} Months")
            lb_result.Text = ($"Total cost Of the Loan  {Final.ToString("n2")}")
            Lb_Warning.Text = ""
        End If
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
End Class
