'------------------------------------------------------------------------------------------------
'Program Info.
'------------------------------------------------------------------------------------------------
'Program: Km Tracker
'Date:06/12/2019
'Autor: Felipe SG
'Program Details: Sales reps may claim their travel costs and have their KMs 
'reimbursed by the company. The company pays .53 cents per KM And If the KMs 
'travelled For a particular day Is over 300 KM, a flat daily rate Of $150 Is paid. 
'The sales rep needs a program To keep track Of mileage And calculate the information 
'To be reimbursed And reported To the company.

'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'06/12/2019         Felipe SG           Initial Version

Public Class Form1

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clear_all(Me)
    End Sub

    Public Sub btn_reset_Click(sender As Object, e As EventArgs) Handles btn_reset.Click
        clear_all(Me)
    End Sub
    Public Sub btn_add_Click(sender As Object, e As EventArgs) Handles btn_add.Click

        'validation of the data entry
        Try
            'If empty return an error
            If txb_km.Text = "" Then
                error_msg(4)
                clear_focus()
                'If letter return an error
            ElseIf Not IsNumeric(txb_km.Text) Then
                error_msg(1)
                clear_focus()
                Return
                'If 0 or less return an error
            ElseIf txb_km.Text <= 0 Then
                error_msg(2)
                clear_focus()
                Return
                'If over 1500 return an error
            ElseIf txb_km.Text > 1500 Then
                error_msg(5)
                clear_focus()
                'If has commas return an error
            ElseIf txb_km.Text.Contains(Chr(44)) Then
                error_msg(7)
                clear_focus()
                'If has dots return an error
            ElseIf txb_km.Text.Contains(Chr(46)) Then
                error_msg(8)
                clear_focus()
            Else

                'copy the text box to list box
                listbox_km.Items.Add(txb_km.Text)

                '------------------RegularRate------------------
                'Print the total KM
                lb_km_Travelled.Text = calculate_distance(listbox_km)
                'Print the total value of the reimbursement
                lb_reimbursement.Text = $"{calculate_value(listbox_km).ToString("c2")}"
                'Print the total days of 
                lb_days_travelled.Text = calulate_days(listbox_km)

                '------------------FlatRate------------------
                'Print the total KM
                lb_km_flatrate.Text = calculate_distance(listbox_km, validation:=True)
                'Print the total value of the reimbursement
                lb_reimbursement_flatrate.Text = $"{calculate_value(listbox_km, validation:=True).ToString("C2")}"
                'Print the total days of 
                lb_days_flatrate.Text = calulate_days(listbox_km, validation:=True)

            End If

        Catch ex As Exception
            error_msg(6)
            clear_focus()
        End Try
    End Sub
End Class
