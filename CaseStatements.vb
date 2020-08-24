'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Case Statments
'Date: 21/10/2019
'Autor: Felipe SG
'The application will use case statements to populate a Models Listbox control based On 
'what was selected In the Make Combobox control. 
'When a New Make Is selected using the combobox, the Models Listbox control should be populated with 
'the relevant models For the selected make. When one Of the Models Is selected In the listbox, a label 
'control should be changed To report the price.

'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'21/10/2019         Felipe SG             Initial Version
'23/10/2019         Felipe SG           Code commented and incremented


Public Class Form1


    'As soon as something is selected on the list, the label should erase previous information and
    'display the price. Here I used the event selected index.
    Private Sub lb_models_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_models.SelectedIndexChanged
        Select Case lb_models.Text
            Case "RAV4"
                lb_price.Text = ""
                lb_price.Text = "$:32.000"
            Case "Camry"
                lb_price.Text = ""
                lb_price.Text = "$:1.000"
            Case "Celica"
                lb_price.Text = ""
                lb_price.Text = "$:3.000"
            Case "Corolla"
                lb_price.Text = ""
                lb_price.Text = "$:20.000"
            Case "Yaris"
                lb_price.Text = ""
                lb_price.Text = "$:15.000"
            Case "Ram_1500"
                lb_price.Text = ""
                lb_price.Text = "$:40.000"
            Case "Avenger"
                lb_price.Text = ""
                lb_price.Text = "$:38.000"
            Case "Viper"
                lb_price.Text = ""
                lb_price.Text = "$:22.000"
            Case "Durango"
                lb_price.Text = ""
                lb_price.Text = "$:12.000"
            Case "Shadow"
                lb_price.Text = ""
                lb_price.Text = "$:112.000"
            Case "Mustang"
                lb_price.Text = ""
                lb_price.Text = "$:332.000"
            Case "Focus"
                lb_price.Text = ""
                lb_price.Text = "$:2.000"
            Case "Fusion"
                lb_price.Text = ""
                lb_price.Text = "$:5.700"
            Case "F150"
                lb_price.Text = ""
                lb_price.Text = "$:9.999"
            Case "Escape"
                lb_price.Text = ""
                lb_price.Text = "$:7.000"
            Case "Tempo"
                lb_price.Text = ""
                lb_price.Text = "$:2.500"
            Case "Model 3"
                lb_price.Text = ""
                lb_price.Text = "$:7.800"
            Case "Model S"
                lb_price.Text = ""
                lb_price.Text = "$:6.000"
            Case "Model X"
                lb_price.Text = ""
                lb_price.Text = "$:500"
            Case "Roadster"
                lb_price.Text = ""
                lb_price.Text = "$:3.000"
        End Select
    End Sub
    'Here is where the models list are generated.
    Private Sub cbx_maker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_maker.SelectedIndexChanged
        Select Case cbx_maker.SelectedIndex
            Case 0
                lb_models.Items.Clear()
                lb_models.Items.Add("Ram_1500")
                lb_models.Items.Add("Avenger")
                lb_models.Items.Add("Viper")
                lb_models.Items.Add("Durango")
                lb_models.Items.Add("Shadow")
            Case 1
                lb_models.Items.Clear()
                lb_models.Items.Add("Mustang")
                lb_models.Items.Add("Focus")
                lb_models.Items.Add("F150")
                lb_models.Items.Add("Fusion")
                lb_models.Items.Add("Escape")
                lb_models.Items.Add("Tempo")
            Case 2
                lb_models.Items.Clear()
                lb_models.Items.Add("Model 3")
                lb_models.Items.Add("Model S")
                lb_models.Items.Add("Model X")
                lb_models.Items.Add("Roadster")
            Case 3
                lb_models.Items.Clear()
                lb_models.Items.Add("Camry")
                lb_models.Items.Add("Celica")
                lb_models.Items.Add("RAV4")
                lb_models.Items.Add("Corolla")
                lb_models.Items.Add("Yaris")

        End Select


    End Sub
End Class
