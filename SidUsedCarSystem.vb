'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Sid Used Car System
'Date:20/11/2019
'Autor: Felipe SG
'Program Details: The program should have:
'Text box to allow a quantity to be entered
'Two List boxes, one to store the details of the Cars ordered, the second to store
'the extended price of the cars being ordered.
'A button to add the quantity of cars selected to the items list box and the extended
'price to the extended price list box
'A button to “Remove” the selected item from both list boxes
'Three Label Controls, one to Display the Subtotal, the second to display Tax and
'the third to display the total order price
'When the “Add to order” button is clicked, the Make, Model and Quantity will be added to the 
'Order details list box, And the extended price will be calculated And added To the prices list box
'The taxes and the total label price labels will be updated each time an item is added to the list box. 
'A loop should be used to calculate the subtotal based on the contents of the listbox And the tax And total 
'labels updated To reflect the New total.
'When the “Remove” button is clicked, if an item is selected in the “Order Details” List box, the selected 
'item will be removed from both list boxes at the index that was selected. When an item Is removed, the tasks 
'from Step 4 should be executed again
'Exception type specific handling should be used to trap any errors that might occur with the quantity text box. 
'A generic Exception handler should also be used To trap For unaccounted problems. Problems should be reported To the user.
'When an item is selected in the “Models” listbox, if a price is not defined for the selected model, 
'Throw And handle an IndexOutOfRangeException. Show the user a user friendly message that the price was Not found. 
'(Add a model To the box without a corresponding price To test this functionality)
'Add Validation to the combobox and listbox to ensure something is selected prior to adding the item to the list box. 
'Working clear button
'working exit button


'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'20/11/2019         Felipe SG            Initial Version
'26/11/2019         Felipe SG            Version 1.1 Functions and sub implemented

Public Class Form1


    Dim Initial_price,
        Counting,
        Removed,
        Final_Price,
        Price_tax,
        sum,
        car_price As Decimal
    Dim i As Integer
    Dim Model_Selected,
        Errormsg,
        Error_field,
        compile As String

    'Constants for prices
    Const tax = 0.15


    'Constants for names
    Const Camry_Model = "Camry",
          RAV4_Model = "RAV4",
          Celica_Model = "Celica",
          Corolla_Model = "Corolla",
          Yaris_Model = "Yaris",
          RAM1500_Model = "RAM 1500",
          Avenger_Model = "Avenger",
          Viper_Model = "Viper",
          Durango_Model = "Durango",
          Shadow_Model = "Shadow",
          Mustang_Model = "Mustang",
          Focus_Model = "Focus",
          Fusion_Model = "Fusion",
          F150_Model = "F 150",
          Escape_Model = "Escape",
          Tempo_Model = "Tempo",
          Model3_Model = "Model 3",
          Models_Model = "Model S",
          Modelx_Model = "Model X",
          Roadster_Model = "Roadster",
           lbt_Model = "Supercar"

    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        Call clearall()
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_Maker.SelectedIndexChanged
        'Here the case statments to fill the combo box
        Call vehiclemakes()

    End Sub
    Private Sub lstbox_Details_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstbox_Details.SelectedIndexChanged
        'Here is the part of the code that when you select something from Details listbox it auto select the same index from Extended Price listbox
        i = lstbox_Details.SelectedIndex
        lstbox_ExtendedPrice.SelectedIndex = i

    End Sub

    Private Sub cbox_Maker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstbox_Model.SelectedIndexChanged
        Model_Selected = ""
        'Case Statment to feed the variables when user seletc a model from list box

        Call populatemodels()

    End Sub


    Private Sub btn_AddQuantity_Click(sender As Object, e As EventArgs) Handles btn_AddQuantity.Click
        'All Exceptions stated here 
        Try
            If cbox_Maker.SelectedIndex < 0 Then
                cbox_Maker.Focus()
                Errormsg = "Please Select a Maker"
                Error_field = "Maker Not selected"
                Throw New Exception
            ElseIf Model_Selected = "" Then
                lstbox_Model.Focus()
                Errormsg = "Please Select some model"
                Error_field = "Model Not selected"
                Throw New Exception
            ElseIf txbox_Quantity.Text = "" Then
                txbox_Quantity.Focus()
                Errormsg = "You forgot To Select the quantity"
                Error_field = "Quantity"
                Throw New Exception
            ElseIf txbox_Quantity.Text = 0 Then
                txbox_Quantity.Focus()
                Errormsg = "0 Is Not allowed"
                Error_field = "Quantity"
                Throw New Exception
            ElseIf txbox_Quantity.Text < 0 Then
                txbox_Quantity.Focus()
                Errormsg = "Negative number Not allowed"
                Error_field = "Quantity wrong"
                Throw New Exception
            ElseIf lb_Subtotal.Text = 0 Then
                Throw New IndexOutOfRangeException

            End If
            sum = 0
            Price_tax = 0
            Final_Price = 0
            lb_Total.Text = 0
            lb_Tax.Text = 0

            If txbox_Quantity.Text > 0 Then
                Initial_price = car_price * txbox_Quantity.Text
                compile = $"{txbox_Quantity.Text}{vbTab}{Model_Selected}{vbTab}${Initial_price}"
                If Counting < 1 Then
                    lstbox_Details.Items.Add($"Quantity{vbTab}Model{vbTab}Price")
                    lstbox_ExtendedPrice.Items.Add("Price after Taxes")
                    Counting = 1
                End If
                lstbox_Details.Items.Add(compile)
                lstbox_ExtendedPrice.Items.Add(Initial_price)

            End If

            ' Here is the loop responsible by sum all values from listbox Extended price
            Call calcLoop()
            Price_tax = sum * tax
            Final_Price = sum + Price_tax
            lb_Total.Text = $"$:{Final_Price.ToString}"
            lb_Tax.Text = $"$:{Price_tax}"
            lb_subtotal_final.Text = $"$:{sum}"

        Catch outofrange As IndexOutOfRangeException
            MsgBox("I'm sorry, the selected model has no price, please select another one")
        Catch exDivbyZero As DivideByZeroException
            MsgBox($"A divide by zero Exception was thrown. This code will only execute for Divide by Zero Exceptions. The Error was:    {exDivbyZero.Message}")
        Catch exInvalidCast As InvalidCastException
            MsgBox($"An invalid Cast Exception was thrown. This code will only execute for invalid Cast Exceptions. The Error was:     {exInvalidCast.Message}")
        Catch ex As Exception
            MsgBox($"Error {Error_field} {Errormsg}")



        End Try
    End Sub
    Private Sub btn_Remove_Click(sender As Object, e As EventArgs) Handles btn_Remove.Click
        Try
            If lstbox_ExtendedPrice.SelectedIndex < 0 Then
                lstbox_ExtendedPrice.Focus()
                Errormsg = "Please select what you want to eliminate from Details box"
                Error_field = "To remove something"
                Throw New Exception


            End If

            'when user click on remove button it will calculate the sum, using the same

            Removed = lstbox_ExtendedPrice.Items.Item(lstbox_ExtendedPrice.SelectedIndex)
            lstbox_ExtendedPrice.Items.Remove(lstbox_ExtendedPrice.Items.Item(lstbox_ExtendedPrice.SelectedIndex))
            lstbox_Details.Items.Remove(lstbox_Details.Items.Item(lstbox_Details.SelectedIndex))
            sum = 0
            Price_tax = 0
            Final_Price = 0
            lb_Total.Text = 0
            lb_Tax.Text = 0
            Call calcLoop()
            Price_tax = sum * tax
            Final_Price = sum + Price_tax
            lb_Total.Text = Final_Price.ToString
            lb_subtotal_final.Text = sum
            lb_Tax.Text = Price_tax

        Catch exDivbyZero As DivideByZeroException
            MsgBox($"A divide by zero Exception was thrown. This code will only execute for Divide by Zero Exceptions. The Error was:    {exDivbyZero.Message}")
        Catch exInvalidCast As InvalidCastException
            MsgBox($"An invalid Cast Exception was thrown. This code will only execute for invalid Cast Exceptions. The Error was:     {exInvalidCast.Message}")
        Catch ex As Exception
            MsgBox($"Error {Error_field} {Errormsg}")
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call populatemakes()
        Call clearall()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Application.Exit()
    End Sub
    Sub calcLoop()


        For i = 0 To lstbox_ExtendedPrice.Items.Count - 1
            sum += Val(lstbox_ExtendedPrice.Items(i))

        Next
    End Sub
    Sub vehiclemakes()
        Select Case cbox_Maker.SelectedIndex
            Case 0
                lstbox_Model.Items.Clear()
                lstbox_Model.Items.Add(RAM1500_Model)
                lstbox_Model.Items.Add(Avenger_Model)
                lstbox_Model.Items.Add(Viper_Model)
                lstbox_Model.Items.Add(Durango_Model)
                lstbox_Model.Items.Add(Shadow_Model)
                lstbox_Model.Items.Add(lbt_Model)
            Case 1
                lstbox_Model.Items.Clear()
                lstbox_Model.Items.Add(Mustang_Model)
                lstbox_Model.Items.Add(Focus_Model)
                lstbox_Model.Items.Add(F150_Model)
                lstbox_Model.Items.Add(Fusion_Model)
                lstbox_Model.Items.Add(Escape_Model)
                lstbox_Model.Items.Add(Tempo_Model)
            Case 2
                lstbox_Model.Items.Clear()
                lstbox_Model.Items.Add(Model3_Model)
                lstbox_Model.Items.Add(Models_Model)
                lstbox_Model.Items.Add(Modelx_Model)
                lstbox_Model.Items.Add(Roadster_Model)
            Case 3
                lstbox_Model.Items.Clear()
                lstbox_Model.Items.Add(Camry_Model)
                lstbox_Model.Items.Add(Celica_Model)
                lstbox_Model.Items.Add(RAV4_Model)
                lstbox_Model.Items.Add(Corolla_Model)
                lstbox_Model.Items.Add(Yaris_Model)
        End Select
    End Sub
    Sub clearall()

        lstbox_Model.Items.Clear()
        lstbox_ExtendedPrice.Items.Clear()
        lb_Subtotal.Text = ""
        lb_subtotal_final.Text = ""
        lb_Tax.Text = ""
        lb_Total.Text = ""
        cbox_Maker.SelectedIndex = -1
        txbox_Quantity.Text = ""
        lstbox_Details.Items.Clear()


    End Sub
    Sub populatemodels()
        car_price = modelprice()

        Select Case lstbox_Model.Text
            Case "RAV4"
                Call bringprice()
                Model_Selected = RAV4_Model
            Case "Camry"
                Call bringprice()
                Model_Selected = Camry_Model
            Case "Celica"
                Call bringprice()
                Model_Selected = Celica_Model
            Case "Corolla"
                Call bringprice()
                Model_Selected = Corolla_Model
            Case "Yaris"
                Call bringprice()
                Model_Selected = Yaris_Model
            Case "RAM 1500"
                Call bringprice()
                Model_Selected = RAM1500_Model
            Case "Avenger"
                Call bringprice()
                Model_Selected = Avenger_Model
            Case "Viper"
                Call bringprice()
                Model_Selected = Viper_Model
            Case "Durango"
                Call bringprice()
                Model_Selected = Durango_Model
            Case "Shadow"
                Call bringprice()
                Model_Selected = Shadow_Model
            Case "Mustang"
                Call bringprice()
                Model_Selected = Mustang_Model
            Case "Focus"
                Call bringprice()
                Model_Selected = Focus_Model
            Case "Fusion"
                Call bringprice()
                Model_Selected = Fusion_Model
            Case "F 150"
                Call bringprice()
                Model_Selected = F150_Model
            Case "Escape"
                Call bringprice()
                Model_Selected = Escape_Model
            Case "Tempo"
                Call bringprice()
                Model_Selected = Tempo_Model
            Case "Model 3"
                Call bringprice()
                Model_Selected = Model3_Model
            Case "Model S"
                Call bringprice()
                Model_Selected = Models_Model
            Case "Model X"
                Call bringprice()
                Model_Selected = Modelx_Model
            Case "Roadster"
                Call bringprice()
                Model_Selected = Roadster_Model
            Case "Supercar"
                lb_Subtotal.Text = 0
                Model_Selected = lbt_Model


        End Select
    End Sub
    Sub populatemakes()
        cbox_Maker.Items.Add("Dodge")
        cbox_Maker.Items.Add("Ford")
        cbox_Maker.Items.Add("Tesla")
        cbox_Maker.Items.Add("Toyota")

    End Sub
    Function modelprice()

        Select Case lstbox_Model.Text
            Case "RAV4"
                car_price = 3200
            Case "Celica"
                car_price = 3000
            Case "Corolla"
                car_price = 20000
            Case "Yaris"
                car_price = 15000
            Case "Corolla"
                car_price = 20000
            Case "Corolla"
                car_price = 20000
            Case "RAM 1500"
                car_price = 40000
            Case "Avenger"
                car_price = 30000
            Case "Viper"
                car_price = 22000
            Case "Durango"
                car_price = 12000
            Case "Shadow"
                car_price = 112000
            Case "Mustang"
                car_price = 332000
            Case "Focus"
                car_price = 2000
            Case "Fusion"
                car_price = 5700
            Case "F 150"
                car_price = 9999
            Case "Escape"
                car_price = 7000
            Case "Tempo"
                car_price = 2500
            Case "Model 3"
                car_price = 7000
            Case "Model S"
                car_price = 6000
            Case "Model X"
                car_price = 500
            Case "Roadster"
                car_price = 3000
            Case "Supercar"
                car_price = 0
        End Select
        Return car_price
    End Function
    Sub bringprice()
        lb_Subtotal.Text = ""
        lb_Subtotal.Text = $"${car_price}"
    End Sub

End Class




