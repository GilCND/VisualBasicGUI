'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Timber Toms Sys
'Date: 06/11/2019
'Autor: Felipe SG
'Add an order an item button and text boxes which will 
'Allow the user to enter an item number
'Allow the user to enter a quantity of items for the item being ordered 
'Add a group box with radio buttons which will allow you to apply discounts to the
'order. They should include:  No Discount, 10% Discount, and 15% Discount 
'Use if or case statements to get the item description and price from a table
'Calculate the extended price (this is the price of the item multiplied by the quantity)
'Add a listbox to track the items that have been ordered. Based on user actions using the 
'previously defined controls, display in the list box (one line per item):
'item number / description / quantity / price / extended price
'Add a second button to complete the order as well as list box control to display invoices.
'When the order completed button is pressed populate the invoice list box in the following manner: 
'The company name, date and time in the invoice list box
'Using a Loop, copy the items from the ordered items list box to the invoice list box
'Calculate and display the discount (based on the radio buttons) 4. Calculate and display the tax. 
'Display the total cost. / Add a working clear button / Add a working exit button 
'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'06/11/2019         Felipe SG        Initial Version


Public Class Form1

    Dim itemid As Integer
    Dim item_description, compile As String
    Dim price, price_sum, extended_price, discount, total_tax As Decimal
    Dim total_price_item, total_price_sum, final_price, total_discount As Decimal
    Dim currenttime As Date
    Const tax As Decimal = 0.15


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btn_completeorder.Click

        currenttime = Now
        lstbox_Invoice.Items.Add($"Timber Tom's Hardware {currenttime}")

        'Loop used to populate the listbox Invoice
        For counter = 0 To (Lstbox_ordered.Items.Count - 1)
            lstbox_Invoice.Items.Add(Lstbox_ordered.Items.Item(counter))

        Next

        Select Case True
            Case radio_10discount.Checked
                discount = 0.1
            Case radio_15discount.Checked
                discount = 0.15
            Case Else
                discount = 0
        End Select
        'calculations
        total_discount = discount * total_price_sum
        total_tax = (total_price_sum - total_discount) * tax
        final_price = (total_price_sum - total_discount) + total_tax


        'Print the reciept 
        lstbox_Invoice.Items.Add($"Discount  ${ total_discount.ToString("n2")}")
        lstbox_Invoice.Items.Add($"TAX  ${ total_tax.ToString("n2")}")
        lstbox_Invoice.Items.Add($"Total  ${ final_price.ToString("n2")}")



    End Sub

    Private Sub btn_addtoorder_Click(sender As Object, e As EventArgs) Handles btn_addtoorder.Click
        lbl_error.Text = ""
        itemid = 0
        If IsNumeric(txt_itemid.Text) Then
            If IsNumeric(txt_quantity.Text) Then

                Select Case txt_itemid.Text
                    Case 100
                        itemid = 100
                        item_description = "Wrench                    "
                        price = 3.5
                    Case 101
                        itemid = 101
                        item_description = "Pipe Wrench               "
                        price = 5.75
                    Case 200
                        itemid = 200
                        item_description = "Rip Saw                   "
                        price = 16.23
                    Case 201
                        itemid = 201
                        item_description = "Framing Hammer            "
                        price = 32.5
                    Case 203
                        itemid = 203
                        item_description = "Framing Square            "
                        price = 27.5
                    Case 305
                        itemid = 305
                        item_description = "Solder (roll)             "
                        price = 6.34
                    Case 306
                        itemid = 306
                        item_description = "Paste                     "
                        price = 4.26
                    Case Else
                        lbl_error.Text = "Please Inform a valid item ID"

                End Select
                If itemid = 0 Then

                Else
                    total_price_item = price * txt_quantity.Text
                    total_price_sum = total_price_item + total_price_sum

                    'Compile variable used to compuile other variables in only one line
                    compile = $"{itemid}{vbTab}{item_description}{vbTab}{txt_quantity.Text}{vbTab}${price:#,##0.00}{vbTab}${total_price_item:#,##0.00}"
                    Lstbox_ordered.Items.Add(compile)
                    item_description = ""

                End If

            Else
                lbl_error.Text = "Error Only numbers allowed."
            End If
        Else
            lbl_error.Text = "Error Only numbers allowed."
        End If




    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        'clear all list box 
        'Clear Invoice listbox

        While lstbox_Invoice.Items.Count > 0
            lstbox_Invoice.Items.RemoveAt(lstbox_Invoice.Items.Count - 1)
        End While
        'Clear Ordered Listbox

        While Lstbox_ordered.Items.Count > 0
            Lstbox_ordered.Items.RemoveAt(Lstbox_ordered.Items.Count - 1)
        End While

        'Clear all data from variables
        itemid = 0
        total_discount = 0
        total_price_item = 0
        total_price_sum = 0
        total_tax = 0


    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Application.Exit()


    End Sub
End Class
