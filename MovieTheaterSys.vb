'------------------------------------------------------------------------------------------------
'Program Info
'------------------------------------------------------------------------------------------------
'Program: Movie Theater Sys
'Date:26/11/2019
'Autor: Felipe SG
'Program Details: the program should do
'Allow a Product and a Quantity to be selected
'Allow a Size to be selected for certain products
'Allow Date to Be selected
'Add order details/prices to the corresponding list boxes.
'Sum the Subtotal by using a loop to sum the listbox prices
'Calculate the Discount
'Calculate the total Tax
'Calculate the total Order price
'Calculate the change based on amount paid
'Handle any Exceptions that may occur using exception handling code on input fields.
'Validate input where exceptions are not practical.
'Should have discount day on Tuesdays of 30% on the tickets


'------------------------------------------------------------------------------------------------
'Change Log
'------------------------------------------------------------------------------------------------
' Date              Programmer              Change
'------------------------------------------------------------------------------------------------
'26/11/2019         Felipe SG            Initial Version



Public Class Form1


    ' Constant Numbers
    Const c_tax = 0.15,
          c_movie_ticket_price = 14.99,
          c_small_popcorn_price = 3.99,
          c_medium_popcorn_price = 4.99,
          c_large_popcorn_price = 5.99,
          c_small_pop_price = 4.0,
          c_medium_pop_price = 5.0,
          c_large_pop_price = 6.5,
          c_bars_price = 2.99,
          c_ticket_discount = 0.3


    ' Constant Names
    Const c_movie_ticket_name = "Movie Ticket",
          c_small_popcorn_name = "Small PopCorn",
          c_medium_popcorn_name = "Medium PopCorn",
          c_large_popcorn_name = "Large PopCorn",
          c_small_pop_name = "Small Pop",
          c_medium_pop_name = "Medium Pop",
          c_large_pop_name = "Large Pop",
          c_bars_name = "Bars"


    'Variables
    Dim product_price,
        ticket_price,
        ticket_discount_total,
        ticket_tax,
        ticket_total,
        product_tax,
        product_total,
        product_final_price,
        total_discount,
        print_order_total As Decimal
    Dim product_name As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Populate the combo box with the products and disable the groupbox
        cbox_product.Items.Add("Movie Ticket")
        cbox_product.Items.Add("Popcorn")
        cbox_product.Items.Add("Pop")
        cbox_product.Items.Add("Bars")
        gbox_size.Enabled = False
        product_name = ""
        lb_ordertotal.Text = ""
        lb_change.Text = ""
        lb_totaldiscount.Text = ""
        lb_totaltax.Text = ""

    End Sub

    Private Sub btn_add_to_order_Click(sender As Object, e As EventArgs) Handles btn_add_to_order.Click
        Try
            'All the basic errors covered by this if statement 
            If cbox_product.SelectedIndex = -1 Then
                MsgBox("Please Select something from the products list")
                cbox_product.Focus()
            ElseIf txbox_quantity.Text = "" Then
                MsgBox("Please insert a number into the quantity field.")
                txbox_quantity.Focus()
            ElseIf Not IsNumeric(txbox_quantity.Text) Then
                MsgBox("Please select a valid number at the quantity field, numbers are not allowed.")
                txbox_quantity.Focus()
            ElseIf txbox_quantity.Text <= 0 Then
                MsgBox("Please insert a valid number on the Quantity field, greater than 0")
                txbox_quantity.Focus()
            ElseIf txbox_quantity.text > 10000 Then
                MsgBox("The maximun of itens per line is 10000, if you need more, please add more lines")
                txbox_quantity.Focus()
            ElseIf print_order_total > 1000000 Then
                MsgBox("I'm sorry this cachier only knows how to count to a million. If you need all that or more please call the C.E.O.")
            Else
                Dim compile As String
                Dim the_date As Date
                Dim sum,
            tax_total,
            print_ticket_discount As Decimal
                the_date = date_picker.Value
                lb_change.Text = ""
                Select Case cbox_product.SelectedIndex
                    Case 0
                        'Movie Ticket
                        'if condition to check for days of the week, if is tuesday, apply the discount

                        If the_date.Date < Today.Date Then
                            MsgBox("Please select a valid date, it's not possible to buy something in the past")
                        ElseIf the_date.DayOfWeek = DayOfWeek.Tuesday Then
                            ticket_discount_total = (c_movie_ticket_price * txbox_quantity.Text) * c_ticket_discount
                            print_ticket_discount = c_movie_ticket_price - (c_movie_ticket_price * c_ticket_discount)
                            total_discount += ticket_discount_total
                            ticket_price = (c_movie_ticket_price * txbox_quantity.Text) - ticket_discount_total
                            ticket_tax = ticket_price * c_tax
                            ticket_total = ticket_price + ticket_tax
                            compile = $"{c_movie_ticket_name}{vbTab}Qnt:{txbox_quantity.Text}{vbTab}Price:{print_ticket_discount.ToString("C2")}"
                            lbox_items.Items.Add(compile)
                            lbox_linetotal.Items.Add(ticket_price.ToString("C2"))


                        Else
                            ticket_discount_total = 0
                            ticket_price = (c_movie_ticket_price * txbox_quantity.Text) - ticket_discount_total
                            ticket_tax = ticket_price * c_tax
                            ticket_total = ticket_price + ticket_tax
                            compile = $"{c_movie_ticket_name}{vbTab}Qnt:{txbox_quantity.Text}{vbTab}Price:{c_movie_ticket_price.ToString("C2")}"
                            lbox_items.Items.Add(compile)
                            lbox_linetotal.Items.Add(ticket_price.ToString("C2"))

                        End If

                    Case 1
                        'Popcorn
                        'if condition to check which one of the radio buttons are checked and perform 
                        'specific actions
                        If rbtn_small.Checked = True Then
                            product_price = c_small_popcorn_price
                            product_name = c_small_popcorn_name
                        ElseIf rbtn_medium.Checked = True Then
                            product_price = c_medium_popcorn_price
                            product_name = c_medium_popcorn_name
                        Else
                            product_price = c_large_popcorn_price
                            product_name = c_large_popcorn_name
                        End If
                        product_total = (product_price * txbox_quantity.Text)
                        product_tax = product_total * c_tax
                        product_final_price = product_total + product_tax
                        compile = $"{product_name}{vbTab}Qnt:{txbox_quantity.Text}{vbTab}Price:{product_price.ToString("C2")}"
                        lbox_items.Items.Add(compile)
                        lbox_linetotal.Items.Add(product_total.ToString("C2"))
                    Case 2
                        'Pop
                        'if condition to check which one of the radio buttons are checked and perform 
                        'specific actions
                        If rbtn_small.Checked = True Then
                            product_price = c_small_pop_price
                            product_name = c_small_pop_name
                        ElseIf rbtn_medium.Checked = True Then
                            product_price = c_medium_pop_price
                            product_name = c_medium_pop_name
                        Else
                            product_price = c_large_pop_price
                            product_name = c_large_pop_name
                        End If
                        product_total = (product_price * txbox_quantity.Text)
                        product_tax = product_total * c_tax
                        product_final_price = product_total + product_tax
                        compile = $"{product_name}{vbTab}Qnt:{txbox_quantity.Text}{vbTab}Price:{product_price.ToString("C2")}"
                        lbox_items.Items.Add(compile)
                        lbox_linetotal.Items.Add(product_total.ToString("C2"))
                    Case 3
                        'Bars  
                        product_total = (product_price * txbox_quantity.Text)
                        product_tax = product_total * c_tax
                        product_final_price = product_total + product_tax
                        compile = $"{product_name}{vbTab}{vbTab}Qnt:{txbox_quantity.Text}{vbTab}Price:{product_price.ToString("C2")}"
                        lbox_items.Items.Add(compile)
                        lbox_linetotal.Items.Add(product_total.ToString("C2"))
                End Select
                'Counting loop on the lbox_Linetotal
                For i = 0 To lbox_linetotal.Items.Count - 1
                    sum += lbox_linetotal.Items(i)
                    tax_total = sum * c_tax

                    'Print the totals to the labels
                    lb_totaltax.Text = tax_total.ToString("C2")
                    print_order_total = sum + tax_total
                    lb_ordertotal.Text = print_order_total.ToString("C2")
                    lb_totaldiscount.Text = total_discount.ToString("C2")
                Next
            End If

        Catch ex As Exception
            MsgBox("ERROR. Please check with the IT responsible")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_product.SelectedIndexChanged
        Select Case cbox_product.Text
            Case "Movie Ticket"

                'Disable the gbox and deselect all radio buttons
                date_picker.Enabled = True
                rbtn_small.Checked = False
                rbtn_medium.Checked = False
                rbtn_large.Checked = False
                gbox_size.Enabled = False
                txbox_quantity.Focus()
                txbox_quantity.Clear()
                product_name = c_movie_ticket_name

            Case "Popcorn"
                'Enable radio box, group box and, disable the calendar
                rbtn_large.Checked = True
                gbox_size.Enabled = True
                date_picker.Enabled = False
                txbox_quantity.Focus()
                txbox_quantity.Clear()


            Case "Pop"
                'Enable radio box, group box and, disable the calendar
                gbox_size.Enabled = True
                date_picker.Enabled = False
                rbtn_large.Checked = True
                txbox_quantity.Focus()
                txbox_quantity.Clear()

            Case "Bars"
                'Disable the groupbox and deselect all radio buttons and, disable the calendar
                rbtn_small.Checked = False
                rbtn_medium.Checked = False
                rbtn_large.Checked = False
                gbox_size.Enabled = False
                date_picker.Enabled = False
                product_price = c_bars_price
                product_name = c_bars_name
                txbox_quantity.Focus()
                txbox_quantity.Clear()
        End Select
    End Sub
    Private Sub btn_calculate_change_Click(sender As Object, e As EventArgs) Handles btn_calculate_change.Click

        Try
            Dim due As Decimal
            'If, to check the value of the money paid performing certain conditions in each scenario
            If tbox_money_paid.Text = "" Then
                MsgBox("Please insert a number on the Value Paid by the customer")
                tbox_money_paid.Clear()
                lb_change.Text = ""
                tbox_money_paid.Focus()
            ElseIf Not IsNumeric(tbox_money_paid.Text) Then
                MsgBox("Only numbers can be accepted")
                tbox_money_paid.Clear()
                lb_change.Text = ""
                tbox_money_paid.Focus()
            ElseIf tbox_money_paid.Text <= 0 Then
                MsgBox("Please Insert a valid number, greater than 0")
                tbox_money_paid.Clear()
                lb_change.Text = ""
                tbox_money_paid.Focus()
            ElseIf tbox_money_paid.Text < print_order_total Then
                due = print_order_total - tbox_money_paid.Text
                MsgBox($"The ammount of money paid it's not enough, wee need more {due.ToString("c2")}")
                tbox_money_paid.Clear()
                lb_change.Text = ""
                tbox_money_paid.Focus()

            Else
                If (tbox_money_paid.Text - print_order_total) > 500 Then
                    MsgBox("We only allowed 500 dolars in cash back, sorry!")
                Else
                    lb_change.Text = (tbox_money_paid.Text - print_order_total).ToString("c2")
                End If
            End If
        Catch ex As Exception
            MsgBox("ERROR. Please check with the IT responsible")
        End Try

    End Sub

End Class
