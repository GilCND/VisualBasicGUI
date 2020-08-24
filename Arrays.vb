'Program Name: Arrays
'Objective: Check for errors when user enters the name of the instructor. 
Programmer: Felipe SG


Public Class Form1

    Dim array_classroom(9) As String
    Dim array_instructor(9) As String

'Check for errors
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            Dim index As Integer
            Dim instructor_last As String = Nothing

            If tb_instructor.Text = Nothing Then
                ErrorMsg(1)
                CleanFocus(tb_instructor)
                Return

            ElseIf IsNumeric(tb_instructor.text) Or tb_instructor.text = Nothing Then
                ErrorMsg(2)
                CleanFocus(tb_instructor)
                Return

            ElseIf tb_classroom.text = Nothing Then
                ErrorMsg(5)
                CleanFocus(tb_classroom)
                Return


            Else
            'Check if the name entered was already there
                For index = 0 To array_instructor.Length - 1
                    If array_instructor(index) = tb_instructor.Text Then
                        MsgBox($"Name already present at {index} at classRoom {array_classroom(index)}")
                        tb_instructor.Clear()
                        tb_classroom.Clear()
                        Return

            'Check if for already assigned info
                    ElseIf array_instructor(index) = tb_instructor.text And array_classroom(index) = tb_classroom.Text Then
                        MsgBox($"Name and classroom already assigned")
                        InstructorClearFocus()
                        ClassRoomClearFocus()
                        Return

                    ElseIf array_instructor(index) = Nothing And array_classroom(index) = Nothing Then
                        array_instructor(index) = tb_instructor.Text
                        array_classroom(index) = tb_classroom.Text

                        Exit For

                    End If

                Next


            End If

'Exceptions added

        Catch ex00B As IndexOutOfRangeException
            MsgBox("Error")
        Catch ex As Exception
            MsgBox("Error2")
        End Try
    End Sub
    
    'When the button was clicked.
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
'Check for errors
            If rad_instructor.Checked = True Then
                If tb_search.Text = Nothing Then
                    ErrorMsg(1)
                    CleanFocus(tb_search)
                    Return
                Else
                    For Index = 0 To array_instructor.Length - 1

                        If array_instructor(Index) = tb_search.Text Then

                            MsgBox($"{array_instructor(Index)} is teaching at Classroom {array_classroom(Index)}")

                            Return
                        End If
                    Next
                    ErrorMsg(6)
                    Return
                End If

            ElseIf rad_classroom.Checked = True Then
                If tb_search.Text = Nothing Then
                    ErrorMsg(1)
                    CleanFocus(tb_search)
                Else
                    For Index = 0 To array_classroom.Length - 1

                        If array_classroom(Index) = tb_search.Text Then

                            MsgBox($"Classroom {array_classroom(Index)} is being used by {array_instructor(Index)}")

                            Return
                        End If
                    Next
                    ErrorMsg(6)
                    Return
                End If
            ElseIf rad_Instructor.Checked = False Or rad_classroom.Checked = False Then

                ErrorMsg(7)

                Return

            ElseIf tb_search.Text = Nothing Then

                ErrorMsg(1)

            End If

        Catch IndexOutofRange As IndexOutOfRangeException

            ErrorMsg(8)
            MsgBox(IndexOutofRange.Message)

        Catch ex As Exception
        End Try

    End Sub

    Sub InstructorClearFocus()
        tb_instructor.Clear()
        tb_instructor.Focus()
    End Sub
    Sub ClassRoomClearFocus()
        tb_classroom.Clear()
        tb_classroom.Focus()
    End Sub

End Class
