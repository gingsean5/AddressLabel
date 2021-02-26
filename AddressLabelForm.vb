Option Strict On
Option Explicit On

'Sean Gingerich
'RCET0265
'Spring 2021
'Address Label
'https://github.com/gingsean5/AddressLabel

Public Class AddressLabelForm
    Private Sub DisplayButton_Click(sender As Object, e As EventArgs) Handles DisplayButton.Click

        ValidateFields()

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        FirstNameTextBox.Text = ""
        LastNameTextBox.Text = ""
        StreetAddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipTextBox.Text = ""
    End Sub

    Function ValidateFields() As Boolean
        Dim problem As String
        If FirstNameTextBox.Text = "" Then
            problem &= "First name is required" & vbNewLine
        End If
        If LastNameTextBox.Text = "" Then
            problem &= "Last name is required" & vbNewLine
        End If
        If StreetAddressTextBox.Text = "" Then
            problem &= "Street Address is required" & vbNewLine
        End If
        If CityTextBox.Text = "" Then
            problem &= "City is required" & vbNewLine
        End If
        If StateTextBox.Text = "" Then
            problem &= "State is required" & vbNewLine
        End If
        If ZipTextBox.Text = "" Then
            problem &= "Zipcode is required" & vbNewLine
        End If
        Dim zipInt As Integer
        Dim zipStr As String
        Try
            zipStr = ZipTextBox.Text
            zipInt = CInt(zipStr)

        Catch ex As Exception
            problem &= "Zipcode must be a number"
        End Try

        If problem <> "" Then
            MsgBox(problem)
        End If

        If problem = "" Then
            DisplayLabel.Text = Trim(FirstNameTextBox.Text) & " " & Trim(LastNameTextBox.Text) & vbNewLine _
            & Trim(StreetAddressTextBox.Text) & vbNewLine _
            & Trim(CityTextBox.Text) & ", " & Trim(StateTextBox.Text) & " " & Trim(ZipTextBox.Text)
        End If
    End Function

End Class
