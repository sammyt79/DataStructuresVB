' Data Structures in a Visual Basic ® Program
' Samuel Tollefson
' POS/408
' March 23, 2015
' Tim Hagan

Imports System.IO

Public Class MainForm

    Dim dblCost As Double ' Variable to hold txtCost input.
    Dim dblPower As Double ' Variable to hold txtPower input.
    Dim dblHours As Double ' Variable to hold txtHours input.
    Dim dblGallons As Double ' Variable to hold txtGallons input.
    Dim dblCostGal As Double ' Variable to hold txtCostGal input.
    Dim dblTotalHr As Double ' Variable to hold txtTotalHr output.
    Dim dblTotalYr As Double ' Variable to hold txtTotalYr output.
    Dim dblTotal As Double ' Variable to hold the running total.
    Dim water As Boolean = False ' Used to determine if the appliance uses water
    Dim allGood As Boolean = True ' Used to test if the program is ready for calculations.
    Dim utilityFile As StreamWriter ' Create StreamWriter object.

    ' Populate the combobox from a text file.
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Define the file path.
        Dim filePath As String = Application.StartupPath + "\Appliances.txt"

        ' Attempt to connect and read from the file.
        Try
            cboAppliance.Items.AddRange(System.IO.File.ReadAllLines(filePath))

            ' If file is not found, warn the user and close the program.
        Catch ex As Exception
            Dim result As Integer = MessageBox.Show("Sorry, this application is missing a vital file. Would you like to search for the file named ""Appliances""?", "input error", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                MessageBox.Show("Good bye")
                Me.Close() ' Close the program.
                ElseIf result = DialogResult.Yes Then
                Dim myStream As Stream = Nothing
                Dim openFileDialog1 As New OpenFileDialog()

                openFileDialog1.InitialDirectory = "c:\"
                openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                openFileDialog1.FilterIndex = 2
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Try
                        myStream = openFileDialog1.OpenFile()
                        If (myStream IsNot Nothing) Then

                            ' Set the new file path.
                            filePath = openFileDialog1.FileName

                            ' Connect and read from the file.
                            cboAppliance.Items.AddRange(System.IO.File.ReadAllLines(filePath))
                        End If
                    Catch readEx As Exception
                        MessageBox.Show("Cannot read file from disk. Original error: " & ex.Message)
                    Finally
                        ' Check this again, since we need to make sure we didn't throw an exception on open. 
                        If (myStream IsNot Nothing) Then
                            myStream.Close()
                        End If
                    End Try
                End If
                End If
        End Try
    End Sub

    ' Determine what happens when the user presses the "Clear" button (btnClear).
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        ' Reset combo box.
        cboAppliance.SelectedIndex = -1

        ' Clear input text boxes.
        txtCost.Clear()
        txtPower.Clear()
        txtHours.Clear()

    End Sub

    ' Determine what happens when the user presses the "Calculate" button (btnCalculate).
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        allGood = True ' Ensure variable is set to true before inputs are read.

        ' Test combobox for valid user selection.
        If cboAppliance.SelectedIndex = -1 Then

            ' Display error message.
            MsgBox("Please select an appliance", , " input error")
            allGood = False ' Not ready for calculations.
        End If

        ' Validate that the cost entered has the correct format and is within a reasonable range.
        Try
            dblCost = CDbl(txtCost.Text)
            If dblCost < 0.05 Or dblCost > 0.5 Then

                ' Display error message.
                MsgBox("Cost / kW-hour must be between " + FormatCurrency("0.05") + "and " + FormatCurrency("0.5"), , " input error")
                txtCost.Clear() ' Clear cost input textbox.
                allGood = False ' Not ready for calculations.
            End If
        Catch ex As Exception

            ' Display error message.
            MsgBox("Please enter a valid amount for amount for ""Cost""", , " input error")
            txtCost.Clear() ' Clear cost input textbox.
            allGood = False ' Not ready for calculations.
        End Try

        ' Validate that the power entered has the correct format and is within a reasonable range.
        Try
            dblPower = CDbl(txtPower.Text)
            If dblPower < 1 Or dblPower > 100 Then

                ' Display error message.
                MsgBox("The value for Power Needed must be between 1.0 and 100", , " input error")
                txtPower.Clear() ' Clear power input textbox.
                allGood = False ' Not ready for calculations.
            End If
        Catch ex As Exception

            ' Display error message.
            MsgBox("Please enter a valid amount for ""Power Needed""", , " input error")
            txtPower.Clear() ' Clear power input textbox.
            allGood = False ' Not ready for calculations.
        End Try

        ' Validate that the hours entered has the correct format and is within a reasonable range.
        Try
            dblHours = CDbl(txtHours.Text)
            If dblHours < 0 Or dblHours > 24 Then

                ' Display error message.
                MsgBox("Hours Used must be within a 24 hour period", , " input error")
                txtHours.Clear() ' Clear hours input Textbox.
                allGood = False ' Not ready for calculations.
            End If
        Catch ex As Exception

            ' Display error message.
            MsgBox("Please enter a valid amount for ""Hours Used""", , " input error")
            txtHours.Clear() ' Clear hours input Textbox.
            allGood = False ' Not ready for calculations.
        End Try

        ' Validate that the number of gallons entered has the correct format and is within a reasonable range.
        If water = True Then
            Try
                dblGallons = CDbl(txtGallons.Text)
                If dblGallons < 1 Or dblGallons > 20 Then

                    ' Display error message.
                    MsgBox("Number of Gallons must be between 1 and 20", , " input error")
                    txtGallons.Clear() ' Clear gallons input Textbox.
                    allGood = False ' Not ready for calculations.
                End If
            Catch ex As Exception

                ' Display error message.
                MsgBox("Please enter a valid amount for ""Number of Gallons""", , " input error")
                txtGallons.Clear() ' Clear gallons input Textbox.
                allGood = False ' Not ready for calculations.
            End Try

            ' Validate that the cost per gallon entered has the correct format and is within a reasonable range.
            Try
                dblCostGal = CDbl(txtCostGal.Text)
                If dblCostGal < 0.1 Or dblCostGal > 5 Then

                    ' Display error message.
                    MsgBox("Cost per Gallons must be between .1 and 5 cents", , " input error")
                    txtCostGal.Clear() ' Clear cost / gallon input Textbox.
                    allGood = False ' Not ready for calculations.
                End If
            Catch ex As Exception

                ' Display error message.
                MsgBox("Please enter a valid amount for ""Cost / Gallon""", , " input error")
                txtCostGal.Clear() ' Clear cost / gallon input Textbox.
                allGood = False ' Not ready for calculations.
            End Try
        End If

        ' Check to see if inputs have been validated.
        If allGood = True Then

            ' Calculation correction, sorry.
            dblCostGal = dblCostGal / 100

            ' Calculate totals.
            If water = True Then
                dblTotalHr = (dblGallons * dblCostGal) + (dblCost * dblPower)
                dblTotalYr = ((dblGallons * dblCostGal) + (dblTotalHr * dblHours)) * 365
            Else
                dblTotalHr = dblCost * dblPower
                dblTotalYr = dblTotalHr * dblHours * 365
            End If

            ' Calculate running total.
            dblTotal = dblTotal + dblTotalYr

            ' Display totals.
            txtTotal.Text = FormatCurrency(dblTotal)
            DataGridView1.Rows.Add(cboAppliance.SelectedItem(), dblHours, FormatCurrency(dblTotalHr), FormatCurrency(dblTotalYr))
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.RowCount - 1
        End If

    End Sub

    ' Determine what happens when the user presses the "Exit" button (btnExit).
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close() ' Close the program.
    End Sub

    ' Determine the state of the water inputs.
    Private Sub cboAppliance_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAppliance.SelectedIndexChanged
        If cboAppliance.SelectedItem().Contains("Washer") Then
            txtGallons.ReadOnly = False
            txtCostGal.ReadOnly = False
            water = True
        Else
            txtGallons.ReadOnly = True
            txtCostGal.ReadOnly = True
            txtGallons.Clear() ' Clear gallons input Textbox.
            txtCostGal.Clear() ' Clear cost / gallon input Textbox.
            water = False
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            Dim numCols As Integer = DataGridView1.ColumnCount
            Dim numRows As Integer = DataGridView1.RowCount - 1

            ' Open the file.
            utilityFile = File.CreateText("utilitylist.txt")

            ' Write the headers and the data.
            For count As Integer = 0 To numRows
                For count2 As Integer = 0 To numCols - 1
                    utilityFile.Write(DataGridView1.Columns(count2).HeaderText)
                    utilityFile.Write(" - ")
                    utilityFile.Write(DataGridView1.Rows(count).Cells(count2).Value)
                    utilityFile.WriteLine()
                Next
                utilityFile.WriteLine()
            Next
            ' Write the total to the file.
            utilityFile.Write("Running Total - " & dblTotal)

            ' Close the file.
            utilityFile.Close()
        Catch
            MessageBox.Show("Error: The file cannot be created.")
        End Try
    End Sub
End Class