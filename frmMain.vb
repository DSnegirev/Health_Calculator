'***************************************************************
' Programmer: Domentyan Snegirev
' Chemeketa Community College
' Date: 3/13/2016
' Class: CIS133VB
' Assigment: Final Project - Health Calculator.
' File Name: frmMain.vb
' Description: An application that calculates many well-known
'              health issues, including smoking.
'***************************************************************

Public Class frmMain

    ' Declare variables.
    Protected Friend intAge As Integer
    Protected Friend dblHeight As Double
    Protected Friend dblWeight As Double
    Protected Friend dblActivityLevel As Double
    Protected Friend strGender As String

    Protected Friend dblBMR As Double
    Protected Friend dblBMI As Double
    Protected Friend dblCaloriesNeeded As Double
    Protected Friend dblHeartRate As Double

    Protected Friend intSmokingYears As Integer
    Protected Friend dblPacksADay As Double
    Protected Friend dblSmokeTotal As Double
    Protected Friend dblDaysLost As Double
    Protected Friend decMoneyLost As Decimal


    ' Calculate BMR in English Units.
    Private Function CalculateEnglishBMR()
        ' Calculate Male BMR.
        If radMale.Checked = True Then
            dblBMR = 66 + (6.23 * dblWeight) + (12.7 * dblHeight) - (6.8 * intAge)
            ' Calculate Female BMR.
        ElseIf radFemale.Checked = True Then
            dblBMR = 655 + (4.35 * dblWeight) + (4.7 * dblHeight) - (4.7 * intAge)
        End If

        Return dblBMR
    End Function

    ' Calculate BMR in Metric Units.
    Private Function CalculateMetricBMR()
        ' Calculate Male BMR.
        If radMale.Checked = True Then
            dblBMR = 66 + (13.7 * dblWeight) + (5 * dblHeight) - (6.8 * intAge)
            ' Calculate Female BMR.
        ElseIf radFemale.Checked = True Then
            dblBMR = 655 + (9.6 * dblWeight) + (1.8 * dblHeight) - (4.7 * intAge)
        End If

        Return dblBMR
    End Function

    ' Calculate the calories needed.
    Private Function CalculateCaloriesNeeded()
        dblCaloriesNeeded = dblBMR * dblActivityLevel
        Return dblCaloriesNeeded
    End Function

    ' Calculate BMI in English Units.
    Private Function CalculateEnglishBMI()
        dblBMI = (dblWeight * 703) / (dblHeight * dblHeight)
        Return dblBMI
    End Function

    ' Calculate BMI in Metric Units.
    Private Function CalculateMetricBMI()
        Dim dblHeightInMeters As Double
        dblHeightInMeters = dblHeight / 100
        dblBMI = (dblWeight) / (dblHeightInMeters * dblHeightInMeters)
        Return dblBMI
    End Function

    ' Calculate Heart Rate.
    Private Function CalculateHeartRate()
        dblHeartRate = 220 - intAge
        Return dblHeartRate
    End Function

    ' Calculate Smoking.
    Private Sub CalculateSmoking()
        dblSmokeTotal = CDbl((dblPacksADay * 365.24) * intSmokingYears)

        ' Average cost of a cigarette pack is $4.96.
        decMoneyLost = CDec(4.96 * dblSmokeTotal)

        ' Every cigarette smoked loses 11 minutes of your life.
        ' Every pack(20 cigs) loses 220 minutes of your life.
        ' 1440 minutes in 1 day.
        dblDaysLost = (dblSmokeTotal * 220) / 1440
    End Sub

    ' Overall method that does all the Calculations.
    Private Sub Calculate()
        ' Does english calculations.
        If mnuMeasurementsEnglish.Checked = True Then
            CalculateEnglishBMR()
            CalculateEnglishBMI()

            ' Does metric calculations
        ElseIf mnuMeasurementsMetric.Checked = True Then
            CalculateMetricBMR()
            CalculateMetricBMI()
        End If

        ' Does the rest of the calculations.
        CalculateCaloriesNeeded()
        CalculateHeartRate()
        CalculateSmoking()
    End Sub

    ' The ClearForm procedure clears the form.
    Private Sub ClearForm()
        ' Clear the text boxes.
        txtAge.Clear()
        txtHeight.Clear()
        txtWeight.Clear()
        txtSmokePacks.Clear()
        txtSmokeYears.Clear()

        ' Reset radio buttons.
        radMale.Checked = True
        radNo.Checked = True

        ' Reset combo box.
        cmbActivity.SelectedIndex = -1

        ' Set the focus.
        txtAge.Focus()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ' Close the form.
        Me.Close()
    End Sub

    Private Sub mnuMeasurementsEnglish_Click(sender As Object, e As EventArgs) Handles mnuMeasurementsEnglish.Click
        ' Changes label units to English units.
        lblHeightLabel.Text = "Height(in):"
        lblWeightLabel.Text = "Weight(lb):"

        mnuMeasurementsMetric.Checked = False
    End Sub

    Private Sub mnuMeasurementsMetric_Click(sender As Object, e As EventArgs) Handles mnuMeasurementsMetric.Click
        ' Changes label units to Metric units.
        lblHeightLabel.Text = "Height(cm):"
        lblWeightLabel.Text = "Weight(kg):"

        mnuMeasurementsEnglish.Checked = False
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        ' Create an instance of the DisplayForm form.
        Dim frmAboutBox As New AboutBox

        frmAboutBox.ShowDialog()
    End Sub

    Private Sub radYes_CheckedChanged(sender As Object, e As EventArgs) Handles radYes.CheckedChanged
        ' Enable the labels and text boxes.
        lblSmokePacks.Enabled = True
        lblSmokeYears.Enabled = True
        txtSmokePacks.Enabled = True
        txtSmokeYears.Enabled = True
    End Sub

    Private Sub radNo_CheckedChanged(sender As Object, e As EventArgs) Handles radNo.CheckedChanged
        ' Disable the labels and text boxes.
        lblSmokePacks.Enabled = False
        lblSmokeYears.Enabled = False
        txtSmokePacks.Enabled = False
        txtSmokeYears.Enabled = False
    End Sub

    Protected Friend Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        ' Try/Catch for overall text boxes.
        Try
            ' When user clicks no to smoking, doesn't try to validate
            ' "false" text boxes that are disabled.
            If radYes.Checked = True Then
                dblPacksADay = CDbl(txtSmokePacks.Text)
                intSmokingYears = CInt(txtSmokeYears.Text)
            End If

            ' Get data from the user.
            intAge = CInt(txtAge.Text)
            dblHeight = CDbl(txtHeight.Text)
            dblWeight = CDbl(txtWeight.Text)

            ' Set the strGender string.
            If radMale.Checked = True Then
                strGender = "Male"
            ElseIf radFemale.Checked = True Then
                strGender = "Female"
            End If

            ' Try/Catch for Activity Level combo box.
            Try
                If cmbActivity.SelectedIndex = 0 Then
                    dblActivityLevel = 1.2
                ElseIf cmbActivity.SelectedIndex = 1 Then
                    dblActivityLevel = 1.375
                ElseIf cmbActivity.SelectedIndex = 2 Then
                    dblActivityLevel = 1.55
                ElseIf cmbActivity.SelectedIndex = 3 Then
                    dblActivityLevel = 1.725
                ElseIf cmbActivity.SelectedIndex = 4 Then
                    dblActivityLevel = 1.9
                ElseIf cmbActivity.SelectedIndex = -1 Then
                    Throw New ArgumentException()
                End If

                ' Performs all the Calculations.
                Calculate()

                ' Create an instance of the frmCalculations form.
                Dim frmCalculations As New frmCalculations

                ' Display the form.
                frmCalculations.ShowDialog()
            Catch ex As ArgumentException
                ' Error message for not selecting activity level.
                MessageBox.Show("Please select an activity level.")
            End Try

        Catch ex As Exception
            MessageBox.Show("Please fill out all text boxes with numbers.")
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clear the form.
        ClearForm()
    End Sub

End Class
