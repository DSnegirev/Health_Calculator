'***************************************************************
' Programmer: Domentyan Snegirev
' Chemeketa Community College
' Date: 3/13/2016
' Class: CIS133VB
' Assigment: Final Project - Health Calculator.
' File Name: frmMain.vb
' Description: Displays all the information that is calculated
'              in the main form.
'***************************************************************

Imports System.IO

Public Class frmCalculations
    Private Sub frmCalculations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Sets the general information.
        lblAge.Text = frmMain.intAge
        lblHeight.Text = frmMain.dblHeight
        lblWeight.Text = frmMain.dblWeight
        lblGender.Text = frmMain.strGender

        ' Sets the Smoking information.
        lblSmokesPerDay.Text = frmMain.dblPacksADay.ToString("#.##")
        lblTotal.Text = frmMain.dblSmokeTotal.ToString("#.##")
        lblMoneyLost.Text = frmMain.decMoneyLost.ToString("c")
        lblDaysLost.Text = frmMain.dblDaysLost.ToString("#.##")

        ' Sets the BMR information.
        lblBMR.Text = frmMain.dblBMR.ToString("#.##")
        lblCaloriesNeeded.Text = frmMain.dblCaloriesNeeded.ToString("#.##")

        ' Sets the BMI information.
        lblBMI.Text = frmMain.dblBMI.ToString("#.##")
        lblHeartRate.Text = frmMain.dblHeartRate
    End Sub

    Private Sub mnuFilePrint_Click(sender As Object, e As EventArgs) Handles mnuFilePrint.Click
        ' Print the current document.
        pdPrint.Print()
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        ' Close the form.
        Me.Close()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        ' Create an instance of the AboutBox.
        Dim AboutBox As New AboutBox

        AboutBox.ShowDialog()
    End Sub
End Class