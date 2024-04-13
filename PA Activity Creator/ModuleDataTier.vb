'Title:         Module Data Tier
'Date:          11-5-14
'Author:        Terry Holmes

'Description:   This module is used for the PA and Activity number for the Cable Report

Option Strict On

Public Class ModuleDataTier


    Private aPAActivityTableAdapter As PAActivtyDataSetTableAdapters.PAActivtyTableAdapter
    Private aPAActivityDataSet As PAActivtyDataSet

    Private aCreatedIDTableAdapter As CreatedIDDataSetTableAdapters.CreateIDTableAdapter
    Private aCreatedIDDataSet As CreatedIDDataSet

    Public Function GetPAActivityInformation() As PAActivtyDataSet

        'This will Get the employee Information
        Try
            'Setting up variable
            aPAActivityDataSet = New PAActivtyDataSet
            aPAActivityTableAdapter = New PAActivtyDataSetTableAdapters.PAActivtyTableAdapter
            aPAActivityTableAdapter.Fill(aPAActivityDataSet.PAActivty)

            'Return to VehicleDailyInspection            
            Return aPAActivityDataSet
        Catch ex As Exception

            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aPAActivityDataSet

        End Try
    End Function
    Public Sub UpdatePAActivityDB(ByVal aPAActivityDataSet As PAActivtyDataSet)

        'This will update the Data Set
        Try
            aPAActivityTableAdapter.Update(aPAActivityDataSet.PAActivty)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetCreatedIDInformation() As CreatedIDDataSet

        'This will Get the employee Information
        Try
            'Setting up variables
            aCreatedIDDataSet = New CreatedIDDataSet
            aCreatedIDTableAdapter = New CreatedIDDataSetTableAdapters.CreateIDTableAdapter
            aCreatedIDTableAdapter.Fill(aCreatedIDDataSet.CreateID)

            'Return to VehicleDailyInspection            
            Return aCreatedIDDataSet
        Catch ex As Exception

            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return aCreatedIDDataSet

        End Try
    End Function
    Public Sub UpdateCreatedIDDB(ByVal aCreatedIDDataSet As CreatedIDDataSet)

        'This will update the Data Set
        Try
            aCreatedIDTableAdapter.Update(aCreatedIDDataSet.CreateID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
