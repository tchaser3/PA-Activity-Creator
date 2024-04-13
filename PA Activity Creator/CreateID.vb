'Title:         Create ID
'Date:          11-5-14
'Author:        Terry Holmes

'Description:   This form is used to create an ID

Option Strict On

Public Class CreateID

    'Setting up the data set
    Private TheCreatedIDDataSet As CreatedIDDataSet
    Private TheCreatedIDDataTier As ModuleDataTier
    Private WithEvents TheCreatedIDBindingSource As BindingSource

    Private editingBoolean As Boolean = False
    Private addingBoolean As Boolean = False
    Private previousSelectedindex As Integer

    Private Sub CreateID_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Setting local variables
        Dim intCreatedTransactionID As Integer

        'This will set the controls
        Try

            'Setting up the data sets
            TheCreatedIDDataTier = New ModuleDataTier
            TheCreatedIDDataSet = TheCreatedIDDataTier.GetCreatedIDInformation
            TheCreatedIDBindingSource = New BindingSource

            'Setting up the binding source
            With TheCreatedIDBindingSource
                .DataSource = TheCreatedIDDataSet
                .DataMember = "CreateID"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = TheCreatedIDBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", TheCreatedIDBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Setting up the rest of the controls\
            txtCreatedID.DataBindings.Add("text", TheCreatedIDBindingSource, "CreatedTransactionID")

            'Creating the ID
            intCreatedTransactionID = cboTransactionID.Items.Count + 1000

            'Creating the new record
            With TheCreatedIDBindingSource
                .EndEdit()
                .AddNew()
            End With

            'Setting up to fill the record
            addingBoolean = True
            SetComboBoxBindings()
            If cboTransactionID.SelectedIndex <> -1 Then
                previousSelectedindex = cboTransactionID.Items.Count - 1
            Else
                previousSelectedindex = 0
            End If
            cboTransactionID.Text = CStr(intCreatedTransactionID)
            txtCreatedID.Text = CStr(intCreatedTransactionID)

            'Saving the record
            Form1.mintCreatedTransactionID = intCreatedTransactionID

            'Updating the data set
            TheCreatedIDBindingSource.EndEdit()
            TheCreatedIDDataTier.UpdateCreatedIDDB(TheCreatedIDDataSet)
            addingBoolean = False
            editingBoolean = False
            SetComboBoxBindings()
            cboTransactionID.SelectedIndex = previousSelectedindex

            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub SetComboBoxBindings()

        'This sets up the bindings for the combo box
        With Me.cboTransactionID
            If (addingBoolean Or editingBoolean) Then
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.OnValidation
                .DropDownStyle = ComboBoxStyle.DropDown
            Else
                .DataBindings!text.DataSourceUpdateMode = DataSourceUpdateMode.Never
                .DropDownStyle = ComboBoxStyle.DropDownList
            End If
        End With

    End Sub
End Class