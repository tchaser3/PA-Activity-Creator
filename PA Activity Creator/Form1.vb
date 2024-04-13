'Title:         Find PA and Activity
'Date:          11-5-14
'Author:        Terry Holmes

'Description:   This form is used to either find an Activity and PA or Create One

Option Strict On

Public Class Form1

    'Setting global variables
    Friend mintCreatedTransactionID As Integer
    Friend mblnCloseForm As Boolean
    Friend mstrErrorMessage As String

    'Setting up the data set
    Private ThePAActivtyDataSet As PAActivtyDataSet
    Private ThePAActivtyDataTier As ModuleDataTier
    Private WithEvents ThePAActivtyBindingSource As BindingSource

    Private editingBoolean As Boolean = False
    Private addingBoolean As Boolean = False
    Private previousSelectedindex As Integer
    Private mblnEditingExisitingRecord As Boolean = False

    'Setting up Selected Index Array to find possible matches
    Dim mintCounter As Integer
    Dim mintSelectedIndex(1000) As Integer
    Dim mintUpperLimit As Integer

    Dim TheInputDataValidation As New InputDataValidation
    
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'This will load up the controls
        Try
            'Loading up the data sets
            ThePAActivtyDataTier = New ModuleDataTier
            ThePAActivtyDataSet = ThePAActivtyDataTier.GetPAActivityInformation
            ThePAActivtyBindingSource = New BindingSource

            'Setting up the Binding Source
            With ThePAActivtyBindingSource
                .DataSource = ThePAActivtyDataSet
                .DataMember = "PAActivty"
                .MoveFirst()
                .MoveLast()
            End With

            'Setting up the combo box
            With cboTransactionID
                .DataSource = ThePAActivtyBindingSource
                .DisplayMember = "TransactionID"
                .DataBindings.Add("text", ThePAActivtyBindingSource, "TransactionID", False, DataSourceUpdateMode.Never)
            End With

            'Setting the rest of the controls
            txtDescription.DataBindings.Add("text", ThePAActivtyBindingSource, "Description")
            txtActivityID.DataBindings.Add("text", ThePAActivtyBindingSource, "ActivityID")

            SetControlsReadOnly(True)

            LoadUpInitialArray()

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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        'This will determine if the program will close
        CloseProgram.ShowDialog()

        If mblnCloseForm = True Then
            Me.Close()
        ElseIf mblnCloseForm = False Then
            MessageBox.Show("The Program Will Not Close", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub
    Private Sub SetControlsReadOnly(ByVal valueBoolean As Boolean)

        'Sets the Controls to Read Only
        txtActivityID.ReadOnly = valueBoolean
        txtDescription.ReadOnly = valueBoolean

    End Sub

    Private Sub btnFindItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindItems.Click

        'Setting Local Variables
        Dim blnFatalError As Boolean
        Dim intKeyWordLengthFromTable As Integer = 0
        Dim intCounter As Integer
        Dim intDescriptionCounter As Integer = 0
        Dim intNumberOfRecords As Integer = 0
        Dim strKeyWordForSearch As String = ""
        Dim strKeyWordFromTable As String = ""
        Dim chaKeyWordInputArray() As Char
        Dim intTransactionIDCounter As Integer = 0
        Dim intKeyWordCounter As Integer = 0
        Dim intCharacterCounter As Integer = 0
        Dim strTempString As String
        Dim intKeyWordLengthForSearch As Integer

        btnNext.Enabled = False
        btnBack.Enabled = False

        strKeyWordForSearch = txtEnterSearchTerm.Text
        strTempString = strKeyWordForSearch
        blnFatalError = TheInputDataValidation.VerifyText(strKeyWordForSearch)

        If blnFatalError = True Then
            MessageBox.Show(mstrErrorMessage, "Check the Enter Search Term", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Routine to Clear the Array
        intNumberOfRecords = cboTransactionID.Items.Count - 1
        blnFatalError = True

        'Clearing the array
        For intCounter = 0 To intNumberOfRecords
            mintSelectedIndex(intCounter) = -1
        Next

        'Getting the length of the for the search
        intKeyWordLengthForSearch = strKeyWordForSearch.Length - 1
        mintCounter = 0
        mintUpperLimit = 0

        'Loop to begin the process
        For intTransactionIDCounter = 0 To intNumberOfRecords

            'Setting the combo box
            cboTransactionID.SelectedIndex = intTransactionIDCounter
            intKeyWordLengthFromTable = txtDescription.Text.Length - 1
            chaKeyWordInputArray = txtDescription.Text.ToCharArray

            'Setting Counter
            intCharacterCounter = intKeyWordLengthForSearch

            'Beginning the Second Loop
            If intKeyWordLengthForSearch <= intKeyWordLengthFromTable Then

                'Beginning the seond loop
                For intDescriptionCounter = 0 To intKeyWordLengthFromTable

                    'Beginning the loop to count the characters
                    For intKeyWordCounter = intDescriptionCounter To intCharacterCounter

                        'Setting up the character array
                        strKeyWordFromTable = strKeyWordFromTable + chaKeyWordInputArray(intKeyWordCounter)

                    Next

                    'Setting up for the next loop
                    If intCharacterCounter < intKeyWordLengthFromTable Then
                        intCharacterCounter += 1
                    End If

                    'comparing the string
                    If strKeyWordForSearch = strKeyWordFromTable Then
                        mintSelectedIndex(mintCounter) = intTransactionIDCounter
                        mintCounter += 1
                        blnFatalError = False
                    End If

                    'Clearing the table string
                    strKeyWordFromTable = ""

                Next

                'Resetting the character counter
                intCharacterCounter = intKeyWordLengthForSearch

            End If

        Next

        If blnFatalError = True Then
            txtEnterSearchTerm.Text = ""
            MessageBox.Show("The Key Word Entered Was Not Found", "Try Again", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        txtEnterSearchTerm.Text = ""

        'Setting up the controls
        mintUpperLimit = mintCounter - 1
        mintCounter = 0

        cboTransactionID.SelectedIndex = mintSelectedIndex(0)

        If mintUpperLimit > 0 Then
            btnNext.Enabled = True
        End If

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        'Setting up local Variables
        Dim strValueForValidation As String
        Dim strControlChecked As String
        Dim blnFatalError As Boolean = False

        'This will run when the Add Button is clicked on
        If btnAdd.Text = "Add" Then

            btnNext.Enabled = False
            btnBack.Enabled = False

            cboTransactionID.Visible = True
            btnEdit.Enabled = False

            'Setting the Binding Source
            With ThePAActivtyBindingSource
                .EndEdit()
                .AddNew()
            End With

            addingBoolean = True
            SetComboBoxBindings()
            If cboTransactionID.SelectedIndex <> -1 Then
                previousSelectedindex = cboTransactionID.Items.Count - 1
            Else
                previousSelectedindex = 0
            End If

            SetControlsReadOnly(False)

            CreateID.Show()
            cboTransactionID.Focus()
            cboTransactionID.Text = CStr(mintCreatedTransactionID)
            btnAdd.Text = "Save"
            mblnEditingExisitingRecord = False

        Else

            'Performing Data Validation
            strValueForValidation = txtDescription.Text
            strControlChecked = "Description"
            blnFatalError = TheInputDataValidation.VerifyText(strValueForValidation)
            If blnFatalError = False Then
                strValueForValidation = txtActivityID.Text
                strControlChecked = "Activity ID"
                blnFatalError = TheInputDataValidation.VerifyText(strValueForValidation)
            End If

            'Performing Check for data validation failure
            If blnFatalError = True Then
                MessageBox.Show(mstrErrorMessage, "Check the " + strControlChecked + " Box", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            'After Validation 
            Try

                'Saving the information
                ThePAActivtyBindingSource.EndEdit()
                ThePAActivtyDataTier.UpdatePAActivityDB(ThePAActivtyDataSet)
                addingBoolean = False
                editingBoolean = False
                cboTransactionID.SelectedIndex = previousSelectedindex
                btnAdd.Text = "Add"
                btnEdit.Enabled = True
                SetControlsReadOnly(True)
                cboTransactionID.Visible = False
                SetComboBoxBindings()

                If mblnEditingExisitingRecord = False Then
                    LoadUpInitialArray()
                End If


            Catch ex As Exception
                MessageBox.Show(ex.Message, "Please Correct", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If

    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

        'This will increment the counter and combo box
        mintCounter += 1
        cboTransactionID.SelectedIndex = mintSelectedIndex(mintCounter)

        'Enabling the button
        btnBack.Enabled = True

        If mintCounter = mintUpperLimit Then
            btnNext.Enabled = False
        End If

    End Sub
    Private Sub LoadUpInitialArray()

        'Creating Local Variables
        Dim intCounter As Integer
        Dim intNumberOfRecords As Integer

        btnNext.Enabled = False
        btnBack.Enabled = False

        intNumberOfRecords = cboTransactionID.Items.Count - 1

        For intCounter = 0 To intNumberOfRecords

            'Incrementing the counter
            cboTransactionID.SelectedIndex = intCounter
            mintSelectedIndex(intCounter) = cboTransactionID.SelectedIndex

        Next

        mintUpperLimit = intNumberOfRecords
        mintCounter = 0
        cboTransactionID.SelectedIndex = 0

        If mintUpperLimit > 0 Then
            btnNext.Enabled = True
        End If

    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        'This will increment the counter and combo box
        mintCounter -= 1
        cboTransactionID.SelectedIndex = mintSelectedIndex(mintCounter)

        'Enabling the button
        btnNext.Enabled = True

        If mintCounter = 0 Then
            btnBack.Enabled = False
        End If

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        SetControlsReadOnly(False)
        editingBoolean = True
        btnAdd.Text = "Save"
        mblnEditingExisitingRecord = True

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        'This Routine will Delete a record
        DeleteConfirmation.ShowDialog()

        If mblnCloseForm = True Then

            ThePAActivtyBindingSource.RemoveCurrent()
            ThePAActivtyDataTier.UpdatePAActivityDB(ThePAActivtyDataSet)
            LoadUpInitialArray()
            MessageBox.Show("The Selected Record Has Been Deleted", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ElseIf mblnCloseForm = False Then

            MessageBox.Show("The Selected Record Was Not Deleted", "Thank You", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
        

    End Sub

    Private Sub txtEnterSearchTerm_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEnterSearchTerm.KeyDown

        'this will allow the user to press enter
        If e.KeyCode = Keys.Enter Then
            btnFindItems.PerformClick()
        End If

    End Sub

    Private Sub btnResetControls_Click(sender As Object, e As EventArgs) Handles btnResetControls.Click

        'This will reset the controls
        LoadUpInitialArray()

    End Sub
End Class
