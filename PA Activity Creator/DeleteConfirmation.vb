'Title:         Delete Confirmation
'Date:          11/6/14
'Author;        Terry Holmes

'Description:   This Form Will confirm that the record will be deleted

Option Strict On

Public Class DeleteConfirmation

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click

        'This will Set the variable
        Form1.mblnCloseForm = True
        Me.Close()

    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        'This will Set the variable
        Form1.mblnCloseForm = False
        Me.Close()

    End Sub
End Class