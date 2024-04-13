'Title:         Close The Program
'Date:          11-5-14
'Author:        Terry Holmes

'Description:   This form is used to determine if the program should be closed

Option Strict On

Public Class CloseProgram

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click

        'This will load the variable
        Form1.mblnCloseForm = True
        Me.Close()

    End Sub

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNo.Click

        'This will load the variable
        Form1.mblnCloseForm = False
        Me.Close()

    End Sub
End Class