'Title:         Input Data Validation
'Date:          11-15-14
'Author:        Terry Holmes

'Description:   This Class is used for Data Validation

Option Strict On

Public Class InputDataValidation

    Public Function VerifyText(ByVal strValueForValidation As String) As Boolean

        'Setting local Variable
        Dim blnFatalError As Boolean = False

        'checking to see if there was a value entered
        If strValueForValidation = "" Then

            'Setting variables
            Form1.mstrErrorMessage = "The Value Enter Is Blank"
            blnFatalError = True

        End If

        'Returning Back to calling function
        Return blnFatalError

    End Function

End Class
