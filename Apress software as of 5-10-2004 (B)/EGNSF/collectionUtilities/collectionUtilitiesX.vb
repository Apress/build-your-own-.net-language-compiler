Option Explicit On 

Friend Class collectionUtilitiesX

    ' *****************************************************************
    ' *                                                               *
    ' * Non-Strict extensions to the collectionUtilities              *
    ' *                                                               *
    ' *****************************************************************

    ' -----------------------------------------------------------------
    ' Attempt to clone the object if it exposes a clone
    '
    '
    Public Shared Function cloneAttempt(ByVal objObject As Object) As Object
        Try
            Return objObject.clone
        Catch
            Return Nothing
        End Try
    End Function

End Class
