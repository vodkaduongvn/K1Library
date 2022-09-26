Namespace Framework.Security
    Friend Class clsBCryptHashAlgorithm
        Implements IHashAlgorithm

        ReadOnly m_salt As String

        ''' <summary>
        ''' If you supply the salt value make sure to store it somewhere else hashed value will not be verified
        ''' </summary>
        ''' <param name="salt"></param>
        ''' <remarks></remarks>
        Sub New(Optional salt As String = "")
            m_salt = salt
        End Sub

        Public Function ComputeHash(ByVal plainText As String) As String Implements IHashAlgorithm.ComputeHash
            Return BCrypt.Net.BCrypt.HashString(plainText & m_salt)
        End Function

        Public Function VerifyHash(ByVal hashedValue As String, ByVal plainText As String) As Boolean Implements IHashAlgorithm.VerifyHash
            Return BCrypt.Net.BCrypt.Verify(plainText & m_salt, hashedValue)
        End Function

    End Class
End Namespace