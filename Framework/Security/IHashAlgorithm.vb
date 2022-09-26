Namespace Framework.Security
    Public Interface IHashAlgorithm

        Function ComputeHash(ByVal plainText As String) As String
        Function VerifyHash(ByVal hashedValue As String, ByVal plainText As String) As Boolean

    End Interface
End NameSpace