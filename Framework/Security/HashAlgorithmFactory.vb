Namespace Framework.Security

    Public Enum HashAlgorithmTypes
        SHA1
        SHA256
        SHA384
        SHA512
        MD5
        BCrypt
    End Enum

    Public Class HashAlgorithmFactory

        ''' <summary>
        ''' Create hash algorithms
        ''' </summary>
        ''' <param name="strSalt"></param>
        ''' <param name="hashType"></param>
        ''' <returns></returns>
        ''' <remarks>Not yet implemented!</remarks>
        Public Shared Function CreateHashAlgorithm(hashType As HashAlgorithmTypes, Optional strSalt As String = "") As IHashAlgorithm

            Select Case hashType
                Case HashAlgorithmTypes.BCrypt
                    Return New clsBCryptHashAlgorithm(strSalt)
                Case Else
                    Return New clsDotNetSecurityHashAlgorithm(strSalt, [Enum].GetName(hashType.GetType(), hashType))
            End Select

        End Function

    End Class

End Namespace