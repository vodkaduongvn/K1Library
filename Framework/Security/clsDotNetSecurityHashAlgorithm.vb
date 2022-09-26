Imports System.Security.Cryptography
Imports System.Text

Namespace Framework.Security
    Friend Class clsDotNetSecurityHashAlgorithm
        Implements IHashAlgorithm

        Private ReadOnly m_hashAlgorithmName As String
        Private m_algorithm As HashAlgorithm
        Private ReadOnly m_saltBytes() As Byte

        Sub New(Optional saltValue As String = "", Optional hashAlgorithmName As String = "SHA256")
            m_hashAlgorithmName = hashAlgorithmName
            m_algorithm = CreateHashAlgorithm()
            m_saltBytes = Encoding.UTF8.GetBytes(saltValue)
        End Sub

        Private Function CreateHashAlgorithm() As HashAlgorithm
            Dim hash As HashAlgorithm
            'Initialize appropriate hashing algorithm class.
            Select Case (m_hashAlgorithmName.ToUpper())

                Case "SHA1"
                    hash = New SHA1Managed()
                Case "SHA256"
                    hash = New SHA256Managed()
                Case "SHA384"
                    hash = New SHA384Managed()
                Case "SHA512"
                    hash = New SHA512Managed()
                Case Else
                    hash = New MD5CryptoServiceProvider()
            End Select

            Return hash
        End Function

        ''' <summary>
        ''' Remember to create the Provider with the same hash salt otherwise the Hash value will not match.
        ''' </summary>
        ''' <param name="plainText"></param>
        ''' <param name="hashValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function VerifyHash(ByVal plainText As String, ByVal hashValue As String) As Boolean Implements IHashAlgorithm.VerifyHash

            Return hashValue = ComputeHash(plainText)

        End Function

        ''' <summary>
        ''' Computes Hash With two pinches of salt
        ''' </summary>
        ''' <param name="plainText"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ComputeHash(ByVal plainText As String) As String Implements IHashAlgorithm.ComputeHash

            ' Convert plain text into a byte array.
            Dim plainTextBytes As Byte()
            plainTextBytes = Encoding.UTF8.GetBytes(plainText)

            ' Allocate array, which will hold plain text and salt.
            Dim plainTextWithSaltBytes() As Byte = New Byte(plainTextBytes.Length + m_saltBytes.Length) {}

            Array.Copy(plainTextBytes, plainTextWithSaltBytes, plainTextBytes.Length)

            'Array.Copy(m_saltBytes, 0, plainTextWithSaltBytes, plainTextBytes.Length - 1, plainTextBytes.Length)
            Array.Copy(m_saltBytes, 0, plainTextWithSaltBytes, plainTextBytes.Length - 1, m_saltBytes.Length)

            ' Compute hash value of our plain text with appended salt.
            Dim hashBytes() As Byte = m_algorithm.ComputeHash(plainTextWithSaltBytes)

            Dim hashWithDoubleSaltedBytes As Byte() = New Byte(hashBytes.Length + m_saltBytes.Length - 1) {}

            'Double salting!!! Its cool ummkay ...
            Array.Copy(hashBytes, hashWithDoubleSaltedBytes, hashBytes.Length)
            Array.Copy(m_saltBytes, 0, hashWithDoubleSaltedBytes, hashBytes.Length - 1, m_saltBytes.Length)

            'Lets Base 64 encode it
            Return Convert.ToBase64String(hashBytes)

        End Function

        ''' <summary>
        ''' Keep this salt value value somewhere safe otherwise you won't be able to verify hash values slated with this 
        ''' </summary>
        ''' <param name="minSaltSize"></param>
        ''' <param name="maxSaltSize"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GenerateSalt(Optional minSaltSize As Integer = 4, Optional maxSaltSize As Integer = 8) As String

            ' Generate a random number for the size of the salt.
            Dim random As Random
            random = New Random()

            Dim saltSize As Integer
            saltSize = random.Next(minSaltSize, maxSaltSize)

            ' Allocate a byte array, which will hold the salt.
            Dim saltBytes = New Byte(saltSize - 1) {}

            ' Initialize a random number generator.
            Dim rng As RNGCryptoServiceProvider = New RNGCryptoServiceProvider()

            ' Fill the salt with cryptographically strong byte values.
            rng.GetNonZeroBytes(saltBytes)

            Return Convert.ToBase64String(saltBytes)

        End Function

    End Class

End Namespace