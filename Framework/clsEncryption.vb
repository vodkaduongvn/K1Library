Imports System.Security.Cryptography

''' <summary>
''' New Encryption Class designed for .NET 2.0
''' This class is used to implement the Rijndael encryption method
''' </summary>
''' <remarks></remarks>
Friend Class clsEncryption

#Region " Members "

    Private m_objCipher As RijndaelManaged
    Private m_blnUseDefaultIV As Boolean

#End Region

#Region " Constants "

    Private ReadOnly cENCRYPTIONBLOCKSIZE As Integer = 256
    Private ReadOnly cENCRYPTIONSTRENGTH As Integer = 256
#End Region

#Region " Constructors "

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        '-- Create an instance of the cipher
        m_objCipher = New RijndaelManaged

        '-- Set the key and block size
        m_objCipher.KeySize = cENCRYPTIONSTRENGTH
        m_objCipher.BlockSize = cENCRYPTIONBLOCKSIZE
        m_objCipher.Mode = CipherMode.CBC

        '-- TODO: When everything has moved to .NET 2.0 this should be updated to ISO10126 mode 
        '-- As it adds random padding bytes, which reduces the predictability of the plain text.

        '-- Always generate a new IV when encrypting
        m_blnUseDefaultIV = False
    End Sub

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="blnInternal">Use internal key or database key</param>
    ''' <remarks>Will generate a new IV for each encryption</remarks>
    Public Sub New(ByVal blnInternal As Boolean)
        Me.New(blnInternal, False)
    End Sub

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="blnInternal">Use internal key or database key</param>
    ''' <param name="blnUseDefaultIV">Use default IV or generate a new one for each encryption</param>
    ''' <remarks>
    ''' If the encrypted value will be compared to a value later on (ie. password) then used the default IV, 
    ''' otherwise use a dynamic IV as it provides better security
    ''' </remarks>
    Public Sub New(ByVal blnInternal As Boolean, ByVal blnUseDefaultIV As Boolean)
        Me.New()

        Dim strPlainKey As String
        '-- Generates a byte array to use as the key
        'If blnUseOldKey Then
        If blnInternal Then
            strPlainKey = "btnModify"
        Else
            strPlainKey = "waterloo"
        End If
        'Else
        'If blnInternal Then
        '    strPlainKey = "#ZXO*2BOnJMfcR3Hpmj8en%"
        'Else
        '    strPlainKey = "OQ&5hZWjw0O4!Q71ruVOiu"
        'End If
        'End If

        If blnUseDefaultIV Then
            m_objCipher.Padding = PaddingMode.PKCS7
        Else
            m_objCipher.Padding = PaddingMode.ISO10126
        End If

        m_objCipher.Key = Me.GetKey(strPlainKey)
        m_blnUseDefaultIV = blnUseDefaultIV
    End Sub

    Public Sub New(ByVal strKey As String)
        Me.New()

        '-- Generates a byte array from the given key
        m_objCipher.Key = Me.GetKey(strKey)
    End Sub

#End Region

#Region " Properties "

    Public Property UseDefaultIV() As Boolean
        Get
            Return m_blnUseDefaultIV
        End Get
        Set(ByVal value As Boolean)
            m_blnUseDefaultIV = value
        End Set
    End Property

#End Region

#Region " Methods "

#Region " Encrypt "

    ''' <summary>
    ''' Encrypts the given plain text
    ''' </summary>
    ''' <param name="strPlainText">Text to be encrypted</param>
    ''' <returns>Base64 string</returns>
    ''' <remarks></remarks>
    Public Function Encrypt(ByVal strPlainText As String) As String
        Return Convert.ToBase64String(Encrypt(System.Text.Encoding.Unicode.GetBytes(strPlainText)))
    End Function

    ''' <summary>
    ''' Encryts the given array of bytes
    ''' </summary>
    ''' <param name="arrPlainText">Array of UTF8 encoded bytes to be encrypted</param>
    ''' <returns>Encrypted array of UTF8 encoded bytes</returns>
    ''' <remarks></remarks>
    Public Function Encrypt(ByVal arrPlainText As Byte()) As Byte()
        '-- Convert byte array to memory stream for encryption
        Dim objOutputStream As IO.MemoryStream = Encrypt(New IO.MemoryStream(arrPlainText))

        Return objOutputStream.ToArray
    End Function

    ''' <summary>
    ''' Encrypts the given memory stream
    ''' </summary>
    ''' <param name="objInputStream">memory stream to be encrypted</param>
    ''' <returns>Encrypted memory stream</returns>
    ''' <remarks>
    ''' If UseDefaultIV property is False then the generated IV used in the encryption 
    ''' is stored as the first 32 bytes of the memory stream
    ''' </remarks>
    Public Function Encrypt(ByVal objInputStream As IO.MemoryStream) As IO.MemoryStream
        Dim objCryptoStream As CryptoStream = Nothing

        Try
            Dim objOutputStream As New IO.MemoryStream

            If m_blnUseDefaultIV Then
                '-- Passwords need a static IV otherwise login is painful
                m_objCipher.IV = GetDefaultIV()
            Else
                '-- Generate a new vector each time we encrypt 
                '-- otherwise information can be leaked
                m_objCipher.GenerateIV()

                '-- Write the vector array to output stream before encrypted data
                objOutputStream.Write(m_objCipher.IV, 0, m_objCipher.IV.Length)
            End If

            '-- Encrypt plain text and write it to our output stream
            objCryptoStream = New CryptoStream(objOutputStream, m_objCipher.CreateEncryptor(), CryptoStreamMode.Write)
            objCryptoStream.Write(objInputStream.ToArray, 0, objInputStream.ToArray.Length)

            '-- Remember padding!  This instructs the transform to pad and finish.
            objCryptoStream.FlushFinalBlock()

            Return objOutputStream
        Catch ex As Exception
            Throw
        Finally
            '-- Make sure we close the crypto stream before exiting
            If objCryptoStream IsNot Nothing Then objCryptoStream.Close()
        End Try
    End Function
#End Region

#Region " Decrypt "

    ''' <summary>
    ''' Decrypts the given cipher text
    ''' </summary>
    ''' <param name="strCipherText">Base64 text to be decrypted</param>
    ''' <returns>Decrypted text</returns>
    ''' <remarks></remarks>
    Public Function Decrypt(ByVal strCipherText As String) As String
        Return System.Text.Encoding.Unicode.GetString(Decrypt(Convert.FromBase64String(strCipherText)))
    End Function

    ''' <summary>
    ''' Decrypts the given array of bytes
    ''' </summary>
    ''' <param name="arrCipherText">Array of base64 encoded bytes to be decrypted</param>
    ''' <returns>Decrypted array of base64 encoded bytes</returns>
    ''' <remarks></remarks>
    Public Function Decrypt(ByVal arrCipherText As Byte()) As Byte()
        '-- Convert byte array to memory stream for decryption
        Dim objOutputStream As IO.MemoryStream = Decrypt(New IO.MemoryStream(arrCipherText))

        Return objOutputStream.ToArray
    End Function

    ''' <summary>
    ''' Decrypts the given memory stream
    ''' </summary>
    ''' <param name="objInputStream">memory stream to be decrypted</param>
    ''' <returns>Decrypted memory stream</returns>
    ''' <remarks>
    ''' If UseDefaultIV property is False then the first 32 bytes of the memory stream
    ''' must be the IV used in the encryption
    ''' </remarks>
    Private Function Decrypt(ByVal objInputStream As IO.MemoryStream) As IO.MemoryStream
        Dim objCryptoStream As CryptoStream = Nothing

        Try
            Dim arrIV(31) As Byte
            Dim intIVLength As Integer
            Dim objOutputStream As New IO.MemoryStream

            If m_blnUseDefaultIV Then
                '-- Passwords need a static IV otherwise login is painful
                arrIV = GetDefaultIV()
                intIVLength = 0
            Else
                '-- Get the vector used when text was encrypted
                objInputStream.Read(arrIV, 0, arrIV.Length)
                intIVLength = arrIV.Length
            End If

            '-- Set Crypto Stream's IV so we can decrypt
            m_objCipher.IV = arrIV

            '-- Get the cipher text from the input stream
            Dim arrCipherText(CInt(objInputStream.Length - intIVLength - 1)) As Byte
            objInputStream.Read(arrCipherText, 0, arrCipherText.Length)

            '-- Decrypt cipher text and write it to our output stream
            objCryptoStream = New CryptoStream(objOutputStream, m_objCipher.CreateDecryptor(), CryptoStreamMode.Write)
            objCryptoStream.Write(arrCipherText, 0, arrCipherText.Length)

            '-- Remember padding!  This instructs the transform to pad and finish.
            objCryptoStream.FlushFinalBlock()

            Return objOutputStream
        Catch ex As Exception
            Throw
        Finally
            '-- Make sure we close the crypto stream before exiting
            If objCryptoStream IsNot Nothing Then objCryptoStream.Close()
        End Try
    End Function
#End Region

#Region " Private "

#Region " Get Default Key "

    ''' <summary>
    ''' Generates a random byte array that we always used as our key
    ''' </summary>
    ''' <returns>Array of bytes that represents a key</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultKey() As Byte()
        Dim arrKey() As Byte = {&H65, &H98, &H5, &H1A, &H80, &H4F, &HB, &HCE, &H4E, &H5F, &HB1, &H9A, &HEC, &HBC, _
            &H2C, &HB3, &H49, &HF0, &HED, &H8E, &H93, &HBF, &HBE, &HA5, &HEF, &H8, &HF, &H78, &H6D, &HEB, &H8A, &HD1}

        Return arrKey
    End Function

#End Region

#Region " Get Key "

    ''' <summary>
    ''' Generates a byte array by hashing the plain key so we can use it as our key
    ''' </summary>
    ''' <param name="strPlainKey">The text that will be used as the key for encryption</param>
    ''' <returns>Array of bytes that represents a key</returns>
    ''' <remarks></remarks>
    Private Function GetKey(ByVal strPlainKey As String) As Byte()
        Dim objHashSHA As New SHA1CryptoServiceProvider
        Dim arrNewKey() As Byte = GetDefaultKey()

        '-- Convert string to Base64
        strPlainKey = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(strPlainKey))
        Dim arrBase64Bytes() As Byte = Convert.FromBase64String(strPlainKey)
        Dim arrHashBytes() As Byte = objHashSHA.ComputeHash(arrBase64Bytes)

        '-- Override default key values with hash values
        Array.Copy(arrHashBytes, arrNewKey, Math.Min(arrHashBytes.Length, arrNewKey.Length))

        Return arrNewKey
    End Function

#End Region

#Region " Get Default IV "

    ''' <summary>
    ''' Generates a static Initalise Vector
    ''' </summary>
    ''' <returns>Array of bytes to be used for IV</returns>
    ''' <remarks>This is used for encryptions when the property UseDefaultIV is True</remarks>
    Private Function GetDefaultIV() As Byte()
        Dim arrIV() As Byte = {&HE, &HA5, &HCA, &H4, &H13, &HA6, &HD8, &HE5, &H80, &H5C, &HF3, &H1C, &H88, &H86, &HF0, &HF5, _
            &HC9, &H97, &H84, &H24, &HC3, &H5D, &H59, &H78, &H8E, &HE0, &H60, &HED, &HCC, &H8D, &H69, &HE9}

        Return arrIV
    End Function

#End Region

#Region " To Hex "

    ''' <summary>
    ''' Converts a byte array into a hexadecimal string for debuging purposes
    ''' </summary>
    ''' <param name="arrBytes">Array for bytes that will be converted</param>
    ''' <returns>Hexadecimal string computed from byte array</returns>
    ''' <remarks></remarks>
    Private Shared Function ToHex(ByVal arrBytes() As Byte) As String
        If arrBytes Is Nothing OrElse arrBytes.Length = 0 Then
            Return ""
        End If
        Const cHexFormat As String = "{0:X2} "
        Dim objStringBuilder As New System.Text.StringBuilder
        For Each bytValue As Byte In arrBytes
            objStringBuilder.Append(String.Format(cHexFormat, bytValue))
        Next
        Return objStringBuilder.ToString
    End Function

#End Region

#End Region

#End Region

End Class