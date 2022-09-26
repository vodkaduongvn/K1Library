<Serializable()> Public Class clsFileInfo

#Region " Members "

    Private m_strFile As String
    Private m_intFileSize As Integer
    Private m_intChunkSize As Integer
    Private m_intIndex As Integer
    Private m_arrBytes As Byte()
    Private m_blnFileComplete As Boolean
#End Region

#Region " Constructors "

    Public Sub New(ByVal strFile As String, ByVal intChunkSize As Long)
        m_blnFileComplete = False
        m_intFileSize = 0
        m_intIndex = 0
        m_strFile = strFile
        m_intChunkSize = CInt(intChunkSize)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property File() As String
        Get
            Return m_strFile
        End Get
    End Property

    Public ReadOnly Property FileSize() As Integer
        Get
            Return m_intFileSize
        End Get
    End Property

    Public ReadOnly Property FileComplete() As Boolean
        Get
            Return m_blnFileComplete
        End Get
    End Property

    Public ReadOnly Property ByteArray() As Byte()
        Get
            Return m_arrBytes
        End Get
    End Property
#End Region

#Region " Methods "

#Region " Read File "

    Public Sub ReadFile()
        Dim objFileStream As IO.FileStream = IO.File.OpenRead(m_strFile)

        If m_intFileSize <= 0 Then
            m_intFileSize = CInt(objFileStream.Length)
        End If

        If m_intIndex + m_intChunkSize >= m_intFileSize Then
            m_intChunkSize = m_intFileSize - m_intIndex
            m_blnFileComplete = True
        End If

        If m_intChunkSize <= 0 Then
            Return
        End If

        ReDim m_arrBytes(m_intChunkSize - 1)

        objFileStream.Position = m_intIndex
        objFileStream.Read(m_arrBytes, 0, m_intChunkSize)
        objFileStream.Close()

        m_intIndex += m_intChunkSize
    End Sub
#End Region

#Region " Serializable "

    Public Function Serialize() As Byte()
        Dim objWriter As System.IO.MemoryStream
        Dim objBinaryFormatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        Try
            ' Create a string writer to store the serialized object
            objWriter = New System.IO.MemoryStream

            ' Create an XML Serializer to serialize the object
            objBinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

            ' Serialize the object
            objBinaryFormatter.Serialize(objWriter, Me)

            ' Close the writer.
            objWriter.Close()
        Catch ex As Exception
            Throw
        End Try

        Return objWriter.GetBuffer
    End Function

    Public Shared Function Deserialize(ByVal arrSerialized As Byte()) As clsFileInfo
        Dim objReader As System.IO.MemoryStream
        Dim objBinaryFormatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim objFileInfo As clsFileInfo

        Try
            ' Create a string writer to store the serialized object
            objReader = New System.IO.MemoryStream(arrSerialized)

            ' Create an XML Serializer to serialize the object
            objBinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

            ' Serialize the object
            objFileInfo = CType(objBinaryFormatter.Deserialize(objReader), clsFileInfo)

            ' Close the writer.
            objReader.Close()
        Catch ex As Exception
            Throw
        End Try

        Return objFileInfo
    End Function
#End Region

#End Region

End Class
