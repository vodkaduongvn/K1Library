#Region " File Information "

'=====================================================================
' This class is a generic parameter object
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date        Description
'---------------------------------------------------------------------
' KSD       11/06/2004  Implemented.
'=====================================================================

#End Region

#End Region

Imports System.Xml.Serialization

<Serializable()> Public Class clsDBParameter
    Implements IDisposable

#Region " Member Variables "

    Private m_strName As String
    Private m_objValue As Object
    Private m_eDirection As System.Data.ParameterDirection
    Private m_eDataType As SqlDbType
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_strName = ""
        m_objValue = Nothing
        m_eDirection = ParameterDirection.Input
    End Sub

    Public Sub New(ByVal strName As String, ByVal objValue As Object)
        Me.New()
        m_strName = strName
        m_objValue = objValue
    End Sub

    Public Sub New(ByVal strName As String, ByVal objValue As Object, ByVal eDirection As System.Data.ParameterDirection)
        Me.New()
        m_strName = strName
        m_objValue = objValue
        m_eDirection = eDirection
    End Sub

    Public Sub New(ByVal strName As String, ByVal objValue As Object, _
    ByVal eDirection As System.Data.ParameterDirection, ByVal eDataType As SqlDbType)
        Me.New()
        m_strName = strName
        m_objValue = objValue
        m_eDirection = eDirection
        m_eDataType = eDataType
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' The name expected in the Stored Procedure for the parameter
    ''' </summary>
    Public Property Name() As String
        Get
            Return m_strName
        End Get
        Set(ByVal Value As String)
            m_strName = Value
        End Set
    End Property

    ''' <summary>
    ''' The value to be passed into the SP
    ''' </summary>
    Public Property Value() As Object
        Get
            Return m_objValue
        End Get
        Set(ByVal Value As Object)
            m_objValue = Value
        End Set
    End Property

    ''' <summary>
    ''' Input, Output, or both for a Stored Procedure Param
    ''' </summary>
    Public Property Direction() As System.Data.ParameterDirection
        Get
            Return m_eDirection
        End Get
        Set(ByVal Value As System.Data.ParameterDirection)
            m_eDirection = Value
        End Set
    End Property

    Public Property SqlDBType() As SqlDbType
        Get
            Return m_eDataType
        End Get
        Set(ByVal Value As SqlDbType)
            m_eDataType = Value
        End Set
    End Property

    Public ReadOnly Property DBType() As DbType
        Get
            Select Case m_eDataType
                Case SqlDBType.BigInt
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Binary
                    Return DBType.Binary
                Case SqlDBType.Bit
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Char
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.DateTime
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Decimal
                    Return DBType.Decimal
                Case SqlDBType.Float
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Image
                    Return DBType.Binary
                Case SqlDBType.Int
                    Return DBType.Int32
                Case SqlDBType.Money
                    Return DBType.Currency
                Case SqlDBType.NChar
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.NText
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.NVarChar
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Real
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.SmallDateTime
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.SmallInt
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.SmallMoney
                    Return DBType.Currency
                Case SqlDBType.Text
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Timestamp
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.TinyInt
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.UniqueIdentifier
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.VarBinary
                    Return DBType.Binary
                Case SqlDBType.VarChar
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case SqlDBType.Variant
                    Return CType(clsDBConstants.cintNULL, DbType)
                Case Else
                    Return CType(clsDBConstants.cintNULL, DbType)
            End Select
        End Get
    End Property
#End Region

#Region " Methods "

    Public Function Serialize() As Byte()
        Dim objWriter As System.IO.MemoryStream
        Dim objSOAPFormatter As System.Runtime.Serialization.Formatters.Soap.SoapFormatter
        'Dim objBinaryFormatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        Try
            ' Create a string writer to store the serialized object
            objWriter = New System.IO.MemoryStream

            ' Create an XML Serializer to serialize the object
            objSOAPFormatter = New System.Runtime.Serialization.Formatters.Soap.SoapFormatter
            'objBinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

            ' Serialize the object
            'objBinaryFormatter.Serialize(objWriter, Me)
            objSOAPFormatter.Serialize(objWriter, Me)

            ' Close the writer.
            objWriter.Close()
        Catch ex As Exception
            Throw
        End Try

        Return objWriter.GetBuffer
    End Function

    Public Shared Function Deserialize(ByVal arrSerialized As Byte()) As clsDBParameter
        Dim objReader As System.IO.MemoryStream
        Dim objSOAPFormatter As System.Runtime.Serialization.Formatters.Soap.SoapFormatter
        Dim objParam As clsDBParameter

        Try
            ' Create a string writer to store the serialized object
            objReader = New System.IO.MemoryStream(arrSerialized)

            ' Create an XML Serializer to serialize the object
            objSOAPFormatter = New System.Runtime.Serialization.Formatters.Soap.SoapFormatter

            ' Serialize the object
            objParam = CType(objSOAPFormatter.Deserialize(objReader), clsDBParameter)

            ' Close the writer.
            objReader.Close()
        Catch ex As Exception
            Throw
        End Try

        Return objParam
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objValue = Nothing
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
