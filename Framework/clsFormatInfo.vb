Public Class clsFormatInfo

#Region " Members "

    Private m_eFormat As clsDBConstants.enumFormatType
    Private m_strFormat As String
    Private m_strKey As String
#End Region

#Region " Constructors "

    Public Sub New(ByVal eFormat As clsDBConstants.enumFormatType, ByVal strFormat As String, ByVal strKey As String)
        m_eFormat = eFormat
        m_strFormat = strFormat
        m_strKey = strKey
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Format() As clsDBConstants.enumFormatType
        Get
            Return m_eFormat
        End Get
    End Property

    Public ReadOnly Property FormatString() As String
        Get
            Return m_strFormat
        End Get
    End Property

    Public ReadOnly Property Key() As String
        Get
            Return m_strKey
        End Get
    End Property
#End Region

End Class
