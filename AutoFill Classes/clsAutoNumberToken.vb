Public Class clsAutoNumberToken

#Region " Members "

    Private m_eTokenType As enumTokenType
    Private m_strValue As String
#End Region

#Region " Enumerations "

    Public Enum enumTokenType
        LETTER = 1
        NUMBER = 2
        ANY = 3
        SEQUENCE = 4
        DAY = 5
        DAY_PAD = 6
        MONTH = 7
        MONTH_PAD = 8
        YEAR_2 = 9
        YEAR_4 = 10
        CONSTANT = 11
        ANY_NOSPACE = 12
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal eTokenType As enumTokenType, _
    Optional ByVal strValue As String = Nothing)
        m_eTokenType = eTokenType
        m_strValue = strValue
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property TokenType() As enumTokenType
        Get
            Return m_eTokenType
        End Get
    End Property

    Public Property Value() As String
        Get
            Return m_strValue
        End Get
        Set(ByVal value As String)
            m_strValue = value
        End Set
    End Property
#End Region

End Class
