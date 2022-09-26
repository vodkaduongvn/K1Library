<Serializable()> Public Class clsSearchValue

#Region " Members "

    Private m_blnVariable As Boolean
    Private m_objValue As Object
    Private m_blnConstant As Boolean
#End Region

#Region " Properties "

    Public Property IsVariable() As Boolean
        Get
            Return m_blnVariable
        End Get
        Set(ByVal value As Boolean)
            m_blnVariable = value
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return m_objValue
        End Get
        Set(ByVal value As Object)
            m_objValue = value
        End Set
    End Property

    Public Property IsConstant() As Boolean
        Get
            Return m_blnConstant
        End Get
        Set(ByVal value As Boolean)
            m_blnConstant = value
        End Set
    End Property
#End Region

#Region " Constructors "

    Public Sub New()
    End Sub

    Public Sub New(ByVal objValue As Object, ByVal eTokenType As clsSearchFilter.enumTokenType)
        m_objValue = objValue
        m_blnVariable = (eTokenType = clsSearchFilter.enumTokenType.VARIABLE)
        m_blnConstant = (eTokenType = clsSearchFilter.enumTokenType.CONSTANT)
    End Sub

#End Region

End Class
