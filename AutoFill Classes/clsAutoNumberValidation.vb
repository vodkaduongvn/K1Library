Public Class clsAutoNumberValidation

#Region " Members "

    Private m_strError As String
    Private m_blnIsAutoGenerated As Boolean = False
    Private m_strMultipleSequenceMask As String
    Private m_blnHasSequence As Boolean = False
    Private m_eErrorType As enumAutonumberErrorType = enumAutonumberErrorType.NONE
#End Region

#Region " Enumerations "

    Public Enum enumAutonumberErrorType
        NONE = 0
        MASK_INCOMPLETE = 1
        OTHER = 2
    End Enum
#End Region

#Region " Properties "

    Public ReadOnly Property ErrorMsg() As String
        Get
            Return m_strError
        End Get
    End Property

    Public Property IsAutoGenerated() As Boolean
        Get
            Return m_blnIsAutoGenerated
        End Get
        Set(ByVal value As Boolean)
            m_blnIsAutoGenerated = value
        End Set
    End Property

    Public Property MultipleSequenceMask() As String
        Get
            Return m_strMultipleSequenceMask
        End Get
        Set(ByVal value As String)
            m_strMultipleSequenceMask = value
        End Set
    End Property

    Public Property HasSequence() As Boolean
        Get
            Return m_blnHasSequence
        End Get
        Set(ByVal value As Boolean)
            m_blnHasSequence = value
        End Set
    End Property

    Public ReadOnly Property ErrorType() As enumAutonumberErrorType
        Get
            Return m_eErrorType
        End Get
    End Property
#End Region

#Region " Methods "

    Public Sub SetError(ByVal strMessage As String, _
    Optional ByVal eType As enumAutonumberErrorType = enumAutonumberErrorType.OTHER)
        m_strError = strMessage
        m_eErrorType = eType
    End Sub
#End Region

End Class
