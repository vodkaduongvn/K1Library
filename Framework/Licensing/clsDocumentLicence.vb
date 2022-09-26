Namespace Licensing

    Public Class clsDocumentLicence

#Region " Members "

        Private m_intSerialNumber As Integer
        Private m_dtStartDate As Date
        Private m_intNumberOfDocuments As Integer
        Private m_eDocLicType As enumDocLicType
#End Region

#Region " Enumerations "

        Public Enum enumDocLicType
            VALID_YEAR = 0
            VALID_ALWAYS = 1
        End Enum
#End Region

#Region " Constructors "

        Public Sub New(ByVal intSerialNumber As Integer, ByVal dtStartDate As Date, _
        ByVal intNumberOfDocuments As Integer, ByVal eDocLicType As enumDocLicType)
            m_intSerialNumber = intSerialNumber
            m_dtStartDate = dtStartDate
            m_intNumberOfDocuments = intNumberOfDocuments
            m_eDocLicType = eDocLicType
        End Sub
#End Region

#Region " Properties "

        Public ReadOnly Property SerialNumber() As Integer
            Get
                Return m_intSerialNumber
            End Get
        End Property

        Public ReadOnly Property StartDate() As Date
            Get
                Return m_dtStartDate
            End Get
        End Property

        Public ReadOnly Property NumberOfDocuments() As Integer
            Get
                Return m_intNumberOfDocuments
            End Get
        End Property

        Public ReadOnly Property DocLicType() As enumDocLicType
            Get
                Return m_eDocLicType
            End Get
        End Property
#End Region

    End Class

End Namespace

