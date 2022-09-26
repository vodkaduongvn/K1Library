Public Class clsSentence

#Region " Members "

    Private m_intValue As Int64
    Private m_strSentence As String
    Private m_strProcessedSentence As String
    Private m_colFoundWords As List(Of String)

#End Region

#Region " Constructors "

    Public Sub New(ByVal strSentence As String, ByVal intValue As Int64, ByVal colFoundWords As List(Of String))
        m_strSentence = strSentence
        m_strProcessedSentence = RemovePunctuations(strSentence)
        m_intValue = intValue
        m_colFoundWords = colFoundWords
    End Sub

#End Region

#Region "Properties "

#Region " Sentence "

    Public Property Sentence() As String
        Get
            Return m_strSentence
        End Get
        Set(ByVal Value As String)
            m_strSentence = Value
            m_strProcessedSentence = RemovePunctuations(m_strSentence)
        End Set
    End Property

#End Region

#Region " Hit Score "

    Public Property HitScore() As Int64
        Get
            Return m_intValue
        End Get
        Set(ByVal Value As Int64)
            m_intValue = Value
        End Set
    End Property

#End Region

#Region " Processed Sentence "

    Public ReadOnly Property ProcessedSentence() As String
        Get
            Return m_strProcessedSentence
        End Get
    End Property

#End Region

    Public Property FoundSeedWords() As List(Of String)
        Get
            Return m_colFoundWords
        End Get
        Set(ByVal value As List(Of String))
            m_colFoundWords = value
        End Set
    End Property

#End Region

#Region " Remove Punctuations "

    ''' <summary>
    ''' Removes unwanted punctuations so we can match words correctly against each other.
    ''' Currently only removes "," and "."
    ''' </summary>
    ''' <param name="strSentence"></param>
    Public Shared Function RemovePunctuations(ByVal strSentence As String) As String
        Try
            'TraceEnter(New StackTrace(True))

            strSentence = Replace(strSentence, ",", " ")
            strSentence = Replace(strSentence, ".", " ")

            Return strSentence
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class
