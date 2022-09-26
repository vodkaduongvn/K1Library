Public Class clsK1Exception
    Inherits Exception
    Implements IDisposable

#Region " Members "

    Private m_intNumber As Integer
    Private m_blnIncludeFooter As Boolean
    Private m_objDB As clsDB
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
    Private m_objInnerException As Exception
#End Region

#Region " Constants "

    Private Const cMessageFooter As String = "Please contact Knowledgeone Corporation at support@knowledgeonecorp.com"

#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal strMessage As String)
        MyBase.New(strMessage)
    End Sub

    Public Sub New(ByVal intNumber As Integer, ByVal strMessage As String)
        MyBase.New(strMessage)
        m_intNumber = intNumber
    End Sub

    Public Sub New(ByVal intNumber As Integer, ByVal strMessage As String, ByVal objInnerException As Exception)
        MyBase.New(strMessage)
        m_intNumber = intNumber
        m_objInnerException = objInnerException
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal intNumber As Integer, ByVal strMessage As String)
        MyBase.New(strMessage)
        m_objDB = objDB
        m_intNumber = intNumber
    End Sub

    Public Sub New(ByVal strMessage As String, ByVal blnIncludeFooter As Boolean)
        MyBase.New(strMessage & " " & cMessageFooter)
    End Sub

    Public Sub New(ByVal intNumber As Integer, ByVal strMessage As String, ByVal blnIncludeFooter As Boolean)
        MyBase.New(strMessage & " " & cMessageFooter)
        m_intNumber = intNumber
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property ErrorNumber() As Integer
        Get
            Return m_intNumber
        End Get
    End Property

    Public Overrides ReadOnly Property Message() As String
        Get
            '2015-09-02 -- Peter Melisi -- Added new exception clsDB_Direct.enumSQLExceptions.USER_EXCEPTION_FAILURE to 
            ' condition so Message gets parsed correctly.
            If m_objDB IsNot Nothing AndAlso (m_intNumber = clsDB_Direct.enumSQLExceptions.USER_EXCEPTION OrElse m_intNumber = clsDB_Direct.enumSQLExceptions.USER_EXCEPTION_FAILURE) Then
                'try to get the error number from the message and interpret it based on language

                Dim strMessage As String = MyBase.Message
                Dim intNum As Integer
                Dim intPos As Integer = strMessage.IndexOf(" "c)

                If intPos >= 0 Then
                    Dim strFirstWord As String = strMessage.Substring(0, intPos)

                    Integer.TryParse(strFirstWord, intNum)
                Else
                    Integer.TryParse(strMessage, intNum)
                End If

                If intNum > 0 Then
                    Dim objErrMsg As clsErrorMessage = clsErrorMessage.GetItemByUIID(intNum, m_objDB)
                    If Not objErrMsg Is Nothing Then
                        Return objErrMsg.StringObj.GetLanguageString( _
                            m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)
                    ElseIf intPos >= 0 Then
                        Return strMessage.Substring(intPos + 1, strMessage.Length - (intPos + 1))
                    End If
                End If
            End If

            Return MyBase.Message
        End Get
    End Property

    Public Overloads ReadOnly Property InnerException() As Exception
        Get
            Return m_objInnerException
        End Get
    End Property
#End Region

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
