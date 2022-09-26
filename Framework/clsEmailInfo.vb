Public Class clsEmailInfo
    Implements IDisposable

#Region " Members "

    Private m_strRootFolder As String
    Private m_eMailType As enumMailType
    Private m_arrFiles As ArrayList
    Private m_strMessage As String
    Private m_strSMTPServer As String
    Private m_blnAsHTML As Boolean
    Private m_blnDisposedValue As Boolean = False
#End Region

#Region " Enumerations "

    Public Enum enumMailType
        SMTP = 1
        MAPI = 2
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal eMailType As enumMailType, ByVal strRootFolder As String, _
    Optional ByVal strSMTPServer As String = Nothing)
        m_eMailType = eMailType
        m_strSMTPServer = strSMTPServer
        m_strRootFolder = strRootFolder
    End Sub
#End Region

#Region " Properties "

    Public Property Files() As ArrayList
        Get
            Return m_arrFiles
        End Get
        Set(ByVal value As ArrayList)
            m_arrFiles = value
        End Set
    End Property

    Public Property Message() As String
        Get
            Return m_strMessage
        End Get
        Set(ByVal value As String)
            m_strMessage = value
        End Set
    End Property

    Public ReadOnly Property MailType() As enumMailType
        Get
            Return m_eMailType
        End Get
    End Property

    Public ReadOnly Property SMTPServer() As String
        Get
            Return m_strSMTPServer
        End Get
    End Property

    Public ReadOnly Property RootFolder() As String
        Get
            Return m_strRootFolder
        End Get
    End Property

    Public Property SendAsHTML() As Boolean
        Get
            Return m_blnAsHTML
        End Get
        Set(ByVal value As Boolean)
            m_blnAsHTML = value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Sub AppendToMessage(ByVal strLine As String)
        If String.IsNullOrEmpty(m_strMessage) Then
            m_strMessage = strLine
        Else
            m_strMessage &= vbCrLf & strLine
        End If
    End Sub

    Public Sub AddAttachment(ByVal strFile As String)
        If m_arrFiles Is Nothing Then
            m_arrFiles = New ArrayList
        End If

        m_arrFiles.Add(strFile)
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                If m_strRootFolder IsNot Nothing Then
                    Try
                        IO.Directory.Delete(m_strRootFolder, True)
                    Catch ex As Exception
                    End Try
                End If

                If Not m_arrFiles Is Nothing Then
                    m_arrFiles.Clear()
                    m_arrFiles = Nothing
                End If
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
