Imports System.ServiceModel.Channels
Imports K1Library.K1ConnectionManager
Imports System.ServiceModel

Public Class clsConnectionManager
    Implements IDisposable

#Region " Members "

    Private m_objProxy As K1ConnectionManager.ConnectionManagerClient
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
    Private m_strServerName As String
    Private m_strCertDistinguishedName As String
    Private m_blnCreateBinding As Boolean

#End Region

#Region " Constructors "

    Public Sub New(ByVal strServer As String, ByVal strCertDistinguishedName As String)
        Me.New(strServer, strCertDistinguishedName, False)
    End Sub

    Public Sub New(ByVal strServer As String, ByVal strCertDistinguishedName As String, _
                   ByVal blnCreateBinding As Boolean)
        Try
            m_strServerName = strServer
            m_strCertDistinguishedName = strCertDistinguishedName
            m_blnCreateBinding = blnCreateBinding

            InitialiseProxy(blnCreateBinding)
        Catch ex As ServiceModel.EndpointNotFoundException
            Throw New Exception("The connection server is not currently available. Please contact your administrator.")
        Catch ex As Exception
            Throw New Exception("Unable to communicate with the server. Please contact your administrator.")
        End Try
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property State() As ServiceModel.CommunicationState
        Get
            If m_objProxy IsNot Nothing Then
                Return m_objProxy.State
            Else
                Return ServiceModel.CommunicationState.Closed
            End If
        End Get
    End Property

    Public ReadOnly Property Proxy() As K1ConnectionManager.ConnectionManagerClient
        Get
            If Not m_objProxy.State = ServiceModel.CommunicationState.Opened Then
                InitialiseProxy(m_blnCreateBinding)
            End If

            Return m_objProxy
        End Get
    End Property

    Public Property OpenTimeout() As TimeSpan
        Get
            Return m_objProxy.Endpoint.Binding.OpenTimeout
        End Get
        Set(ByVal value As TimeSpan)
            m_objProxy.Endpoint.Binding.OpenTimeout = value
        End Set
    End Property

    Public ReadOnly Property ServerName() As String
        Get
            Return m_strServerName
        End Get
    End Property

    Public ReadOnly Property CertificateDistinguishedName() As String
        Get
            Return m_strCertDistinguishedName
        End Get
    End Property

    Public ReadOnly Property IsOpen() As Boolean
        Get
            If Me.State = ServiceModel.CommunicationState.Opened _
            OrElse Me.State = ServiceModel.CommunicationState.Opening _
            OrElse Me.State = ServiceModel.CommunicationState.Created Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

#End Region

#Region " Methods "

    Private Sub InitialiseProxy(ByVal blnCreateBinding As Boolean)

        Try
            Dim objEndPoint As New EndpointAddress(New Uri(String.Format("net.tcp://{0}/K1/ConnectionManagerService/", m_strServerName)),
                                                   New DnsEndpointIdentity(m_strCertDistinguishedName))

            If blnCreateBinding Then
                Dim objBinding As New NetTcpBinding()
                objBinding.Name = "NetTcpBinding_IConnectionManager"
                objBinding.Security.Mode = SecurityMode.Message
                'objBinding.Security.Message.AlgorithmSuite = Security.SecurityAlgorithmSuite.Default  
                objBinding.Security.Message.AlgorithmSuite = Security.SecurityAlgorithmSuite.TripleDesSha256
                objBinding.Security.Message.ClientCredentialType = MessageCredentialType.UserName

                m_objProxy = New ConnectionManagerClient(objBinding, objEndPoint)

                m_objProxy.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = Security.X509CertificateValidationMode.Custom
                m_objProxy.ClientCredentials.ServiceCertificate.Authentication.CustomCertificateValidator = New CustomX509Validator

            Else
                m_objProxy = New ConnectionManagerClient("NetTcpBinding_IConnectionManager", objEndPoint)
            End If

            m_objProxy.ClientCredentials.UserName.UserName = "ConnectionManagerAdmin"
            m_objProxy.ClientCredentials.UserName.Password = "ConnectionManagerAdminPassword"
        Catch ex As EndpointNotFoundException
            Throw New ApplicationException("The connection server is not currently available. Please contact your administrator.")
        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' Checks if the specified system is managed by the connection manager
    ''' </summary>
    ''' <param name="strSystemName">Name of system that we want the information for.</param>
    ''' <returns>Returns True if system is managed; otherwise False.</returns>
    Public Function SystemExists(ByVal strSystemName As String) As Boolean
        Return Proxy.SystemExists(strSystemName)
    End Function

    ''' <summary>
    ''' Returns the names for all the managed systems of the specified applications.
    ''' </summary>
    ''' <param name="eAppType">Application that relates to the managed system.</param>
    ''' <returns>String array of system names.</returns>
    Public Function GetSystemNames(Optional ByVal eAppType As K1ConnectionManager.IConnectionManagerSystemType = K1ConnectionManager.IConnectionManagerSystemType.All, _
                                   Optional ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType = K1ConnectionManager.IConnectionManagerAccessType.All) As String()
        Return Proxy.GetSystemNames(eAppType, eAccessType)
    End Function

    ''' <summary>
    ''' Returns information about the specified system.
    ''' </summary>
    ''' <param name="strSystemName">Name of system that we want the information for.</param>
    Public Function GetSystemInfo(ByVal strSystemName As String, Optional ByVal dblVersion As Double = 0, _
    Optional ByVal strAppName As String = Nothing, Optional ByVal blnSkipCheckVersion As Boolean = False, _
    Optional ByVal dblMinDBVersion As Double = 0) As SystemInfo
        Dim blnIsLatest As Boolean
        Dim arrInfo() As String = Proxy.GetSystemInfo(strSystemName, blnIsLatest, dblVersion, strAppName, blnSkipCheckVersion, dblMinDBVersion)

        If arrInfo Is Nothing Then
            Return Nothing
        ElseIf Not blnIsLatest AndAlso Not blnSkipCheckVersion Then
            Throw New clsK1Exception(ErrorNumber.Old_Version, _
                "Could not connect to the database. The database version is out of date. " & _
                "If you wish to connect to an old database please install " & strAppName & " version 2.2.")
        End If

        If arrInfo(2) Is Nothing Then
            arrInfo(2) = K1ConnectionManager.IConnectionManagerAccessType.Direct.ToString()
        End If

        Return New SystemInfo(strSystemName,
                              CType([Enum].Parse(GetType(K1ConnectionManager.IConnectionManagerSystemType), arrInfo(1)), IConnectionManagerSystemType),
                              CType([Enum].Parse(GetType(K1ConnectionManager.IConnectionManagerAccessType), arrInfo(2)), IConnectionManagerAccessType),
                              arrInfo(0))
    End Function

    ''' <summary>
    ''' Returns the encrypted connection string of the specified system.
    ''' </summary>
    ''' <param name="strSystemName">Name of system that we want the connection string for.</param>
    ''' <returns>Encrypted connection string.</returns>
    Public Function GetEncryptedConnectionString(ByVal dblMinDBVersion As Double, _
    Optional ByVal strSystemName As String = Nothing, _
    Optional ByVal blnSkipCheckDatabaseVersion As Boolean = False) As String
        Dim strConnection As String = Nothing
        Dim strSysNames As String() = GetSystemNames()

        If strSysNames Is Nothing OrElse strSysNames.Count = 0 Then
            Return strConnection
        End If

        Dim strSystem As String = strSystemName
        If strSystem Is Nothing OrElse strSysNames.Count = 1 Then
            strSystem = strSysNames(0)
        End If

        Dim blnIsLatest As Boolean
        Dim arrSysInfo As String() = Proxy.GetSystemInfo(strSystem, blnIsLatest, _
            modGlobal.GetVersion(), My.Application.Info.ProductName, blnSkipCheckDatabaseVersion, dblMinDBVersion)
        If arrSysInfo Is Nothing OrElse arrSysInfo.GetUpperBound(0) = 0 Then
            Throw New clsK1Exception("The system name '" & strSystem & "' does not exist. " & _
                                     "Please seek assistance from your administrator.")
        ElseIf Not blnIsLatest AndAlso Not blnSkipCheckDatabaseVersion Then
            Throw New clsK1Exception(ErrorNumber.Old_Version, _
                                     "Could not connect to the database. The database version is out of date. " & _
                                     "If you wish to connect to an old database please install " & My.Application.Info.ProductName & " version 2.2.")
        End If

        Return arrSysInfo(0)
    End Function

    Public Sub UpdateConnectionString(ByVal strSystem As String, ByVal strConnection As String)
        Proxy.UpdateConnectionString(strSystem, strConnection)
    End Sub

    Public Sub UpdateConnection(ByVal strSystem As String, ByVal strConnection As String, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType)
        Proxy.UpdateConnection(strSystem, strConnection, eAccessType)
    End Sub

    Public Sub InsertConnectionString(ByVal strSystem As String, ByVal strConnection As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType)
        Proxy.InsertConnectionString(strSystem, strConnection, eSystemType)
    End Sub

    Public Sub InsertConnection(ByVal strSystem As String, ByVal strConnection As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType)
        Proxy.InsertConnection(strSystem, strConnection, eSystemType, eAccessType)
    End Sub

    Public Sub DeleteConnection(ByVal strSystem As String)
        Proxy.DeleteSystem(strSystem)
    End Sub

#End Region

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)

        If Not Me.m_blnDisposedValue Then
            If blnDisposing Then
                Dim communicationObject = TryCast(Proxy, ICommunicationObject)
                If (communicationObject IsNot Nothing) Then
                    If (communicationObject.State = CommunicationState.Faulted) Then
                        communicationObject.Abort()
                    Else
                        communicationObject.Close()
                    End If
                Else
                    m_objProxy.Close()
                End If
                m_objProxy = Nothing
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

    Public Class SystemInfo

        Private m_strName As String
        Private m_eType As K1ConnectionManager.IConnectionManagerSystemType
        Private m_eAccessType As K1ConnectionManager.IConnectionManagerAccessType
        Private m_strValue As String

        Public Sub New(ByVal strName As String, _
                       ByVal eType As K1ConnectionManager.IConnectionManagerSystemType, _
                       ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType, _
                       ByVal strValue As String)
            m_strName = strName
            m_eType = eType
            m_eAccessType = eAccessType
            m_strValue = strValue
        End Sub

        Public ReadOnly Property Name() As String
            Get
                Return m_strName
            End Get
        End Property

        Public ReadOnly Property ApplicationType() As K1ConnectionManager.IConnectionManagerSystemType
            Get
                Return m_eType
            End Get
        End Property

        Public ReadOnly Property AccessType() As K1ConnectionManager.IConnectionManagerAccessType
            Get
                Return m_eAccessType
            End Get
        End Property

        Public ReadOnly Property value() As String
            Get
                Return m_strValue
            End Get
        End Property
    End Class
End Class
