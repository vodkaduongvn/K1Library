Imports Microsoft.Identity.Client
Imports System.Collections
Imports System.Collections.Generic
Imports System.Security
Imports System.Threading.Tasks

<Flags>
Public Enum enumIdentityProviderState
    Configured = 1
    ApplicationCreated = 2
    CacheEnabled = 4
End Enum
Public Class clsIdentityProvider
    Implements IDisposable

    Private m_ClientApplication As IPublicClientApplication
    Private m_ConfidentialApplication As IConfidentialClientApplication
    Private m_ClientId As String
    Private m_Tenant As String
    Private m_AppRedirectUri As String
    Private m_WebRedirectUri As String
    Private m_Authority As String
    Private m_LoginPrompt As Prompt
    Private m_UseAutoLogin As Boolean
    Private m_Account As IAccount
    Private m_AutoLoginProduct As Boolean = {"RecFind 6", "Button"}.Contains(My.Application.Info.ProductName)
    Private Locker As New Object()
    Private State As enumIdentityProviderState
    Private ReadOnly TokenPath As String = $"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\Knowledgeone Corporation\{My.Application.Info.AssemblyName}\"
    Private ReadOnly ConsentRedirect As String = "https://login.microsoftonline.com/{0}/oauth2/v2.0/authorize?client_id={1}&response_type=code&scope=user.read&redirect_uri={2}&login_hint={3}"

    Public ReadOnly Property CurrentUserId As String
        Get
            Return If(m_Account.Username, Nothing)
        End Get
    End Property

    Public ReadOnly Property HasAccount As Boolean
        Get
            Return m_Account IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property Authority As String
        Get
            Return m_Authority
        End Get
    End Property

    Public ReadOnly Property AppRedirectUri As String
        Get
            Return m_AppRedirectUri
        End Get
    End Property

    Public ReadOnly Property Tenant As String
        Get
            Return m_Tenant
        End Get
    End Property

    Private ReadOnly Property TokenFilePath As String
        Get
            Return TokenPath + "token"
        End Get
    End Property

    Private ReadOnly Property Scopes As String()
        Get
            Return New List(Of String) From {
                "user.read"
            }.ToArray()
        End Get
    End Property

    Public ReadOnly Property Configured As Boolean
        Get
            Return State.HasFlag(enumIdentityProviderState.Configured)
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Function Configure(objConfiguration As clsK1Configuration) As clsIdentityProvider
        m_ClientId = objConfiguration.ActiveDirectoryClientId
        m_Tenant = objConfiguration.ActiveDirectoryTenant
        m_AppRedirectUri = objConfiguration.ActiveDirectoryAppRedirectUri
        m_WebRedirectUri = objConfiguration.ActiveDirectoryWebRedirectUri
        m_Authority = objConfiguration.ActiveDirectoryAuthority
        m_LoginPrompt = If(objConfiguration.UseAutomaticLogins, Prompt.SelectAccount, Prompt.ForceLogin)
        m_UseAutoLogin = objConfiguration.UseAutomaticLogins
        State = State Or enumIdentityProviderState.Configured
        Return Me
    End Function

    'TODO - Remove and update references    
    Public Function Configure(strClientId As String, strTenantId As String, Optional strRedirectUri As String = "", Optional blnPrompt As Boolean = False) As clsIdentityProvider
        m_ClientId = strClientId
        m_Tenant = strTenantId
        m_AppRedirectUri = strRedirectUri
        m_WebRedirectUri = String.Empty
        m_LoginPrompt = If(blnPrompt, Prompt.ForceLogin, Prompt.NoPrompt)
        m_UseAutoLogin = False
        State = State Or enumIdentityProviderState.Configured
        Return Me
    End Function

    Public Function Configure(strClientId As String, strTenantId As String, strRedirectUri As String, strAuthority As String, blnPrompt As Boolean) As clsIdentityProvider
        m_ClientId = strClientId
        m_Tenant = strTenantId
        m_AppRedirectUri = strRedirectUri
        m_WebRedirectUri = String.Empty
        m_Authority = strAuthority
        m_LoginPrompt = If(blnPrompt, Prompt.ForceLogin, Prompt.NoPrompt)
        m_UseAutoLogin = False
        State = State Or enumIdentityProviderState.Configured
        Return Me
    End Function

    Public Function CreatePublicApplication(Optional blnWithRedirect As Boolean = True) As clsIdentityProvider
        Dim objBuilder = PublicClientApplicationBuilder.Create(m_ClientId) _
            .WithAuthority(AzureCloudInstance.AzurePublic, m_Tenant)

        If blnWithRedirect Then
            objBuilder.WithRedirectUri(m_AppRedirectUri)
        End If

        m_ClientApplication = objBuilder.Build()

        State = State Or enumIdentityProviderState.ApplicationCreated
        Return Me
    End Function

    Public Function CreateConfidentialApplication(strSecret As String) As clsIdentityProvider
        m_ConfidentialApplication = ConfidentialClientApplicationBuilder.Create(m_ClientId) _
            .WithClientSecret(strSecret) _
            .WithAuthority(AzureCloudInstance.AzurePublic, m_Tenant) _
            .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient") _
            .Build()
        Return Me
    End Function

    Public Function EnableTokenCache() As clsIdentityProvider
        If m_UseAutoLogin AndAlso m_AutoLoginProduct Then
            m_ClientApplication.UserTokenCache.SetBeforeAccess(AddressOf GetToken)
            m_ClientApplication.UserTokenCache.SetAfterAccess(AddressOf SaveToken)
            State = State Or enumIdentityProviderState.CacheEnabled
        Else
            Try
                If File.Exists(TokenFilePath) Then
                    File.Delete(TokenFilePath)
                End If
            Catch
                ' Swallow
            End Try
        End If
        Return Me
    End Function

    Public Sub GetToken(e As TokenCacheNotificationArgs)
        Try
            SyncLock Locker
                If File.Exists(TokenFilePath) Then
                    Dim f = File.ReadAllBytes(TokenFilePath)
                    e.TokenCache.DeserializeMsalV3(f)
                End If
            End SyncLock
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SaveToken(e As TokenCacheNotificationArgs)
        If e.HasStateChanged Then
            Try
                SyncLock Locker
                    If Not Directory.Exists(TokenPath) Then
                        Directory.CreateDirectory(TokenPath)
                    End If
                    File.WriteAllBytes(TokenFilePath, e.TokenCache.SerializeMsalV3())
                End SyncLock
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Async Function GetAccountAsync() As Task(Of IAccount)
        If m_Account Is Nothing Then
            Dim accounts = Await m_ClientApplication.GetAccountsAsync()
            m_Account = accounts.FirstOrDefault()
        End If
        Return m_Account
    End Function

    Public Function GetAccount() As IAccount
        If m_Account Is Nothing Then
            Dim accounts = m_ClientApplication.GetAccountsAsync() _
                .GetAwaiter() _
                .GetResult()
            m_Account = accounts.FirstOrDefault()
        End If
        Return m_Account
    End Function

    Private Async Function GetAccountAsync(objProfile As clsUserProfile) As Task(Of IAccount)
        If m_Account Is Nothing Then
            Dim accounts = Await m_ClientApplication.GetAccountsAsync()
            m_Account = accounts.FirstOrDefault(Function(account)
                                                    Return account.Username = objProfile.UserID
                                                End Function)
        End If
        Return m_Account
    End Function

    Private Function GetAccount(objProfile As clsUserProfile) As IAccount
        If m_Account Is Nothing Then
            Dim accounts = m_ClientApplication.GetAccountsAsync() _
                .GetAwaiter() _
                .GetResult()
            m_Account = accounts.FirstOrDefault(Function(account)
                                                    Return account.Username = objProfile.UserID
                                                End Function)
        End If
        Return m_Account
    End Function

    Private Async Function AcquireTokenSilentAsync() As Task(Of AuthenticationResult)
        Dim account = Await GetAccountAsync()
        Try
            Return Await m_ClientApplication.AcquireTokenSilent(Scopes, account) _
                .ExecuteAsync()
        Catch ex As MsalUiRequiredException
            Return Nothing
        Catch
            Throw
        End Try
    End Function

    Private Function AcquireTokenSilent() As AuthenticationResult
        Dim account = GetAccount()
        Try
            Return m_ClientApplication.AcquireTokenSilent(Scopes, account) _
                .ExecuteAsync() _
                .GetAwaiter() _
                .GetResult()
        Catch ex As MsalUiRequiredException
            Return Nothing
        Catch
            Throw
        End Try
    End Function

    Private Async Function AcquireTokenInteractiveAsync() As Task(Of AuthenticationResult)
        Dim account = Await GetAccountAsync()

        Try
            Return Await m_ClientApplication.AcquireTokenInteractive(Scopes) _
                .WithAccount(account) _
                .WithPrompt(m_LoginPrompt) _
                .ExecuteAsync()
        Catch
            Throw
        End Try
    End Function

    Private Function AcquireTokenInteractive() As AuthenticationResult
        Dim account = GetAccount()

        Try
            Return m_ClientApplication.AcquireTokenInteractive(Scopes) _
                .WithAccount(account) _
                .WithPrompt(m_LoginPrompt) _
                .ExecuteAsync() _
                .GetAwaiter() _
                .GetResult()
        Catch
            Throw
        End Try
    End Function

    Private Async Function AcquireTokenByUsernamePasswordAsync(username As String, password As Security.SecureString) As Task(Of AuthenticationResult)
        Try
            Return Await m_ClientApplication.AcquireTokenByUsernamePassword(Scopes, username, password) _
                .ExecuteAsync()
        Catch
            Throw
        End Try
    End Function

    Private Function AcquireTokenByUsernamePassword(username As String, password As Security.SecureString) As AuthenticationResult
        Try
            Return m_ClientApplication.AcquireTokenByUsernamePassword(Scopes, username, password).ExecuteAsync() _
                .GetAwaiter() _
                .GetResult()
        Catch
            Throw
        End Try
    End Function

    Public Async Function LoginAsync(objDB As clsDB) As Task(Of clsUserProfile)
        Dim m_Identity = Await AcquireTokenSilentAsync()
        If m_Identity Is Nothing Then
            m_Identity = Await AcquireTokenInteractiveAsync()
        End If

        Await GetAccountAsync()

        Return GetUserProfile(objDB, m_Identity.Account)
    End Function

    Public Function Login(objDB As clsDB) As clsUserProfile
        Dim m_Identity = AcquireTokenSilent()
        If m_Identity Is Nothing Then
            m_Identity = AcquireTokenInteractive()
        End If

        Dim account = GetAccount()

        Return GetUserProfile(objDB, m_Identity.Account)
    End Function

    Public Async Function LoginAsync(objDB As clsDB, strUsername As String, strPassword As SecureString, Optional blnWithUI As Boolean = False) As Task(Of clsUserProfile)
        Dim objAuthResult As AuthenticationResult
        Dim blnExceptionCaught As Boolean = False
        Try
            objAuthResult = Await AcquireTokenByUsernamePasswordAsync(strUsername, strPassword)
        Catch ex As MsalUiRequiredException When ex.Classification = UiRequiredExceptionClassification.ConsentRequired
            blnExceptionCaught = True
            If Not blnWithUI Then
                Dim objUri As New Uri(String.Format(ConsentRedirect, m_Tenant, m_ClientId, m_AppRedirectUri, strUsername))
                Throw New clsK1MsalAuthorizationRequestException(objUri.ToString())
            End If
        Catch ex As MsalUiRequiredException
            Throw
        End Try

        If blnExceptionCaught AndAlso blnWithUI Then
            objAuthResult = Await AcquireTokenInteractiveAsync()
        End If

        Dim objAccount = Await GetAccountAsync()

        Return GetUserProfile(objDB, objAccount)
    End Function

    Public Function Login(objDB As clsDB, strUsername As String, strPassword As SecureString, Optional blnWithUI As Boolean = False) As clsUserProfile
        Dim objAuthResult As AuthenticationResult

        Try
            objAuthResult = AcquireTokenByUsernamePassword(strUsername, strPassword)
        Catch ex As MsalUiRequiredException When ex.Classification = UiRequiredExceptionClassification.ConsentRequired
            If blnWithUI Then
                objAuthResult = AcquireTokenInteractive()
            Else
                Dim objUri As New Uri(String.Format(ConsentRedirect, m_Tenant, m_ClientId, m_AppRedirectUri, strUsername))
                Throw New clsK1MsalAuthorizationRequestException(objUri.ToString())
            End If
        Catch ex As MsalUiRequiredException
            Throw
        End Try

        Dim objAccount = GetAccount()

        Return GetUserProfile(objDB, objAccount)
    End Function

    Private Function GetUserProfile(objDB As clsDB, objIdentity As IAccount) As clsUserProfile
        Return clsUserProfile.GetUserProfile(objIdentity, objDB)
    End Function

    Public Function GetUserProfile(objDB As clsDB) As clsUserProfile
        Dim objIdentity = GetAccount()
        Return clsUserProfile.GetUserProfile(objIdentity, objDB)
    End Function

    Public Async Function GetUserProfileAsync(objDB As clsDB) As Task(Of clsUserProfile)
        Dim objIdentity = Await GetAccountAsync()
        Return clsUserProfile.GetUserProfile(objIdentity, objDB)
    End Function

    Public Sub Logout()
        Dim account = GetAccount()
        If account IsNot Nothing Then
            Try
                m_ClientApplication.RemoveAsync(account) _
                    .GetAwaiter() _
                    .GetResult()

                If File.Exists(TokenFilePath) Then
                    File.Delete(TokenFilePath)
                End If
            Catch
                ' Swallow
            Finally
                m_Account = Nothing
            End Try
        End If
    End Sub

    Public Async Function LogoutAsync() As Task
        Dim account = Await GetAccountAsync()
        If account IsNot Nothing Then
            Try
                Await m_ClientApplication.RemoveAsync(account)
                If File.Exists(TokenFilePath) Then
                    File.Delete(TokenFilePath)
                End If
            Catch
                ' Swallow
            Finally
                m_Account = Nothing
            End Try
        End If
    End Function

    Public Sub Logout(objProfile As clsUserProfile)
        Dim account = GetAccount(objProfile)
        If account IsNot Nothing Then
            Try
                m_ClientApplication.RemoveAsync(account) _
                    .GetAwaiter() _
                    .GetResult()

                If File.Exists(TokenFilePath) Then
                    File.Delete(TokenFilePath)
                End If
            Catch
                ' Swallow
            Finally
                m_Account = Nothing
            End Try
        End If
    End Sub

    Public Async Function LogoutAsync(objProfile As clsUserProfile) As Task
        Dim account = Await GetAccountAsync(objProfile)
        If account IsNot Nothing Then
            Try
                Await m_ClientApplication.RemoveAsync(account)
                If File.Exists(TokenFilePath) Then
                    File.Delete(TokenFilePath)
                End If
            Catch
                ' Swallow
            Finally
                m_Account = Nothing
            End Try
        End If
    End Function

    Public Shared Function AzureActiveDirectoryEnabled(objConfiguration As clsK1Configuration) As Boolean
        Return objConfiguration.PasswordType = clsDBConstants.enumPasswordType.AZURE_ACTIVE_DIRECTORY
    End Function

#Region "IDisposable"
    Private disposedValue As Boolean

    Private Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

Public Class clsK1MsalAuthorizationRequestException
    Inherits clsK1Exception

    Public Sub New(strMessage As String)
        MyBase.New(strMessage)
    End Sub
End Class
