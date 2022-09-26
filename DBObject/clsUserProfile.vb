''' <summary>
''' This class represents the table Profile in the Database.
''' </summary>
''' 
Public NotInheritable Class clsUserProfile
    Inherits clsUserProfileBase

#Region " Members "
    Private m_colDRMMethods As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colDRMFunctions As FrameworkCollections.K1Dictionary(Of Object)

    Private m_colSecurityGroups As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colAppMethods As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colFields As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colReadOnlyFields As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colTables As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colSecurities As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colMethods As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colTypeFields As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
    Private m_colTypeMethods As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
    Private m_colTypesByTableMethod As FrameworkCollections.K1Dictionary(Of Hashtable)
    Private m_intNominatedSecurityID As Integer
    Private m_objNominatedSecurity As clsSecurity

    Private m_intSecurityGroupID As Integer
    Private m_blnIsSuperAdministrator As Boolean
    Private m_blnIsAdministrator As Boolean
    Private m_blnIsGEMAdmin As Boolean
    Private m_blnIsCaptureAdmin As Boolean
    Private m_blnIsPublic As Boolean
    Private m_strUserID As String
    Private m_strPassword As String
    Private m_objConfig As clsConfig
    Private m_blnSharePointUser As Boolean
    '2017-08-30 -- Peter Melisi -- Changes for Timezones for User Profiles
    Private m_strTimezone As String
    Private m_objTimeDifference As Double = 0
    Private m_ServerTimeZoneInfo As TimeZoneInfo
    Private m_ClientTimeZoneInfo As TimeZoneInfo

    '2021-11-18 -- Emmanuel Cardakaris -- changes for Power BI Dashboard
    Private m_BIReportId As String
    Private m_ConnectorUserKey As String

    Private Const STRDP As String = "K12cdcfjjdfj67364" & "clsUserProfile" & "Shared"
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_intSecurityGroupID = clsDBConstants.cintNULL
        m_intNominatedSecurityID = clsDBConstants.cintNULL
        m_blnIsAdministrator = False
        m_blnIsSuperAdministrator = False
        m_blnIsGEMAdmin = False
        m_blnIsCaptureAdmin = False
        m_blnIsPublic = False
        m_strTimezone = clsDBConstants.cstrNULL
        m_BIReportId = clsDBConstants.cstrNULL
        m_ConnectorUserKey = clsDBConstants.cstrNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intSecurityGroupID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cSECURITYGROUPID, clsDBConstants.cintNULL), Integer)
        m_intNominatedSecurityID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cNOMINATEDSECURITYID, objDB.SysInfo.K1Configuration.DRMDefaultSecurityID), Integer)
        m_blnIsAdministrator = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cISADMINISTRATOR, False), Boolean)
        m_blnIsSuperAdministrator = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cISSUPERADMINISTRATOR, False), Boolean)
        m_blnIsGEMAdmin = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cISGEMADMIN, False), Boolean)
        m_blnIsCaptureAdmin = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cISCAPTUREADMIN, False), Boolean)
        m_blnIsPublic = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cISPUBLIC, False), Boolean)
        m_strUserID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cUSERID, clsDBConstants.cstrNULL), String)
        m_strPassword = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cPASSWORD, clsDBConstants.cstrNULL), String)
        m_strTimezone = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cTIMEZONE, clsDBConstants.cstrNULL), String)
        '2017-08-30 -- Peter Melisi -- Changes for Timezones for User Profiles
        m_objTimeDifference = objDB.GetServerTimeDifference(m_strTimezone)
        m_ClientTimeZoneInfo = TimeZoneInfo.Local

        '2021-11-18 -- Emmanuel Cardakaris -- Changes for Dashboard
        m_BIReportId = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cBIREPORTID, clsDBConstants.cstrNULL), String)
        m_ConnectorUserKey = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cCONNECTORUSERKEY, clsDBConstants.cstrNULL), String)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property SecurityGroupID() As Integer
        Get
            Return m_intSecurityGroupID
        End Get
    End Property

    Public ReadOnly Property SecurityGroup() As clsSecurityGroup
        Get
            If Not m_intSecurityGroupID = clsDBConstants.cintNULL Then
                Return m_objDB.SysInfo.SecurityGroups(CStr(m_intSecurityGroupID))
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SecurityGroups() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            Return m_colSecurityGroups
        End Get
    End Property

    Public ReadOnly Property LinkDRMFunctions() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colDRMFunctions Is Nothing Then
                m_colDRMFunctions = LoadUPHashTable(clsDBConstants.Tables.cLINKUSERPROFILEDRMFUNCTION,
                                                clsDBConstants.Fields.LinkUserProfileDRMFunction.CDRMFUNCTIONID)
            End If
            Return m_colDRMFunctions
        End Get
    End Property

    Public ReadOnly Property LinkDRMmethods() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colDRMMethods Is Nothing Then
                m_colDRMMethods = LoadUPHashTable(clsDBConstants.Tables.cLINKUSERPROFILEDRMMETHOD,
                                                clsDBConstants.Fields.LinkUserProfileDRMMethod.cDRMMETHODID)
            End If
            Return m_colDRMMethods
        End Get
    End Property

    'Added for Multiple Security Groups
    Public ReadOnly Property LinkSecurityGroups() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colSecurityGroups Is Nothing Then
                m_colSecurityGroups = LoadUPHashTable(clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP,
                    clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID)
            End If
            Return m_colSecurityGroups
        End Get
    End Property

    Public ReadOnly Property LinkApplicationMethods() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colAppMethods Is Nothing Then
                m_colAppMethods = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPAPPMETHOD,
                    clsDBConstants.Fields.LinkSecurityGroupAppMethod.cAPPMETHODID)
            End If
            Return m_colAppMethods
        End Get
    End Property

    Public ReadOnly Property LinkFields() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colFields Is Nothing Then
                m_colFields = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPFIELD,
                    clsDBConstants.Fields.LinkSecurityGroupField.cFIELDID)
            End If
            Return m_colFields
        End Get
    End Property

    Public ReadOnly Property LinkReadOnlyFields() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colReadOnlyFields Is Nothing Then
                m_colReadOnlyFields = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPREADONLYFIELDS,
                    clsDBConstants.Fields.LinkSecurityGroupReadOnlyField.cFIELDID)
            End If
            Return m_colReadOnlyFields
        End Get
    End Property

    Public ReadOnly Property LinkSecurities() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colSecurities Is Nothing Then
                m_colSecurities = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPSECURITY,
                    clsDBConstants.Fields.cSECURITYID)
            End If
            Return m_colSecurities
        End Get
    End Property

    Public ReadOnly Property LinkTables() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colTables Is Nothing Then
                m_colTables = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPTABLE,
                    clsDBConstants.Fields.LinkSecurityGroupTable.cTABLEID)
            End If
            Return m_colTables
        End Get
    End Property

    Public ReadOnly Property LinkMethods() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colMethods Is Nothing Then
                m_colMethods = LoadSGHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD,
                    clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID)
            End If
            Return m_colMethods
        End Get
    End Property

    Public ReadOnly Property TypeFields() As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Get
            If m_colTypeFields Is Nothing Then
                m_colTypeFields = clsAccessRightField.GetList(m_objDB, Me)
            End If
            Return m_colTypeFields
        End Get
    End Property

    Public ReadOnly Property TypeMethods() As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Get
            If m_colTypeMethods Is Nothing Then
                m_colTypeMethods = clsAccessRightMethod.GetList(m_objDB, Me)
            End If
            Return m_colTypeMethods
        End Get
    End Property

    Public ReadOnly Property NominatedSecurityID() As Integer
        Get
            Return m_intNominatedSecurityID
        End Get
    End Property

    Public ReadOnly Property NominatedSecurity() As clsSecurity
        Get
            If m_objNominatedSecurity Is Nothing Then
                m_objNominatedSecurity = clsSecurity.GetItem(m_intNominatedSecurityID, m_objDB)
            End If
            Return m_objNominatedSecurity
        End Get
    End Property

    Public ReadOnly Property IsAdministrator() As Boolean
        Get
            Return m_blnIsAdministrator
        End Get
    End Property

    Public ReadOnly Property IsSuperAdministrator() As Boolean
        Get
            Return m_blnIsSuperAdministrator
        End Get
    End Property

    Public ReadOnly Property IsGEMAdministrator() As Boolean
        Get
            Return m_blnIsGEMAdmin
        End Get
    End Property

    Public ReadOnly Property IsCaptureAdministrator() As Boolean
        Get
            Return m_blnIsCaptureAdmin
        End Get
    End Property

    Public ReadOnly Property IsPublic() As Boolean
        Get
            Return m_blnIsPublic
        End Get
    End Property

    Public Property IsSharePointUser() As Boolean
        Get
            Return m_blnSharePointUser
        End Get
        Set(ByVal value As Boolean)
            m_blnSharePointUser = value
        End Set
    End Property

    Public Property Config() As clsConfig
        Get
            If m_objConfig Is Nothing Then
                m_objConfig = New clsConfig
            End If
            Return m_objConfig
        End Get
        Set(ByVal value As clsConfig)
            m_objConfig = value
        End Set
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return m_strUserID
        End Get
    End Property

    Public ReadOnly Property Password() As Integer
        Get
            Return CInt(m_strPassword)
        End Get
    End Property

    Public ReadOnly Property Timezone() As String
        Get
            Return m_strTimezone
        End Get
    End Property

    '2017-08-30 -- Peter Melisi -- Changes for Timezones for User Profiles
    Public Property TimeDifference() As Double
        Get
            Return m_objTimeDifference
        End Get
        Set(value As Double)
            m_objTimeDifference = GetServerTimeDifference(value)
        End Set
    End Property

    Public ReadOnly Property TypesByTableMethod() As FrameworkCollections.K1Dictionary(Of Hashtable)
        Get
            If m_colTypesByTableMethod Is Nothing Then
                AssignTypesByTableMethod()
            End If
            Return m_colTypesByTableMethod
        End Get
    End Property

    Public Function ServerTimeZone(objDB As clsDB) As TimeZoneInfo
        If m_ServerTimeZoneInfo Is Nothing Then
            m_ServerTimeZoneInfo = objDB.GetServerTimeZone
        End If
        Return m_ServerTimeZoneInfo
    End Function

    Public Property ClientTimeZone() As TimeZoneInfo
        Get
            Return m_ClientTimeZoneInfo
        End Get
        Set(value As TimeZoneInfo)
            m_ClientTimeZoneInfo = value
        End Set
    End Property

    Public Property BIReportId() As String
        Get
            Return m_BIReportId
        End Get
        Set(value As String)
            m_BIReportId = value
        End Set
    End Property

    Public Property ConnectorUserKey() As String
        Get
            Return m_ConnectorUserKey
        End Get
        Set(value As String)
            m_ConnectorUserKey = value
        End Set
    End Property
#End Region
    Public Function GetServerTimeDifference(minutesDifference As Double) As Double 'localtimedate
        Try

            Dim dblNow As Double

            Dim objDT As DataTable = m_objDB.GetDataTableBySQL("Select DATEDIFF(minute, GETUTCDATE(), GETDATE())")
            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                dblNow = CDbl(objDT.Rows(0)(0))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            'Dim currentOffset As TimeSpan = Timezone.CurrentTimeZone.GetUtcOffset(Now)

            dblNow = (dblNow - minutesDifference) / 60 ' currentOffset.TotalHours

            Return dblNow
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.ToString)
            Throw
        End Try
    End Function

    ' Ara Melkonian - 2100003652
    ' Convert web server time to database server time.
    Public Function ServerToLocalTime(objDatetime As Date) As Date
        Dim localTimeZone = TimeZoneInfo.Local
        Dim serverTz As TimeZoneInfo = ServerTimeZone(m_objDB)

        Return If(localTimeZone.Equals(serverTz), objDatetime, TimeZoneInfo.ConvertTime(objDatetime, localTimeZone, serverTz))
    End Function

    Private Function LoadUPHashTable(ByVal strTable As String, ByVal strField As String) As FrameworkCollections.K1Dictionary(Of Object)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of Object)

        Try
            Dim objDT As DataTable = m_objDB.GetDataTableBySQL("SELECT [" & strField & "] AS [X] " &
                "FROM [" & strTable & "] WHERE [UserProfileID] = " & m_intID)

            colObjects = New FrameworkCollections.K1Dictionary(Of Object)
            For Each objDR As DataRow In objDT.Rows
                colObjects.Add(CStr(objDR(0)), CInt(objDR(0)))
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function LoadSGHashTable(ByVal strTable As String, ByVal strField As String) As FrameworkCollections.K1Dictionary(Of Object)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of Object)

        Try
            Dim objDT As DataTable = m_objDB.GetDataTableBySQL("SELECT DISTINCT [" & strField & "] AS [X] " &
            "FROM [" & strTable & "] INNER JOIN [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "] " &
            "ON [" & strTable & "].[SecurityGroupID] = [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "].[" & clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID & "]" &
            "WHERE [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "].[" & clsDBConstants.Fields.LinkUserProfileSecurityGroup.cUSERPROFILEID & "] = " & m_intID)

            colObjects = New FrameworkCollections.K1Dictionary(Of Object)
            For Each objDR As DataRow In objDT.Rows
                If Not colObjects.ContainsKey(CStr(objDR(0))) Then
                    colObjects.Add(CStr(objDR(0)), CInt(objDR(0)))
                End If
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub AssignTypesByTableMethod()
        m_colTypesByTableMethod = New FrameworkCollections.K1Dictionary(Of Hashtable)
        Dim colTypeMethods As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod) = TypeMethods

        For Each objAccessMethod As clsAccessRightMethod In colTypeMethods.Values
            Dim colIDs As Hashtable

            colIDs = m_colTypesByTableMethod(CStr(objAccessMethod.TableMethodID))

            If colIDs Is Nothing Then
                colIDs = New Hashtable
                m_colTypesByTableMethod.Add(CStr(objAccessMethod.TableMethodID), colIDs)
            End If

            If colIDs(CStr(objAccessMethod.AppliesToTypeID)) Is Nothing Then
                colIDs.Add(CStr(objAccessMethod.AppliesToTypeID), objAccessMethod.AppliesToTypeID)
            End If
        Next
    End Sub

    '2016-10-06 -- Peter & James -- Bug fix for #1600003221
    Public Sub ResetProfile()
        m_colSecurityGroups = Nothing
        m_colAppMethods = Nothing
        m_colFields = Nothing
        m_colReadOnlyFields = Nothing
        m_colTables = Nothing
        m_colSecurities = Nothing
        m_colMethods = Nothing
        m_colTypeFields = Nothing
        m_colTypeMethods = Nothing
        m_colTypesByTableMethod = Nothing
        m_objNominatedSecurity = Nothing
    End Sub

#Region " Business Logic "

#Region " Get Item From Login "

    ''' <summary>
    ''' Authenticates user and fetches the user profile from the data store.
    ''' </summary>
    ''' <param name="objDB">Data access object. It can be encapsulating a direct Ado.net connection or web services.</param>
    ''' <param name="strUserName">user name in plain text</param>
    ''' <param name="strPassword">password in plain text</param>
    ''' <param name="authenticator">authenticator use this when the data access object encapsulates web services. I.E clsDB_WS has a member called AuthenticateUser</param>
    ''' <param name="enumPasswordType">nullable enum, used to determine if logging in with a recfind or an active directory account</param>
    ''' <returns>Current authenticated user profile from data store or active directory.</returns>
    ''' <remarks>Uses Active Directory or data store to authenticate user.</remarks>
    Public Shared Function GetItemFromLogin(ByVal objDB As clsDB, ByVal strUserName As String, ByVal strPassword As String,
                                            Optional authenticator As Func(Of String, String, Boolean) = Nothing,
                                            Optional enumPasswordType As clsDBConstants.enumPasswordType? = Nothing) As clsUserProfile
        ' Ara Melkonian - 2000003608
        ' Added login type determination - default is a fallback
        If Not enumPasswordType.HasValue Then
            enumPasswordType = objDB.SysInfo.K1Configuration.PasswordType
        End If

        Return GetUserProfile(objDB, strUserName, strPassword, enumPasswordType.Value, authenticator)
    End Function

    Private Shared Function GetUserProfile(ByVal objDB As clsDB, ByVal strUserName As String, ByVal strPassword As String,
                               enumPasswordType As clsDBConstants.enumPasswordType,
                               Optional authenticator As Func(Of String, String, Boolean) = Nothing) As clsUserProfile
        Dim strStoredProcedure As String
        Dim strEncPassword As String
        Dim colParams As New clsDBParameterDictionary

        If enumPasswordType = clsDBConstants.enumPasswordType.RECFIND_DATABASE Then

            Dim objEncryption As New clsEncryption(False, True)
            strEncPassword = objEncryption.Encrypt(strPassword)

            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cUSERID), strUserName))
            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cPASSWORD), strEncPassword))

        Else

            Try
                If (authenticator Is Nothing) Then

                    Using pc As New DirectoryServices.AccountManagement.PrincipalContext(DirectoryServices.AccountManagement.ContextType.Domain,
                                                                         objDB.SysInfo.K1Configuration.ActiveDirectoryName)
                        If pc.ValidateCredentials(strUserName,
                                                  strPassword,
                                                  DirectoryServices.AccountManagement.ContextOptions.Negotiate) = False Then
                            Return Nothing
                        End If
                    End Using
                Else
                    '[Naing] Bug: 1300002430 Fixed
                    If (Not authenticator(strUserName, strPassword)) Then
                        Return Nothing
                    End If
                End If
            Catch ex As Exception
                Throw New clsK1Exception(clsDBConstants.cintNULL, "Could not authenticate user through web services. Please contact your system administrator to check your system configuration.", ex)
            End Try

            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cUSERID), strUserName))

        End If

        strStoredProcedure = clsDBConstants.Tables.cUSERPROFILE &
            clsDBConstants.StoredProcedures.cGETLIST
        Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

        If Not objDT.Rows.Count = 0 Then
            Return New clsUserProfile(objDT.Rows(0), objDB)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetUserProfileWithClaims(ByVal identity As Security.Claims.ClaimsIdentity, ByVal objDB As clsDB) As clsUserProfile
        Dim name = identity?.FindFirst("preferred_username")?.Value

        If String.IsNullOrEmpty(name) Then
            Return Nothing
        End If

        Using colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cUSERID), name))
            Dim strStoredProcedure = clsDBConstants.Tables.cUSERPROFILE & clsDBConstants.StoredProcedures.cGETLIST
            Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        End Using

        Return Nothing
    End Function

    Public Shared Function GetUserProfile(ByVal identity As Microsoft.Identity.Client.IAccount, ByVal objDB As clsDB) As clsUserProfile
        Dim name = identity.Username

        If String.IsNullOrEmpty(name) Then
            Return Nothing
        End If

        Using colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cUSERID), name))
            Dim strStoredProcedure = clsDBConstants.Tables.cUSERPROFILE & clsDBConstants.StoredProcedures.cGETLIST
            Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        End Using

        Return Nothing
    End Function

    Public Overloads Shared Function GetItem(ByVal objDB As clsDB,
    ByVal intID As Integer, ByVal objAdminUser As clsUserProfile) As clsUserProfile
        Dim strStoredProcedure As String

        Try
            If Not objAdminUser.IsAdministrator Then
                Return Nothing
            End If

            Dim colParams As New clsDBParameterDictionary

            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.cID), intID))

            strStoredProcedure = clsDBConstants.Tables.cUSERPROFILE &
                clsDBConstants.StoredProcedures.cGETITEM
            Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function AutomaticLogin(ByVal objDB As clsDB,
    ByVal strUserName As String, ByVal strSystemPassword As String) As clsUserProfile
        If Not strSystemPassword = "$%RER@#FY^GWEFa32SS@2" Then
            Return Nothing
        End If

        Try
            Dim colParams As New clsDBParameterDictionary

            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.UserProfile.cUSERID), strUserName))

            Dim strStoredProcedure As String = clsDBConstants.Tables.cUSERPROFILE &
                clsDBConstants.StoredProcedures.cGETLIST
            Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetDefault(ByVal objDB As clsDB, ByVal strPassword As String) As clsUserProfile
        If Not strPassword = STRDP Then
            Throw New Exception("Access not granted to default profile")
        End If

        Try
            Dim objDT As DataTable = objDB.GetDataTableByField(clsDBConstants.Tables.cUSERPROFILE,
                clsDBConstants.Fields.cID, objDB.SysInfo.K1Configuration.DefaultProfileID)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetSharePointProfile(ByVal objDB As clsDB, ByVal strUserName As String, ByVal strPassword As String) As clsUserProfile
        If Not strPassword = STRDP Then
            clsAuditTrail.CreateUnsuccessfulLoginRecord(objDB, "Default profile - " & strUserName)
            Throw New Exception("Access not granted to default profile")
        End If

        Dim strStoredProcedure As String
        Dim colParams As New clsDBParameterDictionary

        Try
            strStoredProcedure = clsDBConstants.Tables.cUSERPROFILE &
                clsDBConstants.StoredProcedures.cGETLIST

            colParams.Add(New clsDBParameter(clsDB.ParamName(
                clsDBConstants.Fields.UserProfile.cSHAREPOINTUSER), strUserName))

            Dim objDT As DataTable = objDB.GetDataTable(strStoredProcedure, colParams)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetPublicProfile(ByVal objDB As clsDB) As clsUserProfile
        Try
            Dim objDT As DataTable = objDB.GetDataTableByField(clsDBConstants.Tables.cUSERPROFILE,
                clsDBConstants.Fields.UserProfile.cISPUBLIC, True)

            If Not objDT.Rows.Count = 0 Then
                Return New clsUserProfile(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Get Security Where Clause "

    ''' <summary>
    ''' For use with SQL or RowFilter of a dataview
    ''' </summary>
    Public Function GetSecurityWhereClause(Optional ByVal strTableName As String = Nothing) As String
        Dim strWhere As String = "("

        Dim strIDs As String = CreateIDStringFromCollection(
            Me.LinkSecurities.Values)

        If strIDs IsNot Nothing AndAlso strIDs.Length > 0 Then
            If Not String.IsNullOrEmpty(strTableName) Then
                strWhere &= "[" & strTableName & "]."
            End If
            strWhere &= "[" & clsDBConstants.Fields.cSECURITYID & "] IN (" & strIDs & ") OR "
        End If

        If Not String.IsNullOrEmpty(strTableName) Then
            strWhere &= "[" & strTableName & "]."
        End If

        strWhere &= "[" & clsDBConstants.Fields.cSECURITYID & "] IS NULL)"

        Return strWhere
    End Function
    'Public Function GetSecurityWhereClause() As String
    '    Dim strWhere As String = "("

    '    Dim strIDs As String = CreateIDStringFromCollection( _
    '        SecurityGroup.LinkSecurities.Values)

    '    If strIDs IsNot Nothing AndAlso strIDs.Length > 0 Then
    '        strWhere &= "[" & clsDBConstants.Fields.cSECURITYID & "] IN (" & strIDs & ") OR "
    '    End If

    '    strWhere &= "[" & clsDBConstants.Fields.cSECURITYID & "] IS NULL)"

    '    Return strWhere
    'End Function
#End Region

#Region " Get Settings XML "

    Public Function GetSettingsXML() As String
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cUSERPROFILE), m_intID)

        Return CType(colMasks.GetMaskValue(clsDBConstants.Fields.UserProfile.cSETTINGS), String)
    End Function
#End Region

#Region " HasAccess "

    'Pass in a securtity value and this returns whether the profile has access to it
    Public Function HasAccess(ByVal intSecurityID As Integer) As Boolean
        If Not Me.LinkSecurities(CType(intSecurityID, String)) Is Nothing Then
            Return True
        End If
        Return False
    End Function

#Region " HasTableAccess "

    Public Function HasTableAccess(ByVal strTableName As String) As Boolean
        Dim objTable As clsTable = Database.SysInfo.Tables(strTableName)
        Return HasTableAccess(objTable)
    End Function

    Public Function HasTableAccess(ByVal objTable As clsTable) As Boolean
        If (HasAccess(objTable.SecurityID) AndAlso
        Not Me.LinkTables(CType(objTable.ID, String)) Is Nothing) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " HasFieldAccess "

    Public Function HasFieldAccess(ByVal strTableName As String, ByVal strField As String,
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As Boolean
        Dim objTable As clsTable = Database.SysInfo.Tables(strTableName)
        Dim objField As clsField = Database.SysInfo.Fields(objTable.ID & "_" & strField)
        Return HasFieldAccess(objField, intTypeID)
    End Function

    Public Function HasFieldAccess(ByVal objField As clsField,
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As Boolean
        Dim blnHasAccess As Boolean = False

        If (HasAccess(objField.SecurityID) _
        AndAlso Me.LinkFields(CType(objField.ID, String)) IsNot Nothing) Then
            blnHasAccess = True
        Else
            blnHasAccess = False
        End If

        If Not intTypeID = clsDBConstants.cintNULL Then
            '2016-09-22 -- Peter & James -- Bug fix for #1600003206
            Dim colAccessRightField As New List(Of clsAccessRightField)

            For Each intSecurityGroupID As Integer In Database.Profile.LinkSecurityGroups.Values
                If Database.Profile.TypeFields(CStr(objField.ID & "_" & intTypeID & "_" & intSecurityGroupID)) IsNot Nothing Then
                    colAccessRightField.Add(Database.Profile.TypeFields(CStr(objField.ID & "_" & intTypeID & "_" & intSecurityGroupID)))
                End If
            Next

            For Each objAccessRightField As clsAccessRightField In colAccessRightField
                If objAccessRightField IsNot Nothing AndAlso Not objAccessRightField.Visible Then
                    blnHasAccess = False
                Else
                    blnHasAccess = True
                    Exit For
                End If
            Next

            'Dim objARF As clsAccessRightField = Me.TypeFields( _
            '    CStr(objField.ID & "_" & intTypeID))

            'If objARF IsNot Nothing AndAlso Not objARF.Visible Then
            '    blnHasAccess = False
            'End If
        End If

        Return blnHasAccess
    End Function
#End Region

#Region " HasMethodAccess "

    Public Function HasMethodAccess(ByVal strTableName As String, ByVal eMethod As clsMethod.enumMethods,
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As Boolean
        Dim objTable As clsTable = Database.SysInfo.Tables(strTableName)
        Dim objMethod As clsMethod = Database.SysInfo.Methods(CStr(eMethod))
        Dim objTableMethod As clsTableMethod = objTable.TableMethods(objMethod.ID.ToString())

        Return HasMethodAccess(objTableMethod, intTypeID)
    End Function

    Public Function HasMethodAccess(ByVal objTableMethod As clsTableMethod,
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As Boolean
        Dim blnHasAccess As Boolean = False

        '[Naing] Part of fix for Bug: 1300002475
        If (objTableMethod Is Nothing) Then
            Return blnHasAccess
        End If

        '[Naing] Access granted by Inclusion in SecurityGroup.LinkMethods
        If (HasAccess(objTableMethod.SecurityID) AndAlso
            Not Me.LinkMethods(CType(objTableMethod.ID, String)) Is Nothing) Then
            blnHasAccess = True
        Else
            blnHasAccess = False
        End If

        If Not intTypeID = clsDBConstants.cintNULL Then

            '[Naing] Access not granted by Inclusion in SecurityGroup.TypeMethods. Not consistent as above so caused some confusion.
            Dim objArm As clsAccessRightMethod = Me.TypeMethods(CStr(objTableMethod.ID & "_" & intTypeID))

            If objArm IsNot Nothing Then
                blnHasAccess = False
            End If
        End If

        Return blnHasAccess
    End Function
#End Region

#End Region
    Public Function CanModifyLegalHold(ByVal intID As Integer) As Boolean
        Dim intResult As Integer = m_objDB.ExecuteScalar("Select Count(ID) from LegalHold where ID='" & intID & "' and (AuthorizedPersonID='" & MyBase.PersonID & "' or AuthorizedPersonID is null)")
        Return CBool(intResult)
    End Function

    '2017-08-30 -- Peter Melisi -- Changes for Timezones for User Profiles
    Public Function ToLocalTime(ByVal objDatetime As DateTime, Optional ByVal blnToServerTime As Boolean = False) As DateTime 'localdatetime
        Try
            Dim serverTz As TimeZoneInfo = ServerTimeZone(m_objDB)

            If ClientTimeZone Is Nothing _
            OrElse serverTz Is Nothing _
            OrElse objDatetime = Nothing _
            OrElse ClientTimeZone.Equals(serverTz) Then
                Return objDatetime
            End If

            ' Ara Melkonian 2000003639
            ' Correctly convert to the server datetime
            Return If(blnToServerTime,
                TimeZoneInfo.ConvertTime(objDatetime, ClientTimeZone, serverTz),
                TimeZoneInfo.ConvertTime(objDatetime, serverTz, ClientTimeZone))
        Catch ex As Exception
            'log error?
            Throw
        End Try
    End Function

    Public Sub CloneDefaultSecurityGroup(ByVal objDB As clsDB)
        'Dim defaultUser = clsUserProfile.GetDefault(objDB, STRDP)
        Dim objDT As DataTable = m_objDB.GetDataTableBySQL($"SELECT [{clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID}] " &
                                                           $"FROM [{clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP}] " &
                                                           $"WHERE [{clsDBConstants.Fields.LinkUserProfileSecurityGroup.cUSERPROFILEID}]={objDB.SysInfo.K1Configuration.DefaultProfileID}")

        For Each objRow As DataRow In objDT.Rows
            objDB.InsertLink(clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP,
                             clsDBConstants.Fields.LinkUserProfileSecurityGroup.cUSERPROFILEID, m_intID,
                             clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID, CInt(objRow($"{clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID}")))
        Next
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            MyBase.DisposeDBObject()

            If Not m_objPerson Is Nothing Then
                m_objPerson.Dispose()
                m_objPerson = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region


End Class
