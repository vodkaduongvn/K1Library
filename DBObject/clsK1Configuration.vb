#Region " File Information "

'=====================================================================
' This class represents the table K1Configuration in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date            Description
'---------------------------------------------------------------------
' KD        11/05/2005      Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsK1Configuration
    Inherits clsDBObjBase

#Region " Members "

    Private m_intSessionTimeout As Integer
    Private m_intSessionTimeoutScanning As Integer
    Private m_intRecordLockTimeout As Integer
    Private m_blnAuditLogins As Boolean
    Private m_blnAuditLogoffs As Boolean
    Private m_blnAuditUnsuccessfulLogins As Boolean
    Private m_intMinCaptionWidth As Integer
    Private m_intMaxCaptionWidth As Integer
    Private m_intRecordsPerPage As Integer
    Private m_intTacitRefreshInterval As Integer 'Seconds
    Private m_objTacitCurrencyPeriod As clsPeriod
    Private m_intTacitCurrencyPeriodID As Integer
    Private m_intTacitCurrencyPeriodUnit As Integer
    Private m_intTacitDefaultWeightingID As Integer
    Private m_intRecordSetLimit As Integer
    Private m_intDRMDefaultSecurityID As Integer
    Private m_dblDbVersion As Double
    Private m_strAdminEmail As String
    Private m_blnUseAutomaticLogins As Boolean
    Private m_blnSaturdayNonWorkingDay As Boolean
    Private m_blnSundayNonWorkingDay As Boolean
    Private m_blnShowImageColInList As Boolean
    Private m_intDefaultProfileID As Integer
    Private m_objDefaultProfile As clsUserProfileBase
    Private m_strSMTPServer As String
    '2015-05-27 -- Peter Melisi -- O'Neil Integration
    Private m_strWebServicesURL As String
    Private m_intAbstractMaxSentence As Integer
    Private m_blnUseLatestVersions As Boolean
    Private m_intThumbnailSize As Integer
    Private m_strWebSiteURL As String
    Private m_ePasswordType As clsDBConstants.enumPasswordType
    Private m_eSessionTimeoutType As clsDBConstants.enumSessionTimeoutType
    Private m_strActiveDirectoryName As String
    '2017-02-17 -- Peter Melisi -- New Licensing Model
    Private m_strWebServiceLastChecked As String
    Private m_blnPasswordStrength As Boolean
    Private m_strPasswordStrengthSettings As String
    '2020-11-24 --Ara Melkonian -- Azure Active Directory
    Private m_strActiveDirectoryTenant As String
    Private m_strActiveDirectoryClientId As String
    Private m_strActiveDirectoryAuthority As String
    Private m_strActiveDirectoryWebRedirectUri As String
    Private m_strActiveDirectoryAppRedirectUri As String
    '2021-03-31 -- Ara Melkonian -- DocumentHover
    Private m_blnHoverEnabled As Boolean
#End Region

#Region " Constants "

    Public Const cstrAPPVERSION As String = "1.0"
    Public Const cstrMINORAPPVERSION As String = ""
    Public Const cdblVALID_DATABASE_VERSION As Double = 11
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_intSessionTimeout = clsDBConstants.cintNULL
        m_intSessionTimeoutScanning = clsDBConstants.cintNULL
        m_intRecordLockTimeout = clsDBConstants.cintNULL
        m_blnAuditLogins = False
        m_blnAuditLogoffs = False
        m_blnAuditUnsuccessfulLogins = False
        m_intMinCaptionWidth = clsDBConstants.cintNULL
        m_intMaxCaptionWidth = clsDBConstants.cintNULL
        m_intRecordsPerPage = clsDBConstants.cintNULL
        m_intTacitRefreshInterval = clsDBConstants.cintNULL
        m_intTacitCurrencyPeriodID = clsDBConstants.cintNULL
        m_intTacitDefaultWeightingID = clsDBConstants.cintNULL
        m_intTacitCurrencyPeriodUnit = clsDBConstants.cintNULL
        m_intRecordSetLimit = clsDBConstants.cintNULL
        m_intDRMDefaultSecurityID = clsDBConstants.cintNULL
        m_strAdminEmail = clsDBConstants.cstrNULL
        m_blnUseAutomaticLogins = False
        m_blnSaturdayNonWorkingDay = True
        m_blnSundayNonWorkingDay = True
        m_blnShowImageColInList = True
        m_intDefaultProfileID = clsDBConstants.cintNULL
        m_strSMTPServer = clsDBConstants.cstrNULL
        '2015-05-27 -- Peter Melisi -- O'Neil Integration
        m_strWebServicesURL = clsDBConstants.cstrNULL
        m_intAbstractMaxSentence = clsDBConstants.cintNULL
        m_blnUseLatestVersions = False
        m_intThumbnailSize = clsDBConstants.cintNULL
        m_strWebSiteURL = clsDBConstants.cstrNULL
        m_ePasswordType = clsDBConstants.enumPasswordType.RECFIND_DATABASE
        m_eSessionTimeoutType = clsDBConstants.enumSessionTimeoutType.WHEN_LICENCE_NEEDED
        '2017-02-17 -- Peter Melisi -- New Licensing Model
        m_strWebServiceLastChecked = clsDBConstants.cstrNULL
        m_blnPasswordStrength = False
        m_strPasswordStrengthSettings = clsDBConstants.cstrNULL
        '2020-11-24 --Ara Melkonian -- Azure Active Directory
        m_strActiveDirectoryTenant = clsDBConstants.cstrNULL
        m_strActiveDirectoryClientId = clsDBConstants.cstrNULL
        m_strActiveDirectoryAuthority = clsDBConstants.cstrNULL
        m_strActiveDirectoryWebRedirectUri = clsDBConstants.cstrNULL
        m_strActiveDirectoryAppRedirectUri = clsDBConstants.cstrNULL
        '2021-03-31 -- Ara Melkonian -- DocumentHover
        m_blnHoverEnabled = False
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intSessionTimeout = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cSESSIONTIMEOUT, clsDBConstants.cintNULL), Integer)
        m_intSessionTimeoutScanning = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cSESSIONTIMEOUTSCANNING, clsDBConstants.cintNULL), Integer)
        m_intRecordLockTimeout = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cRECORDLOCKTIMEOUT, clsDBConstants.cintNULL), Integer)
        m_blnAuditLogins = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cAUDITLOGINS, False), Boolean)
        m_blnAuditLogoffs = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cAUDITLOGOFFS, False), Boolean)
        m_blnAuditUnsuccessfulLogins = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cAUDITUNSUCCESSFULLOGINS, False), Boolean)
        m_intMinCaptionWidth = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cMINCAPTIONWIDTH, clsDBConstants.cintNULL), Integer)
        m_intMaxCaptionWidth = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cMAXCAPTIONWIDTH, clsDBConstants.cintNULL), Integer)
        m_intRecordsPerPage = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cRECORDSPERPAGE, clsDBConstants.cintNULL), Integer)
        m_intTacitRefreshInterval = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cTACITREFRESHINTERVAL, clsDBConstants.cintNULL), Integer)
        m_intTacitCurrencyPeriodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cTACITCURRENCYPERIODID, clsDBConstants.cintNULL), Integer)
        m_intTacitDefaultWeightingID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cTACITDEFAULTWEIGHTINGID, clsDBConstants.cintNULL), Integer)
        m_intTacitCurrencyPeriodUnit = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cTACITCURRENCYPERIODUNIT, clsDBConstants.cintNULL), Integer)
        m_intRecordSetLimit = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cRECORDSETLIMIT, clsDBConstants.cintNULL), Integer)
        m_dblDbVersion = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cDATABASEVERSION, clsDBConstants.cintNULL), Double)
        m_intDRMDefaultSecurityID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cDRMDEFAULTSECURITYID, clsDBConstants.cintNULL), Integer)
        m_strAdminEmail = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cADMINEMAIL, clsDBConstants.cstrNULL), String).Trim
        m_blnUseAutomaticLogins = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cUSEAUTOMATICLOGINS, True), Boolean)
        m_blnSaturdayNonWorkingDay = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cISSATURDAYNONWORKINGDAY, True), Boolean)
        m_blnSundayNonWorkingDay = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cISSUNDAYNONWORKINGDAY, True), Boolean)
        m_blnShowImageColInList = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cDISPLAYIMAGECOLINLIST, True), Boolean)
        m_intDefaultProfileID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cDEFAULTUSERPROFILEID, clsDBConstants.cintNULL), Integer)
        m_strSMTPServer = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cSMTPSERVER, clsDBConstants.cstrNULL), String).Trim
        '2015-05-27 -- Peter Melisi -- O'Neil Integration
        m_strWebServicesURL = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cWEBSERVICESURL, clsDBConstants.cstrNULL), String).Trim
        m_intAbstractMaxSentence = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cABSTRACTMAXSENTENCE, clsDBConstants.cintNULL), Integer)
        m_blnUseLatestVersions = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cCHECKOUTLATESTVERSIONONLY, False), Boolean)
        m_intThumbnailSize = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cTHUMBNAILSIZE, clsDBConstants.cintNULL), Integer)
        m_strWebSiteURL = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cWEBSITEURL, clsDBConstants.cstrNULL), String).Trim
        m_eSessionTimeoutType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cSESSIONTIMEOUTTYPE, clsDBConstants.enumSessionTimeoutType.WHEN_LICENCE_NEEDED), clsDBConstants.enumSessionTimeoutType)
        m_ePasswordType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cPASSWORDTYPE, clsDBConstants.enumPasswordType.RECFIND_DATABASE), clsDBConstants.enumPasswordType)
        m_strActiveDirectoryName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cACTIVEDIRECTORYNAME, clsDBConstants.cstrNULL), String).Trim
        If m_intSessionTimeout <= 0 Then
            m_intSessionTimeout = 30
        End If
        If m_intSessionTimeoutScanning <= 0 Then
            m_intSessionTimeoutScanning = 30
        End If
        If m_intRecordLockTimeout < 0 Then
            m_intRecordLockTimeout = 0
        End If
        If m_intRecordsPerPage <= 0 Then
            m_intRecordsPerPage = 1
        End If
        If m_intRecordsPerPage > 1000 Then
            m_intRecordsPerPage = 1000
        End If
        If m_intRecordSetLimit <= 0 Then
            m_intRecordSetLimit = 1
        End If
        If m_intRecordSetLimit > 10000000 Then
            m_intRecordSetLimit = 10000000
        End If
        '2017-02-17 -- Peter Melisi -- New Licensing Model
        m_strWebServiceLastChecked = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cWEBSERVICELASTCHECKED, clsDBConstants.cstrNULL), String).Trim
        m_blnPasswordStrength = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cUSEPASSWORDSTRENGTH, False), Boolean)
        m_strPasswordStrengthSettings = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cPASSWORDSTRENGTHSETTINGS, clsDBConstants.cstrNULL), String).Trim
        '2020-11-24 --Ara Melkonian -- Azure Active Directory
        m_strActiveDirectoryTenant = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cActiveDirectoryTenant, clsDBConstants.cstrNULL), String).Trim
        m_strActiveDirectoryClientId = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cActiveDirectoryClientId, clsDBConstants.cstrNULL), String).Trim
        m_strActiveDirectoryAuthority = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cActiveDirectoryAuthority, clsDBConstants.cstrNULL), String).Trim
        m_strActiveDirectoryWebRedirectUri = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cActiveDirectoryWebRedirectUri, clsDBConstants.cstrNULL), String).Trim
        m_strActiveDirectoryAppRedirectUri = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cActiveDirectoryAppRedirectUri, clsDBConstants.cstrNULL), String).Trim
        '2021-03-31 -- Ara Melkonian -- DocumentHover
        m_blnHoverEnabled = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Configuration.cHoverEnabled, False), Boolean)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property SessionTimeout() As Integer
        Get
            Return m_intSessionTimeout
        End Get
    End Property

    Public ReadOnly Property SessionTimeoutScanning() As Integer
        Get
            Return m_intSessionTimeoutScanning
        End Get
    End Property

    Public ReadOnly Property RecordLockTimeout() As Integer
        Get
            Return m_intRecordLockTimeout
        End Get
    End Property

    Public ReadOnly Property AuditLogins() As Boolean
        Get
            Return m_blnAuditLogins
        End Get
    End Property

    Public ReadOnly Property AuditLogoffs() As Boolean
        Get
            Return m_blnAuditLogoffs
        End Get
    End Property

    Public ReadOnly Property AuditUnsuccessfulLogins() As Boolean
        Get
            Return m_blnAuditUnsuccessfulLogins
        End Get
    End Property

    Public ReadOnly Property MinCaptionWidth() As Integer
        Get
            Return m_intMinCaptionWidth
        End Get
    End Property

    Public ReadOnly Property MaxCaptionWidth() As Integer
        Get
            Return m_intMaxCaptionWidth
        End Get
    End Property

    Public ReadOnly Property RecordsPerPage() As Integer
        Get
            Return m_intRecordsPerPage
        End Get
    End Property

    Public ReadOnly Property TacitRefreshInterval() As Integer
        Get
            Return m_intTacitRefreshInterval
        End Get
    End Property

    Public ReadOnly Property TacitCurrencyPeriod() As clsPeriod
        Get
            If m_objTacitCurrencyPeriod Is Nothing Then
                If Not m_intTacitCurrencyPeriodID = clsDBConstants.cintNULL Then
                    m_objTacitCurrencyPeriod = clsPeriod.GetItem(m_intTacitCurrencyPeriodID, Me.Database)
                End If
            End If
            Return m_objTacitCurrencyPeriod
        End Get
    End Property

    Public ReadOnly Property TacitCurrencyPeriodID() As Integer
        Get
            Return m_intTacitCurrencyPeriodID
        End Get
    End Property

    Public ReadOnly Property TacitCurrencyPeriodUnit() As Integer
        Get
            Return m_intTacitCurrencyPeriodUnit
        End Get
    End Property

    Public ReadOnly Property TacitDefaultWeightingID() As Integer
        Get
            Return m_intTacitDefaultWeightingID
        End Get
    End Property

    Public ReadOnly Property RecordSetLimit() As Integer
        Get
            Return m_intRecordSetLimit
        End Get
    End Property

    Public ReadOnly Property DbVersion() As Double
        Get
            Return m_dblDbVersion
        End Get
    End Property

    Public ReadOnly Property AppVersion() As String
        Get
            Return cstrAPPVERSION
        End Get
    End Property

    Public ReadOnly Property DRMDefaultSecurityID() As Integer
        Get
            Return m_intDRMDefaultSecurityID
        End Get
    End Property

    Public ReadOnly Property AdminEmail() As String
        Get
            Return m_strAdminEmail
        End Get
    End Property

    Public ReadOnly Property SMTPServer() As String
        Get
            Return m_strSMTPServer
        End Get
    End Property

    '2015-05-27 -- Peter Melisi -- O'Neil Integration
    Public ReadOnly Property WebServicesURL() As String
        Get
            Return m_strWebServicesURL
        End Get
    End Property

    Public ReadOnly Property UseAutomaticLogins() As Boolean
        Get
            Return m_blnUseAutomaticLogins
        End Get
    End Property

    Public ReadOnly Property IsSaturdayNonWorkingDay() As Boolean
        Get
            Return m_blnSaturdayNonWorkingDay
        End Get
    End Property

    Public ReadOnly Property IsSundayNonWorkingDay() As Boolean
        Get
            Return m_blnSundayNonWorkingDay
        End Get
    End Property

    Public ReadOnly Property ShowImageColInList() As Boolean
        Get
            Return m_blnShowImageColInList
        End Get
    End Property

    Public ReadOnly Property DefaultProfileID() As Integer
        Get
            Return m_intDefaultProfileID
        End Get
    End Property

    Public ReadOnly Property DefaultProfile() As clsUserProfileBase
        Get
            If m_objDefaultProfile Is Nothing Then
                If Not m_intDefaultProfileID = clsDBConstants.cintNULL Then
                    m_objDefaultProfile = clsUserProfileBase.GetItem(m_intDefaultProfileID, Me.Database)
                End If
            End If
            Return m_objDefaultProfile
        End Get
    End Property

    Public ReadOnly Property IsDefaultProfileLoaded() As Boolean
        Get
            Return (m_objDefaultProfile IsNot Nothing)
        End Get
    End Property

    Public ReadOnly Property AbstractMaxSentence() As Integer
        Get
            Return m_intAbstractMaxSentence
        End Get
    End Property

    Public ReadOnly Property CheckOutLatestVersionOnly() As Boolean
        Get
            Return m_blnUseLatestVersions
        End Get
    End Property

    Public ReadOnly Property ThumbnailSize() As Integer
        Get
            Return m_intThumbnailSize
        End Get
    End Property

    Public ReadOnly Property WebSiteURL() As String
        Get
            Return m_strWebSiteURL
        End Get
    End Property

    Public ReadOnly Property SessionTimeoutType As clsDBConstants.enumSessionTimeoutType
        Get
            Return m_eSessionTimeoutType
        End Get
    End Property

    Public ReadOnly Property PasswordType As clsDBConstants.enumPasswordType
        Get
            Return m_ePasswordType
        End Get
    End Property

    Public ReadOnly Property ActiveDirectoryName As String
        Get
            Return m_strActiveDirectoryName
        End Get
    End Property

    '2017-02-17 -- Peter Melisi -- New Licensing Model
    Public ReadOnly Property WebServiceLastChecked As String
        Get
            Return m_strWebServiceLastChecked
        End Get
    End Property

    Public ReadOnly Property UsePasswordStrength As Boolean
        Get
            Return m_blnPasswordStrength
        End Get
    End Property

    Public ReadOnly Property PasswordStrengthSettings As String
        Get
            Return m_strPasswordStrengthSettings
        End Get
    End Property

    '2020-11-24 Azure Active Directory
    Public ReadOnly Property ActiveDirectoryTenant As String
        Get
            Return m_strActiveDirectoryTenant
        End Get
    End Property

    Public ReadOnly Property ActiveDirectoryClientId As String
        Get
            Return m_strActiveDirectoryClientId
        End Get
    End Property

    Public ReadOnly Property ActiveDirectoryAuthority As String
        Get
            Return m_strActiveDirectoryAuthority
        End Get
    End Property

    Public ReadOnly Property ActiveDirectoryWebRedirectUri As String
        Get
            Return m_strActiveDirectoryWebRedirectUri
        End Get
    End Property

    Public ReadOnly Property ActiveDirectoryAppRedirectUri As String
        Get
            Return m_strActiveDirectoryAppRedirectUri
        End Get
    End Property

    '2021-03-31 Document Hover
    Public ReadOnly Property HoverEnabled As Boolean
        Get
            Return m_blnHoverEnabled
        End Get
    End Property
#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsK1Configuration
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cK1CONFIGURATION, intID)

            Return New clsK1Configuration(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Shared Function GetDefault(ByVal objDB As clsDB) As clsK1Configuration
        Try
            Dim objDT As DataTable

            objDT = objDB.GetDataTable(clsDBConstants.StoredProcedures.cUI_K1CONFIG_GETDEFAULT)

            If Not objDT Is Nothing AndAlso objDT.Rows.Count > 0 Then
                Return New clsK1Configuration(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function CanWorkWithLibrary(ByVal strAppVersion As String) As Boolean
        If strAppVersion.Trim = cstrAPPVERSION Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_objTacitCurrencyPeriod IsNot Nothing Then
                m_objTacitCurrencyPeriod.Dispose()
                m_objTacitCurrencyPeriod = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
