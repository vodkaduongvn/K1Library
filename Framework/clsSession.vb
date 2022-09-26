Imports System.Data.SqlClient

Public Class clsSession
    Implements IDisposable

#Region " Members "

    Private m_objDB As clsDB
    Private m_intID As Integer
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
    Private m_intSessionTimeOut As Integer
    Private m_eType As enumSessionType
    Private m_blnExpired As Boolean
    Private WithEvents m_tmrAutoUpdate As Timers.Timer

#End Region

#Region " Enumerations "

    Public Enum enumSessionType
        MAIN = 0
        CLONE = 1
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, ByVal intID As Integer)
        m_objDB = objDB
        m_intID = intID
        m_intSessionTimeOut = 35000
        m_eType = enumSessionType.CLONE
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal eAppType As clsDBConstants.enumApplicationType)
        m_objDB = objDB
        m_eType = enumSessionType.MAIN

        Dim objLicenceInfo As New Licensing.clsLicenceFileInfo(objDB, eAppType)
        Dim intActiveSessions As Integer = 0


        If m_objDB.SysInfo.K1Configuration.SessionTimeoutType = clsDBConstants.enumSessionTimeoutType.WHEN_LICENCE_NEEDED Then
            intActiveSessions = clsSession.GetSessionCount(objDB, eAppType)
            If objLicenceInfo.NumberOfUsers > 0 AndAlso intActiveSessions >= objLicenceInfo.NumberOfUsers Then
                DeleteExpiredSessions(objDB)
                intActiveSessions = clsSession.GetSessionCount(objDB, eAppType)
            End If
        ElseIf m_objDB.SysInfo.K1Configuration.SessionTimeoutType = clsDBConstants.enumSessionTimeoutType.WHEN_IN_USE Then
            DeleteApplicationSessions(objDB, eAppType)
            intActiveSessions = clsSession.GetSessionCount(objDB, eAppType)
        Else
            '-- Remove all expired sessions before starting
            DeleteExpiredSessions(objDB)
            intActiveSessions = clsSession.GetSessionCount(objDB, eAppType)
        End If

        '-- Make sure there are available user licenses
        If objLicenceInfo.NumberOfUsers > 0 AndAlso intActiveSessions >= objLicenceInfo.NumberOfUsers Then
            Throw New clsK1Exception(modErrors.ErrorNumber.No_Available_Licence, _
                                     "Maximum number of concurrent users exceeded. " & _
                                     "Please try again later or contact your Administrator.")
        End If
        '#End If

        Dim colParams As New clsDBParameterDictionary

        Dim strSP As String = clsDBConstants.Tables.cK1SESSION & clsDBConstants.StoredProcedures.cINSERT

        colParams.Add(New clsDBParameter(clsDB.ParamName( _
            clsDBConstants.Fields.cID), clsDBConstants.cintNULL, ParameterDirection.Output))
        colParams.Add(New clsDBParameter(clsDB.ParamName( _
            clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE), CInt(eAppType)))
        colParams.Add(New clsDBParameter(clsDB.ParamName( _
            clsDBConstants.Fields.K1Session.cUSERPROFILEID), objDB.Profile.ID))

        objDB.ExecuteProcedure(strSP, colParams)

        m_intID = CType(colParams(clsDB.ParamName(clsDBConstants.Fields.cID)).Value, Integer)

        '[Naing] 26/07/2013 Let me give a more clearer explanation then below. 
        'SqlDependency is only used when in Smart Client mode. Web client does not use this feature. 
        'Instead it uses some kind of background thread to poll the Session table. See clsSession.vb for implementation.
        If m_objDB.DataAccessType = clsDB.enumDataAccessType.DB_DIRECT AndAlso
            eAppType <> clsDBConstants.enumApplicationType.API Then
            'Only works if using direct access            
            AddSessionExpiredNotification()
        End If
        m_intSessionTimeOut = 35000

    End Sub

#End Region

#Region " Properties "

    Public Property Database() As clsDB
        Get
            Return m_objDB
        End Get
        Set(ByVal value As clsDB)
            m_objDB = value
        End Set
    End Property

    Public ReadOnly Property ID() As Integer
        Get
            Return m_intID
        End Get
    End Property

    Public ReadOnly Property SessionTimeout() As Integer
        Get
            Return m_intSessionTimeOut
        End Get
    End Property

    Public ReadOnly Property Type() As enumSessionType
        Get
            Return m_eType
        End Get
    End Property

    Public ReadOnly Property Expired() As Boolean
        Get
            Return m_blnExpired
        End Get
    End Property

#End Region

#Region " Methods "

    ''' <summary>
    ''' Adds a notification for when a session expires.
    ''' </summary>
    ''' <remarks>Starts the SQL Notification Listener if not started already.</remarks>
    Private Sub AddSessionExpiredNotification()

        ''-- Make sure we have a sql dependency object
        'If m_objDB.SqlDependency Is Nothing Then
        '    m_objDB.SqlDependency = New clsSqlDependency(CType(m_objDB, clsDB_Direct))
        'Else
        '    If m_objDB.SqlDependency.SystemFlags IsNot Nothing Then
        '        m_objDB.SqlDependency.SystemFlags.Refresh(m_objDB)
        '    End If
        'End If

        ''-- Make sure notification broker is started
        'If Not m_objDB.SqlDependency.Started Then
        '    m_objDB.SqlDependency.Start()
        'End If

        Dim objDb = CType(m_objDB, clsDB_Direct)
        If (objDb Is Nothing) Then
            Throw New InvalidCastException("m_objDB is not type of clsDB_Direct")
        End If
        objDb.TryInitializeSqlDependency()
        '2014-02-18 -- Peter Melisi -- Bug fix for #1400002637
        If m_objDB IsNot Nothing Then
            m_objDB.AddSqlNotification(String.Format("SELECT [{0}].[{1}] FROM [dbo].[{0}] WHERE [{1}] = {2}",
                                                     clsDBConstants.Tables.cK1SESSION,
                                                     clsDBConstants.Fields.cID,
                                                     m_intID),
                                                 AddressOf OnSessionExpiredNotification)
        End If
    End Sub


    ''' <summary>
    ''' Handles the notification when the session record in K1Sessions table is deleted.
    ''' </summary>
    Private Sub OnSessionExpiredNotification(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)
        Dim dependency As SqlDependency = CType(sender, SqlDependency)

        '-- Notices are only a one shot deal so remove the existing one 
        '-- so a new one can be added
        RemoveHandler dependency.OnChange, AddressOf OnSessionExpiredNotification

        If Not e.Info = SqlNotificationInfo.Delete AndAlso Not e.Info = SqlNotificationInfo.Update Then
            Throw New Exception(String.Format("SQL Dependency Error. Source: {0}, Info: {1}, Type {2} ", _
                                              e.Source, e.Info, e.ToString))
        End If

        '2014-02-18 -- Peter Melisi -- Bug fix for #1400002637
        If m_objDB IsNot Nothing And Not m_blnDisposedValue Then
            If e.Info = SqlNotificationInfo.Delete Then
                '-- Only raise event if record has been deleted
                m_blnExpired = True
                modEvents.RaiseSessionExpired()
            Else
                '-- Session was updated not deleted, so reapply notification
                AddSessionExpiredNotification()
            End If
        End If
    End Sub

    Public Sub Update()
        Dim colParams As New clsDBParameterDictionary

        Dim strSP As String = clsDBConstants.Tables.cK1SESSION & clsDBConstants.StoredProcedures.cUPDATE

        colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.cID), m_intID))

        m_objDB.ExecuteProcedure(strSP, colParams)
        m_objDB.SessionUpdatedMethod()
    End Sub

    Public Sub Delete()
        Dim colParams As New clsDBParameterDictionary

        Dim strSP As String = clsDBConstants.Tables.cK1SESSION & clsDBConstants.StoredProcedures.cDELETE

        colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.cID), m_intID))

        m_objDB.ExecuteProcedure(strSP, colParams)
    End Sub

    Public Shared Function GetActiveSessions(ByVal objDB As clsDB) As DataTable
        Try
            Dim SessionTimeOut As Integer = objDB.SysInfo.K1Configuration.SessionTimeout
            Dim tsTimeout As New TimeSpan(0, SessionTimeOut, 0)
            Dim dtValidDate As DateTime = objDB.GetCurrentTime.Subtract(tsTimeout)

            '-- Remove all expired sessions before starting
            'DeleteExpiredSessions(objDB, dtValidDate)

            Dim strQuery As String = "SELECT S.[{0}], S.[{1}], S.[{2}], UP.[{3}] as UserName, S.[{4}] " &
                " FROM [{5}] as S" &
                " LEFT JOIN [{6}] as UP ON UP.[{0}] = S.[{2}]"
            strQuery = String.Format(strQuery, clsDBConstants.Fields.cID,
                                     clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE,
                                     clsDBConstants.Fields.K1Session.cUSERPROFILEID,
                                     clsDBConstants.Fields.cEXTERNALID,
                                     clsDBConstants.Fields.K1Session.cLAST_UPDATED,
                                     clsDBConstants.Tables.cK1SESSION,
                                     clsDBConstants.Tables.cUSERPROFILE)

            'Dim colParams As New K1Library.clsDBParameterDictionary()
            'colParams.Add(New K1Library.clsDBParameter("@ValidSessionStart", dtValidDate))

            Dim objDT As DataTable = objDB.GetDataTableBySQL(strQuery) ', colParams)
            If objDT IsNot Nothing Then
                objDT.Columns.Add("Application", GetType(String))

                For Each objRow As DataRow In objDT.Rows
                    'objRow("Application") = CType(objRow(clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE), clsDBConstants.enumApplicationType).ToString
                    objRow("Application") = GetProductName(CType(objRow(clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE), clsDBConstants.enumApplicationType))
                Next
            End If

            Return objDT
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetActiveSessionCount(ByVal objDB As clsDB) As Integer
        Return GetActiveSessionCount(objDB, Nothing)
    End Function


    Public Shared Function GetSessionCount(ByVal objDB As clsDB,
                                             ByVal eAppType As clsDBConstants.enumApplicationType) As Integer
        Try
            '-- Remove all expired sessions before starting
            'DeleteExpiredSessions(objDB, dtValidDate)

            Dim strQuery As String = "SELECT COUNT([{0}]) FROM [{1}]"
            strQuery = String.Format(strQuery, clsDBConstants.Fields.cID, clsDBConstants.Tables.cK1SESSION)

            Dim colParams As New K1Library.clsDBParameterDictionary()

            If eAppType > 0 AndAlso [Enum].IsDefined(GetType(clsDBConstants.enumApplicationType), eAppType) Then
                '-- Filter Sessions
                strQuery &= String.Format(" WHERE [{0}] = @AppType",
                                          clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE)

                colParams.Add(New K1Library.clsDBParameter("@AppType", eAppType))
            End If

            Return objDB.ExecuteScalar(strQuery, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function DeleteApplicationSessions(ByVal objDB As clsDB,
                                         ByVal eAppType As clsDBConstants.enumApplicationType) As Integer
        Try
            Dim strQuery As String = "DELETE FROM [{0}] WHERE [{1}] = @AppType AND [{2}] = @UserID"
            strQuery = String.Format(strQuery, clsDBConstants.Tables.cK1SESSION, clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE, clsDBConstants.Fields.K1Session.cUSERPROFILEID)

            Dim colParams As New K1Library.clsDBParameterDictionary()
            colParams.Add(New K1Library.clsDBParameter("@AppType", eAppType))
            colParams.Add(New K1Library.clsDBParameter("@UserID", objDB.Profile.ID))

            Return objDB.ExecuteScalar(strQuery, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetActiveUserSessionCount(ByVal objDB As clsDB, _
                                     ByVal eAppType As clsDBConstants.enumApplicationType) As Integer
        Try
            Dim strQuery As String = "SELECT COUNT([{0}]) FROM [{1}] WHERE [{2}] = @AppType AND [{3}] = @UserID"
            strQuery = String.Format(strQuery, clsDBConstants.Fields.cID, clsDBConstants.Tables.cK1SESSION, clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE, clsDBConstants.Fields.K1Session.cUSERPROFILEID)

            Dim colParams As New K1Library.clsDBParameterDictionary()
            colParams.Add(New K1Library.clsDBParameter("@AppType", eAppType))
            colParams.Add(New K1Library.clsDBParameter("@UserID", objDB.Profile.ID))

            Return objDB.ExecuteScalar(strQuery, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetActiveSessionCount(ByVal objDB As clsDB, _
                                                 ByVal eAppType As clsDBConstants.enumApplicationType) As Integer
        Try
            Dim SessionTimeOut As Integer = objDB.SysInfo.K1Configuration.SessionTimeout
            Dim tsTimeout As New TimeSpan(0, SessionTimeOut, 0)
            Dim dtValidDate As DateTime = objDB.GetCurrentTime.Subtract(tsTimeout)

            '-- Remove all expired sessions before starting
            'DeleteExpiredSessions(objDB, dtValidDate)

            Dim strQuery As String = "SELECT COUNT([{0}]) FROM [{1}] WHERE [{2}] > @ValidSessionStart"
            strQuery = String.Format(strQuery, clsDBConstants.Fields.cID, clsDBConstants.Tables.cK1SESSION, _
                                     clsDBConstants.Fields.K1Session.cLAST_UPDATED)

            Dim colParams As New K1Library.clsDBParameterDictionary()
            colParams.Add(New K1Library.clsDBParameter("@ValidSessionStart", dtValidDate))

            If eAppType > 0 AndAlso [Enum].IsDefined(GetType(clsDBConstants.enumApplicationType), eAppType) Then
                '-- Filter Sessions
                strQuery &= String.Format(" AND [{0}] = @AppType", _
                                          clsDBConstants.Fields.K1Session.cAPPLICATION_TYPE)

                colParams.Add(New K1Library.clsDBParameter("@AppType", eAppType))
            End If

            Return objDB.ExecuteScalar(strQuery, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function CheckSessionExists(ByVal objDB As clsDB, ByVal intSessionID As Integer) As Boolean
        Try
            Dim strQuery As String = "SELECT COUNT([{0}]) FROM [{1}] WHERE [{2}] = @SessionID"
            strQuery = String.Format(strQuery, clsDBConstants.Fields.cID, clsDBConstants.Tables.cK1SESSION, _
                                     clsDBConstants.Fields.cID)

            Dim colParams As New K1Library.clsDBParameterDictionary()
            colParams.Add(New K1Library.clsDBParameter("@SessionID", intSessionID))

            If (objDB.ExecuteScalar(strQuery, colParams) = 0) Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Sub DeleteExpiredSessions(ByVal objDB As clsDB, ByVal dtExpiredSessionDate As DateTime)
        Try
            Dim colParams As New K1Library.clsDBParameterDictionary
            colParams.Add(New K1Library.clsDBParameter("@expireTime", dtExpiredSessionDate))

            If objDB.SysInfo.K1Configuration.AuditLogoffs Then
                Dim objDT As DataTable = objDB.GetDataTableBySQL(String.Format( _
                    "SELECT UserProfileID, ApplicationType FROM [{0}] WHERE [{1}] < @expireTime", _
                    clsDBConstants.Tables.cK1SESSION, clsDBConstants.Fields.K1Session.cLAST_UPDATED), colParams)
                For Each objRow As DataRow In objDT.Rows
                    Dim strAppType As String = Nothing
                    Try
                        Dim eType As clsDBConstants.enumApplicationType = CType(objRow(1), clsDBConstants.enumApplicationType)

                        If eType = clsDBConstants.enumApplicationType.RecFind Then
                            strAppType = "RecFind 6/Web Client"
                        Else
                            strAppType = CType(eType, String)
                        End If
                        clsAuditTrail.CreateSessionTimeoutRecord(objDB, CInt(objRow(0)), strAppType)
                    Catch ex As Exception

                    End Try
                Next
            End If
            Dim strQuery As String = "DELETE FROM [{0}] WHERE [{1}] < @expireTime"

            strQuery = String.Format(strQuery, clsDBConstants.Tables.cK1SESSION, clsDBConstants.Fields.K1Session.cLAST_UPDATED)

            objDB.ExecuteSQL(strQuery, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Shared Sub DeleteExpiredSessions(ByVal objDB As clsDB)
        Try
            Dim tsSessionLength As New TimeSpan(0, objDB.SysInfo.K1Configuration.SessionTimeout, 0)
            Dim dtExpiredSessionDate As DateTime = objDB.GetCurrentTime.Subtract(tsSessionLength)

            DeleteExpiredSessions(objDB, dtExpiredSessionDate)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    Public Function IsSessionActive() As Boolean
        Try
            Dim SessionTimeOut As Integer = m_objDB.SysInfo.K1Configuration.SessionTimeout
            Dim tsTimeout As New TimeSpan(0, SessionTimeOut, 0)
            Dim dtCurrent As DateTime = m_objDB.GetCurrentTime
            Dim dtValidDate As DateTime = dtCurrent.Subtract(tsTimeout)

            Dim strQuery As String = "SELECT COUNT([ID]) FROM [{0}] WHERE [{1}] = @sessionID AND [{2}] > @validSessionTime"
            strQuery = String.Format(strQuery, clsDBConstants.Tables.cK1SESSION, clsDBConstants.Fields.cID, _
                                     clsDBConstants.Fields.K1Session.cLAST_UPDATED)

            Dim colParams As New K1Library.clsDBParameterDictionary()
            colParams.Add(New K1Library.clsDBParameter("@sessionID", m_intID))
            colParams.Add(New K1Library.clsDBParameter("@validSessionTime", dtValidDate, ParameterDirection.Input, SqlDbType.Date))

            Return CType(m_objDB.ExecuteScalar(strQuery, colParams), Boolean)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub StartAutoUpdate()
        EndAutoUpdate()

        Dim intTimeout As Integer = m_objDB.SysInfo.K1Configuration.SessionTimeout - 1
        intTimeout = intTimeout * 60000
        If intTimeout = 0 Then intTimeout = 30000
        m_tmrAutoUpdate = New Timers.Timer(intTimeout)
        m_tmrAutoUpdate.Start()

        m_objDB.SessionUpdate()
    End Sub

    Public Sub EndAutoUpdate()
        If m_tmrAutoUpdate IsNot Nothing Then
            m_tmrAutoUpdate.Stop()
            m_tmrAutoUpdate.Dispose()
            m_tmrAutoUpdate = Nothing
        End If
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                EndAutoUpdate()
                m_objDB = Nothing
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Private Sub m_tmrAutoUpdate_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles m_tmrAutoUpdate.Elapsed
        Try
            m_tmrAutoUpdate.Stop()
            Me.Database.SessionUpdate()
            m_tmrAutoUpdate.Start()
        Catch ex As Exception
        End Try

    End Sub
End Class
