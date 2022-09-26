'[Naing] 12:23:00 PM obsolete do not use this class in future. No reason for it to exist.

Imports System.Data.SqlClient
Imports System.Security

Public Class clsSqlDependency
    Implements IDisposable

    '[Naing Begin] Fix for Bug 1300002478 was moved to clsDB_Direct dispose method
    Private Shared Internal_SqlDependency_Started As Boolean = False ' To detect redundant calls to SqlDependency.Start

    Protected m_objDB As clsDB_Direct
    Protected m_blnStarted As Boolean
    Protected m_objSystemFlags As K1Library.clsSystemFlags
    Private disposedValue As Boolean = False ' To detect redundant calls

    Public Sub New(ByVal objDb As clsDB_Direct)
        m_objDB = objDb
        m_objSystemFlags = New clsSystemFlags(m_objDB)
        m_blnStarted = False
    End Sub

    Public ReadOnly Property SystemFlags() As clsSystemFlags
        Get
            Return m_objSystemFlags
        End Get
    End Property

    Public ReadOnly Property Started() As Boolean
        Get
            Return m_blnStarted
        End Get
    End Property

    <SqlClientPermission(Permissions.SecurityAction.Assert, unrestricted:=True)> _
    Public Sub Start()
        Try

            Dim strConnection As String = m_objDB.ConnectionString

            '-- In order to use the callback feature of the SqlDependency, the application must have permission.
            'Dim perm As New SqlClientPermission(Permissions.PermissionState.Unrestricted)
            'perm.Demand()

            '-- Kill any existing notifications
            StopInt()

            '[Naing Begin] Fix for Bug 1300002478 was moved to clsDB_Direct dispose method
            If (Not Internal_SqlDependency_Started) Then
                '-- Start Sql Notification listener
                SqlDependency.Start(strConnection)
                Internal_SqlDependency_Started = True
            End If

            m_blnStarted = True

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Shared Sub StopBrokerAndResetFlag(connectionString As String)
        Try
            Internal_SqlDependency_Started = False
            SqlDependency.Stop(connectionString)
        Catch ex As Exception
            Dim msg = ex.Message
            If ex.InnerException IsNot Nothing Then
                msg &= ex.InnerException.Message
            End If

            Throw
        End Try

    End Sub

    Public Sub [Stop]()
        StopInt()
    End Sub

    Private Sub StopInt()
        If (m_objDB IsNot Nothing) Then
            Dim strConnection As String = m_objDB.ConnectionString
            'Kill any existing notifications
            '[Naing Begin] Fix for Bug 1300002478 was moved to clsDB_Direct dispose method
            'SqlDependency.Stop(strConnection)            
        End If
        m_blnStarted = False

    End Sub

#Region " Add DRM Notifications "

    ''' <summary>
    ''' Starts the SQL Notification Listener and adds a notification
    ''' for DRM_Refresh_Date and DRM_Forced_Logoff_Date
    ''' </summary>
    Public Sub AddDRMNotifications()
        Try
            If Not m_blnStarted Then
                Me.Start()
            End If

            '-- Add notification when Refresh Date is changed
            AddDRMNotification()

            '-- Add notification when Force Logoff Date is changed
            'AddDRMForcedLogoffNotification()

        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region


#Region " DRM Notification "

    ''' <summary>
    ''' Adds an SQL notification on the K1SystemFlags table
    ''' </summary>
    Private Sub AddDRMNotification()

        Dim strSql As String = String.Format("select [{0}].[DRM_Running], [{0}].[DRM_UserProfileID], [{0}].[DRM_Check_Date], [{0}].[DRM_Refresh_Date], [{0}].[DRM_Forced_Logoff_Date], [{0}].[DRM_Locked], [{0}].[ID] from [dbo].[{0}]", clsDBConstants.Tables.cK1SYSTEMFLAGS)

        m_objDB.AddSqlNotification(strSql, AddressOf OnDRMNotification)

    End Sub

#End Region

    '#Region " DRM Forced Logoff Notification "

    '    ''' <summary>
    '    ''' Adds an sql notification on DRM_Forced_Logoff_Date of K1SystemFlags table
    '    ''' </summary>
    '    Private Sub AddDRMForcedLogoffNotification()
    '        m_objDB.AddSqlNotification(clsDBConstants.Tables.cK1SYSTEMFLAGS, _
    '            clsDBConstants.Fields.K1SystemFlags.cDRM_LOCKED, AddressOf OnDRMForcedLogoffNotification)
    '    End Sub

    '#End Region

    '#Region " DRM Refresh Notification "

    '    ''' <summary>
    '    ''' Adds an sql notification on DRM_Refresh_Date of K1SystemFlags table
    '    ''' </summary>
    '    Private Sub AddDRMRefreshNotification()
    '        m_objDB.AddSqlNotification(clsDBConstants.Tables.cK1SYSTEMFLAGS, _
    '            clsDBConstants.Fields.K1SystemFlags.cDRM_REFRESH_DATE, AddressOf OnDRMRefreshNotification)
    '    End Sub

    '#End Region


#Region " On DRM Notification "

    ''' <summary>
    ''' Handles the notification when the K1SystemFlags table is changed
    ''' </summary>
    Private Sub OnDRMNotification(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)
        Dim dependency As SqlDependency = CType(sender, SqlDependency)

        '-- Notices are only a one shot deal so remove the existing one 
        '-- so a new one can be added
        RemoveHandler dependency.OnChange, AddressOf OnDRMNotification

        If m_objSystemFlags Is Nothing Then
            Return
        End If

        '2019-08-20 -- Emmanuel -- Cleaning up some code while looking for mem leak #1900003560
        If Not e.Info = SqlNotificationInfo.Update And Not (e.Source = SqlNotificationSource.Timeout Or e.Source = SqlNotificationSource.Statement) Then
            Throw New Exception(String.Format("SQL Dependency Error. Source: {0}, Info: {1}, Type {2} ",
                                              e.Source, e.Info, e.ToString))
        End If

        AddDRMNotification()

        Dim objSystemFlags As New K1Library.clsSystemFlags(m_objDB)
        If m_objSystemFlags IsNot Nothing AndAlso objSystemFlags.DRMForcedLogoffDate > m_objSystemFlags.DRMForcedLogoffDate Then
            '-- Only raise event if forced log-off date has changed
            modEvents.RaiseDRMLocked()
            m_objSystemFlags = objSystemFlags
        End If

        If m_objDB IsNot Nothing AndAlso objSystemFlags.DRMRefreshedDate > m_objDB.LastRefresh Then
            '-- Only raise event if refresh date has actually changed
            modEvents.RaiseRefreshEvent(objSystemFlags.DRMRefreshedDate)
            m_objSystemFlags = objSystemFlags
        End If
    End Sub

#End Region

#Region " AddNotification "

    ''' <summary>
    ''' Starts the SQL Notification Listener and adds a notification to the specified table
    ''' </summary>
    Public Sub AddNotification(ByVal strTable As String, ByVal strField As String, _
                               ByVal objCallbackHandler As System.Data.SqlClient.OnChangeEventHandler)
        Try
            If Not m_blnStarted Then
                Me.Start()
            End If

            '-- Add notification to table for changes
            m_objDB.AddSqlNotification(strTable, strField, objCallbackHandler)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Starts the SQL Notification Listener and adds a notification to the specified table
    ''' </summary>
    Public Sub AddNotification(ByVal strSQL As String, ByVal objCallbackHandler As System.Data.SqlClient.OnChangeEventHandler)
        Try
            If Not m_blnStarted Then
                Me.Start()
            End If

            '-- Add notification to table for changes
            m_objDB.AddSqlNotification(strSQL, objCallbackHandler)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                If m_objDB IsNot Nothing AndAlso m_blnStarted Then
                    StopInt()
                End If

                If m_objSystemFlags IsNot Nothing Then
                    m_objSystemFlags.Dispose()
                    m_objSystemFlags = Nothing
                End If
            End If
            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
