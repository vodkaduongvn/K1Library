Imports System.Web
Imports System.Web.SessionState
Imports System.Threading

Public Class clsDB_Web
    Inherits clsDB

    Private m_objDB As clsDB_Direct
    Private m_colDatabase As FrameworkCollections.K1Dictionary(Of clsDB_Direct) 'Databases by session
    Private m_strProduct As String

#Region " Constructors "

    ''' <summary>
    ''' Creates a new database object using the connection string provided
    ''' </summary>
    Public Sub New(ByVal strConnection As String, ByVal strProduct As String)
        m_eDataAccess = enumDataAccessType.WEB_SESSION
        m_objDB = New clsDB_Direct(strConnection, CType(Nothing, clsSysInfo))
        m_objDB.MultiThreaded = False
        m_colDatabase = New FrameworkCollections.K1Dictionary(Of clsDB_Direct)
        m_objSysInfo = New clsSysInfo(Me)
        m_objDB.SysInfo = m_objSysInfo
        DB.SysInfo = m_objSysInfo
        m_strProduct = strProduct
    End Sub
#End Region

#Region " Properties "

    Public Overrides Property Profile() As K1Library.clsUserProfile
        Get
            Return DB.Profile
        End Get
        Set(ByVal value As K1Library.clsUserProfile)
            DB.Profile = value
        End Set
    End Property

    Public Overrides ReadOnly Property HasTransaction() As Boolean
        Get
            Return DB.HasTransaction()
        End Get
    End Property

    ''' <summary>
    ''' If K1 is setup as a training version, this is record limit associated
    ''' </summary>
    Public Overrides Property RecordLimit() As Integer
        Get
            Return DB.RecordLimit
        End Get
        Set(ByVal value As Integer)
            DB.RecordLimit = value
        End Set
    End Property

    ''' <summary>
    ''' When reading/writing bytes to or from a blob field, this is the size of the chunks used
    ''' </summary>
    Public Overrides Property BlobChunkSize() As Integer
        Get
            Return m_intBlobChunkSize
        End Get
        Set(ByVal value As Integer)
            m_intBlobChunkSize = value
        End Set
    End Property

    Public Overrides Property ThreadedOperationCancelled() As Boolean
        Get
            Return DB.ThreadedOperationCancelled
        End Get
        Set(ByVal value As Boolean)
            DB.ThreadedOperationCancelled = value
        End Set
    End Property

    Public Overrides Property SqlDependency() As clsSqlDependency
        Get
            Return DB.SqlDependency
        End Get
        Set(ByVal value As clsSqlDependency)
            DB.SqlDependency = value
        End Set
    End Property

    Public Overrides Property LastRefresh() As DateTime
        Get
            Return m_dtLastRefreshed
        End Get
        Set(ByVal value As DateTime)
            m_dtLastRefreshed = value
        End Set
    End Property

    Public Overrides Property Session() As clsSession
        Get
            Return DB.Session
        End Get
        Set(ByVal value As clsSession)
            DB.Session = value
        End Set
    End Property

    Public Overrides Property State() As enumState
        Get
            Return DB.State
        End Get
        Set(ByVal value As enumState)
            DB.State = value
            'RaiseEvent StateChanged()
            DBStateChanged()
        End Set
    End Property

    Protected Overrides Sub DBStateChanged()
        If DB.State = enumState.DRM_LOCKED_SESSION Then
            DB.ThreadedOperationCancelled = True
        End If
    End Sub

    Public ReadOnly Property Product() As String
        Get
            Return m_strProduct
        End Get
    End Property

    Public ReadOnly Property ShouldCleanUp(ByVal strSession As String) As Boolean
        Get
            Dim objDB As clsDB_Direct = Nothing
            If strSession IsNot Nothing Then
                objDB = m_colDatabase(strSession)
            End If
            If objDB Is Nothing OrElse objDB.LastActivity = Date.MinValue Then
                Return False
            End If
            If Now.Subtract(objDB.LastActivity).Minutes >= objDB.SysInfo.K1Configuration.SessionTimeout Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
#End Region

#Region " Session Database "

    Public ReadOnly Property MasterDatabase() As clsDB_Direct
        Get
            Return m_objDB
        End Get
    End Property

    'Public Overrides ReadOnly Property dbSystem As clsDB_System
    '    Get
    '        Return m_objDB.dbSystem  
    '    End Get
    'End Property

    Private Function DB() As clsDB_Direct
        Dim objDB As clsDB_Direct = Nothing
        Dim strSession As String = String.Empty

        If System.Web.HttpContext.Current IsNot Nothing Then
            If System.Web.HttpContext.Current.Session IsNot Nothing Then
                strSession = System.Web.HttpContext.Current.Session.SessionID
            ElseIf HttpContext.Current.Cache.Item("Session" & Threading.Thread.CurrentThread.ManagedThreadId) IsNot Nothing Then
                strSession = HttpContext.Current.Cache.Item("Session" & Threading.Thread.CurrentThread.ManagedThreadId).ToString()
            End If
        End If

        If Not String.IsNullOrEmpty(strSession) Then
            objDB = m_colDatabase(strSession)

            If objDB Is Nothing Then
                objDB = New clsDB_Direct(m_objDB.ConnectionString, m_objSysInfo)
                objDB.MultiThreaded = False
                m_colDatabase.Add(strSession, objDB)
            End If
        End If

        If objDB Is Nothing Then
            objDB = m_objDB
        Else
            objDB.LastActivity = Now
        End If

        Return objDB
    End Function
#End Region

#Region " Overloads "

    Public Overloads Overrides Sub AddSqlNotification(ByVal strTable As String, ByVal strField As String, ByVal objCallBack As System.Data.SqlClient.OnChangeEventHandler)
        DB.AddSqlNotification(strTable, strField, objCallBack)
    End Sub

    Public Overloads Overrides Sub AddSqlNotification(ByVal strSQL As String, ByVal objCallBack As System.Data.SqlClient.OnChangeEventHandler)
        DB.AddSqlNotification(strSQL, objCallBack)
    End Sub

    Public Overloads Overrides Sub BeginTransaction()
        DB.BeginTransaction()
    End Sub

    Public Overloads Overrides Sub BeginTransaction(ByVal eIsolationLevel As System.Data.IsolationLevel)
        DB.BeginTransaction(eIsolationLevel)
    End Sub

    Public Overrides Sub BulkInsert(ByVal objDT As System.Data.DataTable, ByVal strDestTable As String)
        DB.BulkInsert(objDT, strDestTable)
    End Sub

    Public Overrides Sub DeleteLink(ByVal strLinkTable As String, ByVal strLinkField1 As String, ByVal intLinkID1 As Integer, ByVal strLinkField2 As String, ByVal intLinkID2 As Integer)
        DB.DeleteLink(strLinkTable, strLinkField1, intLinkID1, strLinkField2, intLinkID2)
    End Sub

    Public Overloads Overrides Sub DeleteRecord(ByVal strTable As String, ByVal intID As Integer)
        DB.DeleteRecord(strTable, intID)
    End Sub

    Public Overrides Sub EndTransaction(ByVal blnCommit As Boolean)
        DB.EndTransaction(blnCommit)
    End Sub

    Public Overloads Overrides Sub ExecuteProcedure(ByVal strStoredProcedure As String)
        DB.ExecuteProcedure(strStoredProcedure)
    End Sub

    Public Overloads Overrides Sub ExecuteProcedure(ByVal strStoredProcedure As String, ByRef colParams As K1Library.clsDBParameterDictionary)
        DB.ExecuteProcedure(strStoredProcedure, colParams)
    End Sub

    Public Overloads Overrides Function ExecuteScalar(ByVal strSQL As String) As Integer
        Return DB.ExecuteScalar(strSQL)
    End Function

    Public Overloads Overrides Function ExecuteScalar(ByVal strSQL As String, ByRef colParams As K1Library.clsDBParameterDictionary) As Integer
        Return DB.ExecuteScalar(strSQL, colParams)
    End Function

    Public Overloads Overrides Function ExecuteSQL(ByVal strSQL As String) As Integer
        Return DB.ExecuteSQL(strSQL)
    End Function

    Public Overloads Overrides Function ExecuteSQL(ByVal strSQL As String, ByRef colParams As K1Library.clsDBParameterDictionary) As Integer
        Return DB.ExecuteSQL(strSQL, colParams)
    End Function

    Public Overrides Function GetBLOBSize(ByVal strTableName As String, ByVal strFieldName As String, ByVal intID As Integer) As Integer
        Return DB.GetBLOBSize(strTableName, strFieldName, intID)
    End Function

    Public Overrides Sub GetDatabaseInfo(ByRef strServer As String, ByRef strDatabase As String,
    ByRef strUserID As String, Optional ByVal dblVersion As Double = 0,
    Optional ByVal strAppName As String = Nothing, Optional ByVal dblMinDBVersion As Double = 0)
        DB.GetDatabaseInfo(strServer, strDatabase, strUserID, dblVersion, strAppName, dblMinDBVersion)
    End Sub

    Public Overrides Sub GetDatabaseGroupInfo(ByRef strGroup As String, ByRef strServer As String, ByRef strDatabase As String, ByRef strUserID As String, Optional dblVersion As Double = 0, Optional strAppName As String = Nothing, Optional dblMinDBVersion As Double = 0)
        DB.GetDatabaseGroupInfo(strGroup, strServer, strDatabase, strUserID, dblVersion, strAppName, dblMinDBVersion)
    End Sub

    Public Overrides Function CheckHostName(ByVal strServer As String) As String
        Return DB.CheckHostName(strServer)
    End Function

    Public Overrides Function CheckSQLServerName() As String
        Return DB.CheckSQLServerName()
    End Function

    Public Overloads Overrides Function GetDataTable(ByVal objTable As K1Library.clsTable) As System.Data.DataTable
        Return DB.GetDataTable(objTable)
    End Function

    Public Overloads Overrides Function GetDataTable(ByVal strStoredProcedure As String) As System.Data.DataTable
        Return DB.GetDataTable(strStoredProcedure)
    End Function

    Public Overloads Overrides Function GetDataTable(ByVal strStoredProcedure As String, ByVal colParams As K1Library.clsDBParameterDictionary) As System.Data.DataTable
        Return DB.GetDataTable(strStoredProcedure, colParams)
    End Function

    Public Overloads Overrides Function GetDataTableByField(ByVal strTable As String, ByVal colParams As K1Library.clsDBParameterDictionary) As System.Data.DataTable
        Return DB.GetDataTableByField(strTable, colParams)
    End Function

    Public Overloads Overrides Function GetDataTableByField(ByVal strTable As String, ByVal strField As String, ByVal objValue As Object) As System.Data.DataTable
        Return DB.GetDataTableByField(strTable, strField, objValue)
    End Function

    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQL As String) As System.Data.DataTable
        Return DB.GetDataTableBySQL(strSQL)
    End Function

    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQL As String, ByVal colParams As K1Library.clsDBParameterDictionary) As System.Data.DataTable
        Return DB.GetDataTableBySQL(strSQL, colParams)
    End Function

    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQL As String, ByVal colParams As K1Library.clsDBParameterDictionary, ByVal blnValidateType As Boolean) As System.Data.DataTable
        Return DB.GetDataTableBySQL(strSQL, colParams, blnValidateType)
    End Function

    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQLFormat As String, ByVal colParams As K1Library.clsDBParameterDictionary, ByVal ParamArray args() As Object) As System.Data.DataTable
        Return DB.GetDataTableBySQL(strSQLFormat, colParams, args)
    End Function

    Public Overrides Function GetItem(ByVal strTable As String, ByVal intID As Integer) As System.Data.DataTable
        Return DB.GetItem(strTable, intID)
    End Function

    Public Overrides Function GetList(ByVal strTable As String, ByVal strByField As String, ByVal objValue As Object) As System.Data.DataTable
        Return DB.GetList(strTable, strByField, objValue)
    End Function

    Public Overrides Sub InitializeServerObjects()
        DB.InitializeServerObjects()
    End Sub

    Public Overrides Sub InsertLink(ByVal strLinkTable As String, ByVal strLinkField1 As String, ByVal intLinkID1 As Integer, ByVal strLinkField2 As String, ByVal intLinkID2 As Integer)
        DB.InsertLink(strLinkTable, strLinkField1, intLinkID1, strLinkField2, intLinkID2)
    End Sub

    Public Overrides Sub ReadBLOB(ByVal strTableName As String, ByVal strFieldName As String, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)
        DB.ReadBLOB(strTableName, strFieldName, intID, strFileName, blnShowTransferUI)
    End Sub

    Public Overrides Function ReadBLOBToMemory(ByVal strTableName As String, ByVal strFieldName As String, ByVal intID As Integer) As Byte()
        Return DB.ReadBLOBToMemory(strTableName, strFieldName, intID)
    End Function

    Public Overrides Function RecordCountExceeded(ByVal strTable As String) As Boolean
        Return DB.RecordCountExceeded(strTable)
    End Function

    Public Overloads Overrides Sub WriteBLOB(ByVal objTable As K1Library.clsTable, ByVal objField As K1Library.clsField, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)
        DB.WriteBLOB(objTable, objField, intID, strFileName, blnShowTransferUI)
    End Sub

    Public Overloads Overrides Sub WriteBLOB(ByVal strTable As String, ByVal strField As String, ByVal eDataType As System.Data.SqlDbType, ByVal intDataLength As Integer, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)
        DB.WriteBLOB(strTable, strField, eDataType, intDataLength, intID, strFileName, blnShowTransferUI)
    End Sub
#End Region

#Region " System Information Refresh "

    Public Overrides Sub RefreshSysInfo()
        m_dtLastRefreshed = m_objDB.GetCurrentTime
        Dim objSysInfoReload As New clsSysInfo(Me)
        m_objSysInfoReload = objSysInfoReload
    End Sub
#End Region

#Region " Session Functionality "

    Public Overrides Sub SessionStart(ByVal eAppType As clsDBConstants.enumApplicationType)
        DB.Session = New clsSession(Me, eAppType)
    End Sub

    Public Overrides Sub SessionStop()
        If Not DB.Session Is Nothing Then
            DB.Session.Delete()
            DB.Session.Dispose()
            DB.Session = Nothing
        End If
    End Sub

    Public Overrides Sub SessionUpdate()
        If Not DB.Session Is Nothing Then
            Dim objSessionThread As New Threading.Thread(AddressOf SessionUpdateThread)
            objSessionThread.IsBackground = True
            objSessionThread.Name = "Session Updater"
            objSessionThread.Start()
        End If
    End Sub

    Protected Overrides Sub SessionUpdateThread()
        Try
            If DB.Session Is Nothing Then
                Return
            End If

            If DB.Session.Expired Then
                RaiseSessionExpired()
                Return
            End If

            DB.Session.Update()

        Catch ex As Exception
            '-- ignore
        End Try

    End Sub
#End Region

#Region " Clean Up and Shut Down "

    Public Sub CleanUpAndShutDown(ByVal objSession As HttpSessionState)
        Dim objDB As clsDB_Direct = Nothing

        If objSession IsNot Nothing Then
            objDB = m_colDatabase(objSession.SessionID)
        End If

        If objDB IsNot Nothing Then
            Try
                If objDB.Profile IsNot Nothing Then
                    clsAuditTrail.CreateLogoffRecord(objDB, objDB.Profile.ID, Product)

                    Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cUSERPROFILE)
                    Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
                        objTable, objDB.Profile.ID)

                    colMasks.UpdateMaskObj(clsDBConstants.Fields.UserProfile.cLASTSIGNOFF, objDB.GetCurrentTime)

                    colMasks.Update(objDB, False)
                ElseIf objDB.Profile IsNot Nothing Then
                    Diagnostics.Trace.WriteLine("No Profile, cannot create log off record.")
                End If

                If objDB.Session IsNot Nothing AndAlso objDB.Profile IsNot Nothing Then

                    Dim strTempFolder As String
                    strTempFolder = ProperPath(My.Computer.FileSystem.SpecialDirectories.Temp) & "K1WebClient\Temp\" & objDB.Profile.ID & "\" & objDB.Session.ID & "\"

                    DeleteFolder(strTempFolder)

                    'If IO.Directory.Exists(strTempFolder) Then
                    '    IO.Directory.Delete(strTempFolder, True)
                    'End If

                    strTempFolder = ProperPath(HttpRuntime.AppDomainAppPath) & "Images\Temp\" & objDB.Profile.ID & "\" & objDB.Session.ID & "\"

                    DeleteFolder(strTempFolder)

                    'If IO.Directory.Exists(strTempFolder) Then
                    '    IO.Directory.Delete(strTempFolder, True)
                    'End If

                End If
            Catch ex As Exception

            End Try

            objDB.SessionStop()
            objDB.Dispose()

            m_colDatabase.Remove(objSession.SessionID)
        End If
    End Sub

    Public Sub DeleteFolder(strPath As String, Optional blnLastAttempt As Boolean = False)

        If (Not Directory.Exists(strPath)) Then
            Return
        End If

        Dim files As String() = Directory.GetFiles(strPath)

        For Each strFile As String In files
            File.SetAttributes(strFile, FileAttributes.Normal)
            File.Delete(strFile)
        Next

        Dim dirs As String() = Directory.GetDirectories(strPath)

        For Each dir As String In dirs
            DeleteFolder(dir)
        Next

        Try
            Directory.Delete(strPath, False)
        Catch ex As Exception
            If (blnLastAttempt) Then
                Return
            End If
            Thread.Sleep(200)
            DeleteFolder(strPath, True)
        End Try

    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                MyBase.Dispose(blnDisposing)

                If m_colDatabase IsNot Nothing Then
                    For Each objDB As clsDB In m_colDatabase.Values
                        objDB.Dispose()
                    Next
                    m_colDatabase.Clear()
                End If

                If m_objDB IsNot Nothing Then
                    m_objDB.Dispose()
                    m_objDB = Nothing
                End If
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Overrides Sub CreateStoredProcedure(sbProcName As String, sbProcedure As String)
        DB.CreateStoredProcedure(sbProcName,sbProcedure) 
    End Sub

    Public Overrides Function CheckStoredProcedureExists(sbProcName As String) As Boolean
        Return DB.CheckStoredProcedureExists (sbProcName)
    End Function
#End Region

End Class
