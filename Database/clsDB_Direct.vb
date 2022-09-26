Imports System.Data.SqlClient

Public Class clsDB_Direct
    Inherits clsDB

#Region " Members "

    'Variables related to database connection
    Protected m_strConnectionString As String
    Protected m_intCommandTimeout As Integer = 300
    Protected m_strApplicationName As String

    'Variables used to hold connection and transaction information
    Protected m_objDBTransaction As clsDBTransaction

    'Variables used for the training version of the database
    Protected m_blnTrainingVersion As Boolean = False

    Protected m_strIsysIndexPath As String
    Protected m_colTransactions As New FrameworkCollections.K1Dictionary(Of clsDBTransaction)
    Protected m_dtLastActivity As Date = Date.MinValue
    Protected m_dbSystem As clsDB_System = Nothing
#End Region

#Region " Enumerations "

    Public Enum enumSPType
        Insert = 0
        Update = 1
        Delete = 2
        GetItem = 3
        GetList = 4
    End Enum

    Public Enum enumSavedSearchType
        FILTER = 1
        ADVANCED = 2
        [BOOLEAN] = 3
    End Enum

    Public Enum enumSQLExceptions
        NO_SUCH_FIELD = 1
        NO_SUCH_TABLE = 2
        RECORD_LOCKED = 3
        INDEX_VIOLATION = 4
        FOREIGN_LINK_VIOLATION = 5
        USER_EXCEPTION = 6
        '2015-09-02 -- Peter Melisi -- Added new exception for warnings that need failure - USER_EXCEPTION_FAILURE
        USER_EXCEPTION_FAILURE = 7
    End Enum
#End Region

#Region " Constructors "


    ''' <summary>
    ''' Creates a new database object using the connection string provided
    ''' </summary>
    Public Sub New(ByVal strConnection As String)
        m_strConnectionString = strConnection
        m_eDataAccess = enumDataAccessType.DB_DIRECT

        m_objSysInfo = New clsSysInfo(Me)

        '[Naing Begin] Fix for Bug 1300002478
        TryInitializeSqlDependency()

    End Sub

    ''' <summary>
    ''' Creates a new database object using the connection string.
    ''' Limits SqlDependency initialisation
    ''' </summary>
    Public Sub New(ByVal strConnection As String, ByVal eAppType As clsDBConstants.enumApplicationType)
        m_strConnectionString = strConnection
        m_eDataAccess = enumDataAccessType.DB_DIRECT

        m_objSysInfo = New clsSysInfo(Me)

        ' Ara Melkonian - 2100003733
        ' Web applications and services should have multithreaded set to false.
        ' This fixes an issue where transaction state is incorrect.
        Select Case eAppType
            Case clsDBConstants.enumApplicationType.RecFindActivationKey, clsDBConstants.enumApplicationType.K1ActivationKey, clsDBConstants.enumApplicationType.K1,
                 clsDBConstants.enumApplicationType.RecFind, clsDBConstants.enumApplicationType.Tacit, clsDBConstants.enumApplicationType.Button,
                 clsDBConstants.enumApplicationType.Scan, clsDBConstants.enumApplicationType.Mini_API, clsDBConstants.enumApplicationType.GEM,
                 clsDBConstants.enumApplicationType.RecCapture, clsDBConstants.enumApplicationType.WebClient, clsDBConstants.enumApplicationType.SharePoint,
                 clsDBConstants.enumApplicationType.OneilIntegration, clsDBConstants.enumApplicationType.Archive
                ' Use the default
                Exit Select
            Case clsDBConstants.enumApplicationType.API
                MultiThreaded = False
                Exit Select
            Case Else
                ' Use the default
                Exit Select
        End Select

        If eAppType <> clsDBConstants.enumApplicationType.API _
        AndAlso eAppType <> clsDBConstants.enumApplicationType.RF6Connector Then
            TryInitializeSqlDependency()
        End If
    End Sub

    ''' <summary>
    ''' Creates a new database object using the connection string provided
    ''' </summary>
    Public Sub New(ByVal strConnection As String,
                   ByVal strApplicationName As String)

        Dim objConnectionBuilder As New SqlConnectionStringBuilder(strConnection)

        m_strApplicationName = strApplicationName
        objConnectionBuilder.ApplicationName = "RecFind" '-- Always recfind because VERS expects it

        m_strConnectionString = objConnectionBuilder.ConnectionString
        m_eDataAccess = enumDataAccessType.DB_DIRECT

        m_objSysInfo = New clsSysInfo(Me)

        '[Naing Begin] Fix for Bug 1300002478
        TryInitializeSqlDependency()

    End Sub

    '[Naing] Fix for Bug 1300002478
    '[Naing] This field is passed into the constructor from the sub class clsDB_Web to ensure that only 
    'the Main instance of clsDB_Direct will start SqlDependency once and only once and when that main 
    'instance is disposed then it will call SqlDependency stop once and only once.
    Private m_blnUseSqlDependency As Boolean = True

    ''' <summary>
    ''' Creates a new database object using the connection string provided and
    ''' using a shared Sys Info object
    ''' </summary>
    Public Sub New(ByVal strConnection As String,
                   ByVal objSysInfo As clsSysInfo,
                   Optional blnUseSqlDependency As Boolean = True)

        MyBase.New()

        m_strConnectionString = strConnection
        m_eDataAccess = enumDataAccessType.DB_DIRECT

        m_objSysInfo = objSysInfo

        '[Naing Begin] Fix for Bug 1300002478
        If (blnUseSqlDependency) Then
            TryInitializeSqlDependency()
        End If

        m_blnUseSqlDependency = blnUseSqlDependency

    End Sub

    Public Overridable Function TryInitializeSqlDependency() As Boolean

        Try
            '-- Make sure we have a sql dependency object
            If SqlDependency Is Nothing Then
                SqlDependency = New clsSqlDependency(Me)
            Else
                If SqlDependency.SystemFlags IsNot Nothing Then
                    SqlDependency.SystemFlags.Refresh(Me)
                End If
            End If

            '-- Make sure notification broker is started
            If Not SqlDependency.Started Then
                SqlDependency.Start()
            End If

        Catch ex As Exception
            '[Naing] suppress exception here for now, this code above will be run again at a later stage from another location.
            'I know same code in different locations breaks the D.R.Y rule but I didn't write it, I am just maintaining it.
            If (SqlDependency IsNot Nothing) Then
                SqlDependency.Dispose()
                SqlDependency = Nothing
            End If

        End Try

        Return SqlDependency IsNot Nothing

    End Function

#End Region

#Region " Properties "

    ''' <summary>
    ''' This contains the database connection string (encrypted)
    ''' </summary>
    Public ReadOnly Property ConnectionString() As String
        Get
            Return m_strConnectionString
        End Get
    End Property

    ''' <summary>
    ''' The time in seconds before a database command times out
    ''' </summary>
    Public Property CommandTimeout() As Integer
        Get
            Return m_intCommandTimeout
        End Get
        Set(ByVal value As Integer)
            m_intCommandTimeout = value
        End Set
    End Property

    ''' <summary>
    ''' This is the path to the Isys Index used when performing Isys searches
    ''' </summary>
    Public ReadOnly Property IsysIndexPath() As String
        Get
            Return m_strIsysIndexPath
        End Get
    End Property

    Public Overrides ReadOnly Property HasTransaction() As Boolean
        Get
            If m_blnMultiThreaded Then
                Dim intThreadID As Integer = Threading.Thread.CurrentThread.ManagedThreadId

                Return Not (m_colTransactions(CStr(intThreadID)) Is Nothing)
            Else
                Return Not (m_objDBTransaction Is Nothing)
            End If
        End Get
    End Property

    Public ReadOnly Property SystemFlags() As clsSystemFlags
        Get
            If Me.m_objSQLDependency IsNot Nothing Then
                Return Me.m_objSQLDependency.SystemFlags
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' When using a direct connection, this should be true and is a way of keeping transaction threaded
    ''' When created through web services, this should be false, as each web service call would be on a separate thread
    ''' </summary>
    Public Property MultiThreaded() As Boolean
        Get
            Return m_blnMultiThreaded
        End Get
        Set(ByVal value As Boolean)
            m_blnMultiThreaded = value
        End Set
    End Property

    Public Overrides Property Profile() As clsUserProfile
        Get
            Return m_objProfile
        End Get
        Set(ByVal value As clsUserProfile)
            If m_objSQLDependency IsNot Nothing Then
                m_objSQLDependency.Dispose()
                m_objSQLDependency = Nothing
            End If

            If Not String.IsNullOrEmpty(m_strConnectionString) AndAlso
            Not String.IsNullOrEmpty(m_strApplicationName) Then
                If value IsNot Nothing Then
                    Dim objBuilder As New SqlConnectionStringBuilder(m_strConnectionString)
                    objBuilder.ApplicationName = m_strApplicationName & " : " & value.ID
                    m_strConnectionString = objBuilder.ToString
                Else
                    Dim objBuilder As New SqlConnectionStringBuilder(m_strConnectionString)
                    objBuilder.ApplicationName = m_strApplicationName
                    m_strConnectionString = objBuilder.ToString
                End If
            End If

            m_objProfile = value
        End Set
    End Property

    Public Property LastActivity As Date
        Get
            Return m_dtLastActivity
        End Get
        Set(ByVal value As Date)
            m_dtLastActivity = value
        End Set
    End Property

    Private Function dbSystem() As clsDB_System
        If Me.GetType() Is GetType(clsDB_System) Then
            Return CType(Me, clsDB_System)
        End If

        If m_dbSystem Is Nothing Then
            m_dbSystem = New clsDB_System(Me.ConnectionString)
        End If
        Return m_dbSystem
    End Function

#End Region

#Region " Methods "

#Region " Connections and Transactions "

#Region " Connections "

    Protected Function GetDBTransaction() As clsDBTransaction
        If m_blnMultiThreaded AndAlso Not m_colTransactions Is Nothing Then
            Dim intThreadID As Integer = Threading.Thread.CurrentThread.ManagedThreadId

            Return m_colTransactions(CStr(intThreadID))
        Else
            Return m_objDBTransaction
        End If
    End Function

    ''' <summary>
    ''' Creates or assigns already open connection if one exists
    ''' </summary>
    Protected Function GetConnection(Optional ByVal blnOpen As Boolean = True) As SqlConnection

        Dim conDB As SqlConnection

        Dim objDBTransaction As clsDBTransaction = GetDBTransaction()

        If objDBTransaction Is Nothing Then
            conDB = New SqlConnection(m_strConnectionString)
        Else
            conDB = objDBTransaction.Connection
        End If

        If blnOpen AndAlso conDB.State = ConnectionState.Closed Then
            OpenConnectionWithContext(conDB)
        End If

        Return conDB

    End Function

    ''' <summary>
    ''' Closes an open connection, if there is no current transaction
    ''' </summary>
    ''' <param name="conDB"></param>
    ''' <remarks></remarks>
    Protected Sub CloseConnection(ByRef conDB As SqlConnection)
        'only close the connection if it is not attached to a transaction.
        Dim objDBTransaction As clsDBTransaction = GetDBTransaction()
        If Not conDB Is Nothing AndAlso objDBTransaction Is Nothing Then
            If conDB.State = ConnectionState.Open Then
                conDB.Close()
            End If

            '[Naing] Why is this here? This code is messing with connection pooling and causing a lot of problems in web deployments.
            'Why is this code necessary? Why interfere with normal connection pooling of the framework.
            'If Not m_blnMultiThreaded Then
            '    SqlConnection.ClearPool(conDB)
            'End If

            conDB.Dispose()
            conDB = Nothing
        End If

    End Sub

    ''' <summary>
    ''' Tries to open a database connection.  
    ''' This command should be encapsulated in a try... catch block
    ''' </summary>
    Public Shared Sub TestConnection(ByVal strConnection As String)
        Dim objConnection As New SqlConnection(strConnection)
        objConnection.Open()
        objConnection.Close()
    End Sub

    ''' <summary>
    ''' Returns a valid RecFind connection string given the parameters supplied.
    ''' </summary>
    Public Shared Function GetConnectionString(ByVal strServer As String, ByVal strDatabase As String,
    ByVal strUserName As String, ByVal strPassword As String) As String
        'Dim strConnection As String

        'strConnection = "Server=" & strServer & ";"
        'strConnection &= "Database=" & strDatabase & ";"
        'strConnection &= "user id=" & strUserName & ";"
        'strConnection &= "password=" & strPassword & ";"
        'strConnection &= "Persist Security Info=true;"

        'Return strConnection

        Return GetConnectionString("RecFind 6", strServer, strDatabase, strUserName, strPassword)
    End Function

    ''' <summary>
    ''' Returns a valid Recfind connection string given the parameters supplied.
    ''' </summary>
    Public Shared Function GetConnectionString(ByVal strAppName As String, ByVal strServer As String,
                                               ByVal strDatabase As String, ByVal strUserName As String,
                                               ByVal strPassword As String) As String
        Dim objConnectionBuilder As New SqlClient.SqlConnectionStringBuilder()

        objConnectionBuilder.DataSource = strServer
        objConnectionBuilder.InitialCatalog = strDatabase
        objConnectionBuilder.UserID = strUserName
        objConnectionBuilder.Password = strPassword
        objConnectionBuilder.ApplicationName = strAppName
        objConnectionBuilder.PersistSecurityInfo = True

        Return objConnectionBuilder.ConnectionString
    End Function

    ''' <summary>
    ''' Returns insensitive information regarding the current connection
    ''' </summary>
    Public Overrides Sub GetDatabaseInfo(ByRef strServer As String, ByRef strDatabase As String,
    ByRef strUserID As String, Optional ByVal dblVersion As Double = 0,
    Optional ByVal strAppName As String = Nothing,
    Optional ByVal dblMinDBVersion As Double = 0)
        Dim objBuilder As New SqlConnectionStringBuilder(m_strConnectionString)
        strServer = objBuilder.DataSource
        strDatabase = objBuilder.InitialCatalog
        strUserID = objBuilder.UserID

    End Sub

    Public Overrides Sub GetDatabaseGroupInfo(ByRef strGroup As String, ByRef strServer As String, ByRef strDatabase As String, ByRef strUserID As String, Optional dblVersion As Double = 0, Optional strAppName As String = Nothing, Optional dblMinDBVersion As Double = 0)
        GetDatabaseInfo(strServer, strDatabase, strUserID, dblVersion, strAppName, dblMinDBVersion)
        strGroup = GetAvailabilityGroupName()
    End Sub

    Public Overrides Function CheckHostName(ByVal strServer As String) As String
        Return GetHostName(strServer)
    End Function

    Public Overrides Function CheckSQLServerName() As String
        Return GetSQLServerName(Me)
    End Function
#End Region

#Region " Transactions "

#Region " Get Transaction "

    ''' <summary>
    ''' Returns a transaction object if one exists
    ''' </summary>
    Protected Function GetTransaction() As SqlTransaction
        Dim objDBTransaction As clsDBTransaction = GetDBTransaction()

        If objDBTransaction Is Nothing Then
            Return Nothing
        Else
            Return objDBTransaction.Transaction
        End If
    End Function
#End Region

#Region " Begin Transaction "

    ''' <summary>
    ''' Starts a database transaction with a default IsolationLevel of Snapshot
    ''' </summary>
    ''' <remarks>
    ''' With an Isolation Level of Snapshot, there is no shared locks on updating and reading.
    ''' If someone reads a record while somone else is updating it in a transaction, they will
    ''' get the record as it existed prior to the transaction.
    ''' </remarks>
    Public Overloads Overrides Sub BeginTransaction()
        Me.BeginTransaction(IsolationLevel.ReadCommitted)
    End Sub

    ''' <summary>
    ''' Creates a new transaction for the connection using the isolation type specified
    ''' </summary>
    Public Overloads Overrides Sub BeginTransaction(ByVal eIsolationLevel As System.Data.IsolationLevel)
        Try
            Dim objDBTransaction As clsDBTransaction = GetDBTransaction()

            If objDBTransaction IsNot Nothing Then
                Throw New Exception("A transaction is already open.")
            End If

            Dim objConnection As New SqlConnection(m_strConnectionString)
            '[Naing] Let's set the connection context here. This should become an extension method maybe?
            OpenConnectionWithContext(objConnection)
            'objConnection.Open()

            Dim objTransaction As SqlTransaction
            Try
                objTransaction = objConnection.BeginTransaction(eIsolationLevel, "RecFind")
            Catch ex As SqlException When ex.Number = 3903
                objTransaction = objConnection.BeginTransaction(eIsolationLevel, "RecFind")
            End Try

            objDBTransaction = New clsDBTransaction(objConnection, objTransaction)

            If m_blnMultiThreaded Then
                Dim intThreadID As Integer = Threading.Thread.CurrentThread.ManagedThreadId
                m_colTransactions.Add(CStr(intThreadID), objDBTransaction)
            Else
                m_objDBTransaction = objDBTransaction
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' This is to ensure that the db server has some context of the client connection.
    ''' </summary>
    ''' <param name="objConnection"></param>
    ''' <remarks>Adds the current user profile id.</remarks>
    Private Sub OpenConnectionWithContext(objConnection As IDbConnection)

        If (Not objConnection.State = ConnectionState.Open) Then
            objConnection.Open()
        End If

        If (Profile Is Nothing) Then
            Return
        End If

        Const paramName As String = "@SessionInfo"

        Dim strSqlCmd = String.Format("declare @Length tinyint " + Environment.NewLine +
                                   "declare @Ctx varbinary(128) select @Length = len({0})" + Environment.NewLine +
                                   "select @Ctx = convert(binary(1), @Length) + convert(varbinary(127), {0})" + Environment.NewLine +
                                   "set context_info @Ctx", paramName)

        Dim cmd As IDbCommand = objConnection.CreateCommand()
        Using (cmd)
            cmd.CommandType = CommandType.Text
            cmd.CommandText = strSqlCmd
            Dim param = cmd.CreateParameter()
            param.DbType = DbType.AnsiString
            param.ParameterName = paramName
            param.Size = 127
            param.Value = Profile.ID
            'for future use
            'param.Value = Profile.ID + "," + My.Application.Info.ProductName
            cmd.Parameters.Add(param)
            cmd.ExecuteNonQuery()
        End Using

    End Sub

#End Region

#Region " End Transaction "

    ''' <summary>
    ''' Ends any transactions which exist on open connections
    ''' </summary>
    ''' <param name="blnCommit">True - Commit transaction, False - Rollback transaction</param>
    Public Overrides Sub EndTransaction(ByVal blnCommit As Boolean)
        Try
            Dim objDBTransaction As clsDBTransaction = GetDBTransaction()

            If Not objDBTransaction Is Nothing Then
                If Not objDBTransaction.Transaction Is Nothing Then
                    Try
                        If blnCommit Then
                            objDBTransaction.Commit()
                        Else
                            objDBTransaction.Rollback()
                        End If
                    Catch ex As Exception
                    End Try
                End If

                If Not objDBTransaction.Connection Is Nothing Then
                    Try
                        If objDBTransaction.Connection.State = ConnectionState.Open Then
                            objDBTransaction.Connection.Close()
                        End If
                    Catch ex As Exception

                    End Try
                End If

                objDBTransaction.Dispose()

                If m_blnMultiThreaded Then
                    Dim intThreadID As Integer = Threading.Thread.CurrentThread.ManagedThreadId
                    m_colTransactions.Remove(CStr(intThreadID))
                Else
                    m_objDBTransaction = Nothing
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#End Region

#End Region

#Region " GetItem, GetList Methods "

    ''' <summary>
    ''' Gets the record with the given ID from the specified table
    ''' </summary>
    Public Overrides Function GetItem(ByVal strTable As String, ByVal intID As Integer) As DataTable
        Dim objDT As DataTable
        Dim colParams As clsDBParameterDictionary
        Dim strStoredProcedure As String

        Try
            strStoredProcedure = strTable & clsDBConstants.StoredProcedures.cGETITEM

            colParams = GetParamCollection(enumSPType.GetItem, strTable, clsDBConstants.Fields.cID, intID)

            objDT = GetDataTable(strStoredProcedure, colParams)

            Return objDT
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Gets a list of records from the specified table where strByField matches objValue
    ''' </summary>
    Public Overrides Function GetList(ByVal strTable As String, ByVal strByField As String,
    ByVal objValue As Object) As DataTable
        Dim objDT As DataTable
        Dim strStoredProcedure As String
        Dim colParams As clsDBParameterDictionary

        Try
            strStoredProcedure = strTable & clsDBConstants.StoredProcedures.cGETLIST

            colParams = GetParamCollection(enumSPType.GetList, strTable, strByField, objValue)

            objDT = Me.GetDataTable(strStoredProcedure, colParams)

            Return objDT
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " GetDataTable Methods "

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strCommandText"></param>
    ''' <param name="eCommandType"></param>
    ''' <param name="colParams"></param>
    ''' <param name="objCallBack"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function FillDataTable(ByVal strCommandText As String,
                                     ByVal eCommandType As CommandType,
                                     ByVal colParams As clsDBParameterDictionary,
                                     ByVal objCallBack As OnChangeEventHandler,
                                     Optional ByVal blnValidateType As Boolean = False) As DataTable

        Dim conDB As SqlConnection
        Dim objDT As DataTable
        'get or Create the Connection            
        conDB = GetConnection()

        Try

#If TRACE Then
            Trace.WriteVerbose("(FillDataTable) " & strCommandText.Replace(vbCr, "").Replace(vbLf, " "), "")
#End If
            Dim cmdDB As SqlCommand = conDB.CreateCommand()
            cmdDB.CommandText = strCommandText
            cmdDB.CommandType = eCommandType
            cmdDB.CommandTimeout = m_intCommandTimeout
            If colParams IsNot Nothing AndAlso colParams.Any() Then
                For Each objParam As clsDBParameter In colParams.Values
                    If Not blnValidateType Then
                        cmdDB.Parameters.AddWithValue(objParam.Name, objParam.Value)
                    Else
                        cmdDB.Parameters.Add(objParam.Name, objParam.SqlDBType).Value = objParam.Value
                    End If
                Next
            End If

            'join a SQL transaction if it exists
            Dim currentTransaction = GetTransaction()
            If (currentTransaction IsNot Nothing) Then
                cmdDB.Transaction = currentTransaction
            End If

            If objCallBack IsNot Nothing Then
                '--  Clear any existing notifications
                cmdDB.Notification = Nothing

                '--  Create the dependency for this command
                Dim objSqlDependency As SqlDependency = New SqlDependency(cmdDB)

#If TRACE Then
                Trace.WriteVerbose("(FillDataTable) Id - " & objSqlDependency.Id & " with " & objCallBack.Method.Name, "")
#End If
                AddHandler objSqlDependency.OnChange, objCallBack
            End If

            Using objAdapter As New SqlDataAdapter() With {.SelectCommand = cmdDB}
                objDT = New DataTable("DT")
                objAdapter.Fill(objDT)
            End Using
            Return objDT
        Catch ex As SqlException
            'this throws exceptions so no need to worry about returning nothing
            ProcessSQLExceptionExecute(ex)
            Return Nothing
        Catch ex As Exception
            Throw
        Finally
            CloseConnection(conDB)
        End Try

    End Function

    ''' <summary>
    ''' Returns all records for the specified table
    ''' </summary>
    Public Overloads Overrides Function GetDataTable(ByVal objTable As clsTable) As DataTable
        Dim strStoredProcedure As String

        strStoredProcedure = objTable.DatabaseName & clsDBConstants.StoredProcedures.cGETLIST

        Return GetDataTable(strStoredProcedure, Nothing)
    End Function

    ''' <summary>
    ''' Returns all records selected in the stored procedure
    ''' </summary>
    Public Overloads Overrides Function GetDataTable(ByVal strStoredProcedure As String) As DataTable
        Return GetDataTable(strStoredProcedure, Nothing)
    End Function

    ''' <summary>
    ''' Returns all records selected in the stored procedure matching the designated SP's parameters
    ''' </summary>
    Public Overloads Overrides Function GetDataTable(ByVal strStoredProcedure As String,
    ByVal colParams As clsDBParameterDictionary) As DataTable
        Try
            Return FillDataTable(strStoredProcedure, CommandType.StoredProcedure, colParams, Nothing)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement
    ''' </summary>
    Public Overrides Function GetDataTableBySQL(ByVal strSQL As String) As DataTable
        Try
            Return GetDataTableBySQL(strSQL, Nothing)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement with the designated parameters
    ''' </summary>
    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQL As String,
                                                          ByVal colParams As clsDBParameterDictionary) As DataTable
        Try
            Return FillDataTable(strSQL, CommandType.Text, colParams, Nothing)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQL As String,
                                                          ByVal colParams As clsDBParameterDictionary,
                                                          ByVal blnValidateType As Boolean) As DataTable
        Try
            Return FillDataTable(strSQL, CommandType.Text, colParams, Nothing, blnValidateType)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement with the designated parameters.
    ''' Replaces all the format items with the string equivalent value in the corresponding array.
    ''' </summary>
    Public Overloads Overrides Function GetDataTableBySQL(ByVal strSQLFormat As String,
    ByVal colParams As clsDBParameterDictionary, ByVal ParamArray args() As Object) As DataTable
        Try
            Dim strSQL As String = strSQLFormat
            If args.GetUpperBound(0) > 0 Then
                strSQL = String.Format(strSQLFormat, args)
            End If

            Return GetDataTableBySQL(strSQL, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Returns all records for the specified table where strField matches objValue
    ''' </summary>
    Public Overrides Function GetDataTableByField(ByVal strTable As String,
    ByVal strField As String, ByVal objValue As Object) As DataTable
        Dim colParams As New clsDBParameterDictionary

        colParams = GetParamCollection(enumSPType.GetList, strTable, strField, objValue)

        Return GetDataTableByField(strTable, colParams)
    End Function

    ''' <summary>
    ''' Returns all records for the specified table where strField matches objValue
    ''' </summary>
    Public Overrides Function GetDataTableByField(ByVal strTable As String,
    ByVal colParams As clsDBParameterDictionary) As DataTable
        Dim objDataTable As DataTable = GetDataTable(
            strTable & clsDBConstants.StoredProcedures.cGETLIST, colParams)

        Return objDataTable
    End Function

#End Region

#Region " GetParamCollection "

    ''' <summary>
    ''' Returns a collection of parameters with values for a particular 
    ''' stored procedure for a table (single parameter)
    ''' </summary>
    Private Function GetParamCollection(ByVal eSPType As enumSPType, ByVal strTable As String,
    ByVal strFieldName As String, ByVal objValue As Object) As clsDBParameterDictionary
        Dim arrFieldName() As String = {strFieldName}
        Dim arrValue() As Object = {objValue}

        Return GetParamCollection(m_objSysInfo, eSPType, strTable, arrFieldName, arrValue)
    End Function

    ''' <summary>
    ''' Returns a collection of parameters with values for a particular 
    ''' stored procedure for a table (multiple parameters)
    ''' </summary>
    Public Shared Function GetParamCollection(ByVal objSysInfo As clsSysInfo,
    ByVal eSPType As enumSPType, ByVal strTable As String,
    ByVal arrFieldName() As String, ByVal arrValue() As Object) As clsDBParameterDictionary
        Dim colParams As New clsDBParameterDictionary

        'Find the Table from the Database
        Dim objTable As clsTable = objSysInfo.Tables(strTable)
        If objTable Is Nothing Then
            'TODO: Get from table ErrorMessage?
            Throw New Exception("Database Error: Table '" & strTable & "' not found.")
        End If

        For intLoop As Integer = 0 To arrFieldName.GetUpperBound(0)
            'Find the Field from the Database
            Dim objFld As clsField = objSysInfo.Fields(objTable.ID & "_" & arrFieldName(intLoop))
            If objFld Is Nothing Then
                'TODO: Get from table ErrorMessage
                Throw New Exception("Database Error: Field  '" & arrFieldName(intLoop) & "' not found in table '" & objTable.DatabaseName & "'.")
            End If

            AddParameterToCollection(colParams, eSPType, objFld, arrValue(intLoop))
        Next

        Return colParams
    End Function

    ''' <summary>
    ''' Returns a collection of parameters with values for a particular mask collection
    ''' </summary>
    Public Shared Function GetParamCollection(ByVal eSPType As enumSPType,
    ByVal colMasks As clsMaskFieldDictionary) As clsDBParameterDictionary
        Dim colParams As New clsDBParameterDictionary
        Dim objMask As clsMaskField

        If eSPType = enumSPType.Delete Then
            objMask = colMasks(clsDBConstants.Fields.cID)
            colParams.Add(New clsDBParameter(ParamName(objMask.Field.DatabaseName),
                objMask.Value1.Value, ParameterDirection.Input, objMask.Field.DataType))
        Else
            For Each objMask In colMasks.Values
                If Not (eSPType = enumSPType.Update AndAlso objMask.Field.IsBinaryType) Then
                    AddParameterToCollection(colParams, eSPType, objMask.Field, objMask.Value1.Value)
                End If
            Next
        End If

        Return colParams
    End Function

    ''' <summary>
    ''' Adds a parameter to the parameter collection
    ''' </summary>
    ''' <remarks>
    ''' The ID Field will be designated as inputoutput for the insert stored procedure type
    ''' </remarks>
    Private Shared Sub AddParameterToCollection(ByRef colParams As clsDBParameterDictionary,
    ByVal eSPType As enumSPType, ByVal objField As clsField, ByVal objValue As Object)
        'Create Our new Database Parameter, and add it to the collection
        Dim objDBIParam As clsDBParameter

        If eSPType = enumSPType.Insert AndAlso objField.IsIdentityField Then
            objDBIParam = New clsDBParameter(ParamName(objField.DatabaseName),
                objValue, ParameterDirection.InputOutput, objField.DataType)
        Else
            objDBIParam = New clsDBParameter(ParamName(objField.DatabaseName),
                objValue, ParameterDirection.Input, objField.DataType)
        End If

        colParams.Add(objDBIParam)
    End Sub
#End Region

#Region " ExecuteProcedure, ExecuteSQL, ExecuteScalar, DeleteRecord "

    ''' <summary>
    ''' Executes the specified stored procedure or query
    ''' </summary>
    Protected Function ExecuteNonQuery(ByVal strCommandText As String,
                                       ByVal eCommandType As CommandType,
                                       ByRef colParams As clsDBParameterDictionary) As Integer
        Dim conDB As SqlConnection = GetConnection(True)
        Dim intRows As Integer = 0

        Try
#If TRACE Then
            Trace.WriteVerbose("(ExecuteNonQuery) " & strCommandText.Replace(vbCr, "").Replace(vbLf, " "), "")
#End If
            Dim cmdDB As New SqlCommand(strCommandText, conDB)
            cmdDB.CommandType = eCommandType
            cmdDB.CommandTimeout = m_intCommandTimeout

            If Not colParams Is Nothing Then
                For Each objParam As clsDBParameter In colParams.Values

                    Dim objSqlParam As SqlParameter = cmdDB.Parameters.AddWithValue(objParam.Name, objParam.Value)
                    Dim eType As DbType = objParam.DBType
                    If Not eType = clsDBConstants.cintNULL Then
                        objSqlParam.DbType = eType
                        If objSqlParam.DbType = DbType.Binary Then
                            objSqlParam.Size = 100
                        End If
                    End If
                    objSqlParam.Direction = objParam.Direction
                Next
            End If

            cmdDB.Transaction = GetTransaction()
            intRows = cmdDB.ExecuteNonQuery()

            If Not colParams Is Nothing Then
                For Each objParam As clsDBParameter In colParams.Values
                    If objParam.Direction <> ParameterDirection.Input Then
                        objParam.Value = cmdDB.Parameters(objParam.Name).Value
                    End If
                Next
            End If
        Catch ex As SqlException
            ProcessSQLExceptionExecute(ex)
        Catch ex As Exception
            Throw
        Finally
            CloseConnection(conDB)
        End Try

        Return intRows

    End Function

    ''' <summary>
    ''' Executes the specified stored procedure
    ''' </summary>
    Public Overloads Overrides Sub ExecuteProcedure(ByVal strStoredProcedure As String)
        ExecuteProcedure(strStoredProcedure, Nothing)
    End Sub

    ''' <summary>
    ''' Executes the specified stored procedure with the designated SP Parameters
    ''' </summary>
    Public Overloads Overrides Sub ExecuteProcedure(ByVal strStoredProcedure As String,
    ByRef colParams As clsDBParameterDictionary)
        ExecuteNonQuery(strStoredProcedure, CommandType.StoredProcedure, colParams)
    End Sub

    ''' <summary>
    ''' Executes the specified SQL statement
    ''' </summary>
    Public Overloads Overrides Function ExecuteSQL(ByVal strSQL As String) As Integer
        Return ExecuteSQL(strSQL, Nothing)
    End Function

    ''' <summary>
    ''' Executes the specified SQL statement with the designated Parameters
    ''' </summary>
    Public Overloads Overrides Function ExecuteSQL(ByVal strSQL As String,
    ByRef colParams As clsDBParameterDictionary) As Integer
        Return ExecuteNonQuery(strSQL, CommandType.Text, colParams)
    End Function

    ''' <summary>
    ''' Executes the specified SQL SELECT statement with the designated Parameters, and returns first column (should be integer)
    ''' </summary>
    Public Overrides Function ExecuteScalar(ByVal strSQL As String,
        ByRef colParams As clsDBParameterDictionary) As Integer
        Dim conDB As SqlConnection = GetConnection(True)
        Dim intResult As Integer = 0
        Try

#If TRACE Then
            Trace.WriteVerbose("(ExecuteScalar) " & strSQL.Replace(vbCr, "").Replace(vbLf, " "), "")
            If colParams IsNot Nothing Then
                For Each param In colParams
                    Trace.WriteVerbose($"        {param.Key}={param.Value.Value}", "")
                Next
            End If
#End If
            Dim cmdDB As New SqlCommand(strSQL, conDB)
            cmdDB.CommandTimeout = m_intCommandTimeout

            If Not colParams Is Nothing Then
                For Each objParam As clsDBParameter In colParams.Values
                    Dim objNewParam = cmdDB.Parameters.AddWithValue(objParam.Name, objParam.Value)

                    If objParam.DBType = DbType.Binary Then
                        objNewParam.Size = CInt(objParam.Value)
                    End If
                Next
            End If

            cmdDB.Transaction = GetTransaction()

            intResult = CInt(cmdDB.ExecuteScalar())

            Return intResult
        Catch ex As SqlException
            ProcessSQLExceptionExecute(ex)
        Catch ex As Exception
            Throw
        Finally
            CloseConnection(conDB)
        End Try

        Return intResult
    End Function

    ''' <summary>
    ''' Executes the specified SQL SELECT statement and returns first column (should be integer)
    ''' </summary>
    Public Overrides Function ExecuteScalar(ByVal strSQL As String) As Integer
        Return ExecuteScalar(strSQL, Nothing)
    End Function

    ''' <summary>
    ''' Deletes the record with the designated ID from the specified table
    ''' </summary>
    Public Overrides Sub DeleteRecord(ByVal strTable As String, ByVal intID As Integer)
        If strTable = clsDBConstants.Tables.cEDOC Then
            Dim intDB As Integer = ExecuteScalar("SELECT DBID FROM K1Archive_Link WHERE EDOCID=" & CStr(intID))

            If intDB > 0 Then
                Dim objArchiveDT As DataTable = GetDataTableBySQL("SELECT ExternalID FROM K1Archive WHERE ID=" & CStr(intDB))
                If objArchiveDT IsNot Nothing AndAlso objArchiveDT.Rows.Count = 1 Then
                    Dim objEncryption As New clsEncryption(True)
                    Dim objConnection As New SqlClient.SqlConnection(objEncryption.Decrypt(CStr(objArchiveDT.Rows(0)(0))))
                    objConnection.Open()
                    ExecuteSQL("DELETE FROM K1Archive_Link WHERE EDOCID=" & CStr(intID))

                    Dim cmdDB As SqlClient.SqlCommand = objConnection.CreateCommand()
                    cmdDB.CommandText = "DELETE FROM EDOC WHERE ID=" & CStr(intID)
                    cmdDB.CommandType = CommandType.Text
                    cmdDB.ExecuteNonQuery()
                End If

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objArchiveDT IsNot Nothing Then
                    objArchiveDT.Dispose()
                    objArchiveDT = Nothing
                End If
            End If
        End If

        Dim colParams As clsDBParameterDictionary

        colParams = GetParamCollection(enumSPType.Delete, strTable, clsDBConstants.Fields.cID, intID)

        ExecuteProcedure(strTable & clsDBConstants.StoredProcedures.cDELETE, colParams)
    End Sub
#End Region

#Region " Error Messages Interpretations "

    ''' <summary>
    ''' will throw different exception types depending on the error type
    ''' it is up to the parent routine to handle these exceptions differently (if needed)
    ''' </summary>
    ''' <exception cref=" clsK1Exception"></exception>
    ''' <exception cref="SqlException"></exception>
    Private Sub ProcessSQLExceptionExecute(ByVal ex As SqlException)
        Select Case ex.Number
            Case 207    ' A field does not exist (or has been deleted)
                Throw New clsK1Exception(enumSQLExceptions.NO_SUCH_FIELD, ex.Message, ex)
            Case 208    ' the table does not exist (or has been deleted)
                Throw New clsK1Exception(enumSQLExceptions.NO_SUCH_TABLE, ex.Message, ex)
            Case 1205   ' A database lock exists (from transaction locking)
                'TODO: Load from error messages
                Throw New clsK1Exception(enumSQLExceptions.RECORD_LOCKED,
                    "This record is in use and locked, please try again later.", ex)
            Case 2601   ' Violates an index (usually a unique value constraint)
                ThrowIndexError(ex)
            Case 547    ' There is a mandatory link to this record
                If ex.Message.StartsWith("The DELETE") Then
                    ThrowDeleteError(ex)
                Else
                    ThrowUpdateError(ex)
                End If

            Case 515
                ThrowForeignLinkViolation(ex)
            Case 50000  ' A user error from a trigger
                Dim strMessage As String

                If ex.Errors IsNot Nothing AndAlso ex.Errors.Count >= 1 Then
                    strMessage = ex.Errors(0).Message
                Else
                    strMessage = ex.Message
                End If
                '2015-09-02 -- Peter Melisi -- Pass in clsDB_Direct.enumSQLExceptions.USER_EXCEPTION_FAILURE exception 
                ' for severity level of 18.
                If ex.Class = 18 Then
                    Throw New clsK1Exception(Me, enumSQLExceptions.USER_EXCEPTION_FAILURE, strMessage)
                Else
                    Throw New clsK1Exception(Me, enumSQLExceptions.USER_EXCEPTION, strMessage)
                End If
            Case Else
                Throw ex
        End Select
    End Sub

    Private Sub ThrowForeignLinkViolation(ByVal exSql As SqlException)
        Dim arrParts As String() = exSql.Message.Replace("'"c, """"c).Split(""""c)

        If arrParts.Count > 2 Then
            Throw New clsK1Exception(enumSQLExceptions.FOREIGN_LINK_VIOLATION,
                "There is currently no value in the field '" & arrParts(1) & "'." & vbCrLf &
                "This is a mandatory field as designated by the system.", exSql)
        Else
            Throw New clsK1Exception(enumSQLExceptions.FOREIGN_LINK_VIOLATION,
                 "This is a mandatory field as designated by the system.", exSql)
        End If
    End Sub

    ''' <summary>
    ''' Handles unique index errors (creates a nice error message)
    ''' </summary>
    Private Sub ThrowIndexError(ByVal exSql As SqlException)
        Dim strNewException As String

        Try
            Dim strMessage As String = exSql.Message
            Dim arrParts As String() = strMessage.Split("'"c)

            If Not arrParts.Length = 5 Then
                Throw exSql
            End If

            Dim strTable As String = arrParts(1)
            Dim strIndex As String = arrParts(3)

            Dim objDT As DataTable = Me.GetDataTableBySQL("EXEC sp_helpindex [" & strTable & "]")
            objDT.DefaultView.RowFilter = "index_name = '" & SQLString(strIndex) & "'"

            If Not objDT.DefaultView.Count = 1 Then
                Throw exSql
            End If

            Dim intIndex As Integer = strTable.LastIndexOf(".")
            Dim strJustTable As String
            Dim objTable As clsTable
            Dim strCaption As String = strTable

            If intIndex = -1 Then
                strJustTable = strTable
            Else
                strJustTable = strTable.Substring(intIndex + 1, strTable.Length - (intIndex + 1))
            End If

            objTable = SysInfo.Tables(strJustTable)
            If objTable IsNot Nothing Then
                strCaption = objTable.CaptionText
            End If

            If strJustTable.ToUpper = "K1RECORDLOCK" Then
                strNewException = "Attempted to update a record which is locked by another user."
            Else
                Dim strKeys As String = TryCast(objDT.DefaultView.Item(0).Item("index_keys"), String)

                'TODO: load from error messages
                If strKeys.IndexOf(",") >= 0 Then
                    strNewException = "The fields '" & strKeys & "' must have unique values within table '" & strCaption & "'.  " &
                        "Another record already exists with the selected values. Please change the values of " &
                        "the fields '" & strKeys & "' to something unique before saving this record."
                Else
                    strNewException = "The field '" & strKeys & "' must have a unique value within table '" & strCaption & "'.  " &
                        "Another record already exists with the selected value. Please change the value of " &
                        "field '" & strKeys & "' to something unique before saving this record."
                End If
            End If
        Catch ex As Exception
            Throw exSql
        End Try

        Throw New clsK1Exception(enumSQLExceptions.INDEX_VIOLATION, strNewException)
    End Sub

    ''' <summary>
    ''' Handles mandatory link errors when trying to delete a record (creates a nice error message)
    ''' </summary>
    Private Sub ThrowUpdateError(ByVal exSql As SqlException)
        Dim strNewException As String = String.Empty
        Try
            Dim strMessage As String = exSql.Message
            Dim arrParts As String() = strMessage.Replace("'"c, """"c).Split(""""c)

            strNewException = "There is a field which is referencing a record which no longer exists." & vbCrLf &
                 "Please consult your RecFind Administrator for further assistance." & vbCrLf & vbCrLf &
                 "The specific database error:" & vbCrLf & vbCrLf &
                 exSql.Message

            If arrParts.Length > 1 Then

                Dim strFKeyConstraint As String = arrParts(1)

                Dim strSQL As String = "SELECT [CONSTRAINT_NAME] AS FOREIGN_KEY_NAME," & vbCrLf &
                    "[COLUMN_NAME] as FOREIGN_KEY," & vbCrLf &
                    "[TABLE_NAME] as FOREIGN_KEY_TABLE" & vbCrLf &
                    "FROM [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE]" & vbCrLf &
                    "WHERE [CONSTRAINT_NAME] = '" & strFKeyConstraint & "'"

                Dim objDT As DataTable = GetDataTableBySQL(strSQL)

                If objDT.Rows.Count = 1 Then
                    Dim strTable As String = CStr(NullValue(objDT.Rows(0)("FOREIGN_KEY_TABLE"), ""))
                    Dim strField As String = CStr(NullValue(objDT.Rows(0)("FOREIGN_KEY"), ""))

                    Dim objTable As clsTable = SysInfo.Tables(strTable)
                    Dim objField As clsField = Nothing
                    If objTable IsNot Nothing Then
                        objField = SysInfo.Fields(objTable.ID & "_" & strField)
                    End If

                    If objField IsNot Nothing Then
                        strNewException = "The field '" & objField.CaptionText & "' is referencing " &
                            "a record which no longer exists (perhaps it was deleted by another user)." & vbCrLf &
                            vbCrLf & "Please clear this field or select a different value."
                    End If
                End If

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If
            End If
        Catch ex As Exception
        End Try

        Throw New clsK1Exception(enumSQLExceptions.FOREIGN_LINK_VIOLATION, strNewException)
    End Sub

    Private Sub ThrowDeleteError(ByVal exSql As SqlException)
        Dim strNewException As String = "The record could not be deleted. The record is being referenced by field ‘{0}’ on table ‘{1}’." &
                                        vbCrLf & vbCrLf & "Please consult your RecFind Administrator for further assistance."

        Try
            Dim strMessage As String = exSql.Message
            Dim arrParts As String() = strMessage.Replace("'"c, """"c).Split(""""c)
            Dim strTable As String = "Unknown"
            Dim strField As String = "Unknown"

            'strNewException = "There is a field which is referencing a record which no longer exists." & vbCrLf & _
            '     "Please consult your RecFind Administrator for further assistance." & vbCrLf & vbCrLf & _
            '     "The specific database error:" & vbCrLf & vbCrLf & _
            '     exSql.Message

            If arrParts.Length > 1 Then
                Dim strFKeyConstraint As String = arrParts(1)

                Dim strSQL As String = "SELECT [CONSTRAINT_NAME] AS FOREIGN_KEY_NAME," & vbCrLf &
                    "[COLUMN_NAME] as FOREIGN_KEY," & vbCrLf &
                    "[TABLE_NAME] as FOREIGN_KEY_TABLE" & vbCrLf &
                    "FROM [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE]" & vbCrLf &
                    "WHERE [CONSTRAINT_NAME] = '" & strFKeyConstraint & "'"

                Dim objDT As DataTable = GetDataTableBySQL(strSQL)

                If objDT.Rows.Count = 1 Then
                    strTable = CStr(NullValue(objDT.Rows(0)("FOREIGN_KEY_TABLE"), ""))
                    strField = CStr(NullValue(objDT.Rows(0)("FOREIGN_KEY"), ""))

                    '        Dim objTable As clsTable = SysInfo.Tables(strTable)
                    '        Dim objField As clsField
                    '        If objTable IsNot Nothing Then
                    '            objField = SysInfo.Fields(objTable.ID & "_" & strField)
                    '        End If

                    '        If objField IsNot Nothing Then
                    '            strNewException = "The field '" & objField.CaptionText & "' is referencing " & _
                    '                "a record which no longer exists (perhaps it was deleted by another user)." & vbCrLf & _
                    '                vbCrLf & "Please clear this field or select a different value."
                    '        End If
                End If

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If
            End If

            strNewException = String.Format(strNewException, strField, strTable)
        Catch ex As Exception
        End Try

        Throw New clsK1Exception(enumSQLExceptions.FOREIGN_LINK_VIOLATION, strNewException)
    End Sub

    Private Function ReverseFindPeriodString(ByVal strValue As String) As String
        If String.IsNullOrEmpty(strValue) Then
            Return strValue
        End If

        Dim intIndex As Integer = strValue.LastIndexOf("."c)
        If intIndex >= 0 Then
            strValue = strValue.Substring(intIndex + 1, strValue.Length - (intIndex + 1))
        End If

        Return strValue
    End Function
#End Region

#Region " Link Table Methods "

    ''' <summary>
    ''' Creates a Link Table record
    ''' </summary>
    Public Overrides Sub InsertLink(ByVal strLinkTable As String, ByVal strPrimeLink As String,
    ByVal intPrimeLinkID As Integer, ByVal strForeignLink As String, ByVal intForeignLinkID As Integer)
        ExecuteLinkSP(strLinkTable, enumSPType.Insert, strPrimeLink, intPrimeLinkID, strForeignLink, intForeignLinkID)
    End Sub

    ''' <summary>
    ''' Deletes a Link Table record
    ''' </summary>
    Public Overrides Sub DeleteLink(ByVal strLinkTable As String, ByVal strPrimeLink As String,
    ByVal intPrimeLinkID As Integer, ByVal strForeignLink As String, ByVal intForeignLinkID As Integer)
        If intForeignLinkID = clsDBConstants.cintNULL Then
            DeleteLinkSP(strLinkTable, strPrimeLink, intPrimeLinkID)
        Else
            ExecuteLinkSP(strLinkTable, enumSPType.Delete, strPrimeLink, intPrimeLinkID, strForeignLink, intForeignLinkID)
        End If
    End Sub

    ''' <summary>
    ''' Common code used for both the deleting and inserting of link records
    ''' </summary>
    Private Sub ExecuteLinkSP(ByVal strLinkTable As String, ByVal eSPType As enumSPType,
    ByVal strPrimeLink As String, ByVal intPrimeLinkID As Integer, ByVal strForeignLink As String,
    ByVal intForeignLinkID As Integer)
        Dim colParams As clsDBParameterDictionary
        Dim strStoredProcedure As String

        Try
            Dim arrFieldName() As String = {strPrimeLink, strForeignLink}
            Dim arrValue() As Object = {intPrimeLinkID, intForeignLinkID}

            Select Case eSPType
                Case enumSPType.Insert
                    strStoredProcedure = strLinkTable & clsDBConstants.StoredProcedures.cINSERT

                Case enumSPType.Delete
                    strStoredProcedure = strLinkTable & clsDBConstants.StoredProcedures.cDELETE
                Case Else
                    Throw New ArgumentException("Only an Insert or a Delete is allowed.", "eSPType")
            End Select

            colParams = GetParamCollection(m_objSysInfo, eSPType, strLinkTable, arrFieldName, arrValue)

            Me.ExecuteProcedure(strStoredProcedure, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Used when we just want to delete all link records by a single prime id
    ''' </summary>
    Private Sub DeleteLinkSP(ByVal strLinkTable As String,
    ByVal strPrimeLink As String, ByVal intPrimeLinkID As Integer)
        Dim colParams As clsDBParameterDictionary
        Dim strStoredProcedure As String

        Try
            Dim arrFieldName() As String = {strPrimeLink}
            Dim arrValue() As Object = {intPrimeLinkID}

            strStoredProcedure = strLinkTable & clsDBConstants.StoredProcedures.cDELETE
            colParams = GetParamCollection(m_objSysInfo, enumSPType.Delete, strLinkTable,
                arrFieldName, arrValue)

            Me.ExecuteProcedure(strStoredProcedure, colParams)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region " Blob Methods "

    ''' <summary>
    ''' Retrieves the size in bytes of the blob field
    ''' </summary>
    Public Overrides Function GetBLOBSize(ByVal strTable As String,
    ByVal strField As String, ByVal intID As Integer) As Integer
        Dim intDataLength As Integer = 0
        Dim objDT As DataTable = Nothing

        '2017/03/23 -- James -- Checking EDOCStatus = Archived and Image has a value in it before trying to load it from the Archive database
        Dim objEdocStatusCodesRepository = New clsEdocStatusCodesRepository(Me)
        Dim intArchivedID As Integer = objEdocStatusCodesRepository.GetStatusId(clsDBConstants.clsEdocStatusCodes.Archived)
        Dim intEDOCStatus As Integer = 0
        Dim intImageSize As Integer = 0

        '2017-08-31 -- Peter Melisi -- Bug fix for #1700003433
        If strTable.Equals(clsDBConstants.Tables.cEDOC) Then
            intEDOCStatus = ExecuteScalar("SELECT EDOCStatusID FROM EDOC WHERE ID=" & CStr(intID) & " AND EDOCStatusID IS NOT NULL")
            intImageSize = ExecuteScalar("SELECT datalength([Image]) FROM EDOC WHERE ID=" & CStr(intID))
        End If

        If intEDOCStatus > 0 AndAlso intEDOCStatus = intArchivedID AndAlso intImageSize <= 1 AndAlso strField <> "Thumbnail" Then
            Dim intDB As Integer = ExecuteScalar("SELECT DBID FROM K1Archive_Link WHERE EDOCID=" & CStr(intID))

            If intDB > 0 Then
                Dim objArchiveDT As DataTable = GetDataTableBySQL("SELECT ExternalID FROM K1Archive WHERE ID=" & CStr(intDB))

                If objArchiveDT IsNot Nothing AndAlso objArchiveDT.Rows.Count = 1 Then
                    Dim objEncryption As New clsEncryption(True)
                    Dim objConnection As New SqlConnection(objEncryption.Decrypt(CStr(objArchiveDT.Rows(0)(0))))
                    OpenConnectionWithContext(objConnection)

                    Dim cmdDB As SqlCommand = objConnection.CreateCommand()
                    cmdDB.CommandText = "SELECT TEXTPTR([" & strField & "]) POINTER, " &
                        "DATALENGTH([" & strField & "]) FILESIZE FROM [" & strTable & "] " &
                        "WHERE [" & clsDBConstants.Fields.cID & "] = " & intID
                    cmdDB.CommandType = CommandType.Text

                    Using objAdapter As New SqlDataAdapter() With {.SelectCommand = cmdDB}
                        objDT = New DataTable("DT")
                        objAdapter.Fill(objDT)
                    End Using
                End If

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objArchiveDT IsNot Nothing Then
                    objArchiveDT.Dispose()
                    objArchiveDT = Nothing
                End If
            End If
        Else
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@ID", intID))
            objDT = GetDataTableBySQL("SELECT TEXTPTR([" & strField & "]) POINTER, " &
                "DATALENGTH([" & strField & "]) FILESIZE FROM [" & strTable & "] " &
                "WHERE [" & clsDBConstants.Fields.cID & "] = @ID", colParams)
            colParams.Dispose()
        End If

        If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
            If Not objDT.Rows(0)("POINTER") Is DBNull.Value Then
                If (IsDBNull(objDT.Rows(0)("FILESIZE"))) Then
                    intDataLength = 0
                Else
                    intDataLength = CType(objDT.Rows(0)("FILESIZE"), Integer)
                End If
            End If
        Else
            Throw New Exception("The file does not exist in the database.") 'it may have been deleted by another user
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If

        Return intDataLength
    End Function

    ''' <summary>
    ''' Saves the blob to the specified file
    ''' </summary>
    Public Overrides Sub ReadBLOB(ByVal strTable As String,
                                  ByVal strField As String,
                                  ByVal intID As Integer,
                                  ByVal strFile As String,
                                  Optional ByVal blnShowTransferUI As Boolean = True)
        Dim objFS As IO.FileStream = Nothing

        Try
            m_blnCancel = False

            Dim intLength As Integer = GetBLOBSize(strTable, strField, intID)

            If blnShowTransferUI Then
                RaiseFileTransferInit(enumDataAccessType.DB_DIRECT,
                    CInt(Math.Ceiling(intLength / m_intBlobChunkSize)), enumTransferType.DOWNLOAD)
            End If

            Dim intOffSet As Integer = 0

            If Me.Session IsNot Nothing Then Me.Session.StartAutoUpdate()

            While intOffSet < intLength AndAlso Not m_blnCancel
                Dim arrBuffer() As Byte = GetBLOB(strTable, strField, intID, m_intBlobChunkSize, intOffSet)

                If intOffSet = 0 Then
                    CreateDir(IO.Path.GetDirectoryName(strFile))
                    objFS = New System.IO.FileStream(strFile, IO.FileMode.Create, IO.FileAccess.Write)
                End If
                objFS.Write(arrBuffer, 0, arrBuffer.Length)

                intOffSet += arrBuffer.Length

                If blnShowTransferUI Then
                    RaiseFileTransferStep()
                End If
            End While

            If objFS IsNot Nothing Then
                objFS.Close()
            End If

            If blnShowTransferUI Then
                RaiseFileTransferEnd()
            End If
        Catch ex As Exception
            m_blnCancel = True
            Throw
        Finally
            If Me.Session IsNot Nothing Then Me.Session.EndAutoUpdate()

            If objFS IsNot Nothing Then
                objFS.Close()
                objFS = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Saves the blob to the specified file
    ''' </summary>
    Public Overrides Function ReadBLOBToMemory(ByVal strTable As String,
    ByVal strField As String, ByVal intID As Integer) As Byte()
        m_blnCancel = False

        Dim intLength As Integer = GetBLOBSize(strTable, strField, intID)
        Dim arrBlob(intLength - 1) As Byte
        Dim intOffSet As Integer = 0

        'RaiseFileTransferInit(enumDataAccessType.DB_DIRECT, _
        '    CInt(Math.Ceiling(intLength / m_intBlobChunkSize)), enumTransferType.DOWNLOAD)

        While intOffSet < intLength AndAlso Not m_blnCancel
            Dim arrBuffer() As Byte = GetBLOB(strTable, strField, intID, m_intBlobChunkSize, intOffSet)

            If arrBuffer.Length > 0 Then
                arrBuffer.CopyTo(arrBlob, intOffSet)
            End If

            intOffSet += arrBuffer.Length

            ' RaiseFileTransferStep()
        End While

        Return arrBlob
        ' RaiseFileTransferEnd()
    End Function

    ''' <summary>
    ''' Saves the blob to the specified file
    ''' </summary>
    Public Function GetBLOB(ByVal strTable As String, ByVal strField As String, ByVal intID As Integer,
    ByVal intChunkSize As Integer, ByRef intOffSet As Integer) As Byte()
        Dim objDR As SqlClient.SqlDataReader = Nothing
        Dim objConnection As SqlClient.SqlConnection = Nothing
        Dim blnArchive As Boolean = False

        Try
#If TRACE Then
            Trace.WriteVerbose("(GetBLOB) " & strTable & "." & strField, "")
#End If

            Dim arrPointer(15) As Byte
            Dim intDataLength As Integer
            Dim arrBuffer As Byte()

            Dim objPointerParam As SqlClient.SqlParameter
            Dim objSizeParam As SqlClient.SqlParameter
            Dim objOffsetParam As SqlClient.SqlParameter

            Dim objDT As DataTable = Nothing
            Dim objTransaction As SqlTransaction = Nothing

            '2017/03/23 -- James -- Checking EDOCStatus = Archived and Image has a value in it before trying to load it from the Archive database
            Dim objEdocStatusCodesRepository = New clsEdocStatusCodesRepository(Me)
            Dim intArchivedID As Integer = objEdocStatusCodesRepository.GetStatusId(clsDBConstants.clsEdocStatusCodes.Archived)
            Dim intEDOCStatus As Integer = 0
            Dim intImageSize As Integer = 0

            '2017-08-31 -- Peter Melisi -- Bug fix for #1700003433
            If strTable.Equals(clsDBConstants.Tables.cEDOC) Then
                intEDOCStatus = ExecuteScalar("SELECT EDOCStatusID FROM EDOC WHERE ID=" & CStr(intID) & " AND EDOCStatusID IS NOT NULL")
                intImageSize = ExecuteScalar("SELECT datalength([Image]) FROM EDOC WHERE ID=" & CStr(intID))
            End If

            If intEDOCStatus > 0 AndAlso intEDOCStatus = intArchivedID AndAlso intImageSize <= 1 AndAlso strField <> "Thumbnail" Then
                Dim intDB As Integer = ExecuteScalar("SELECT DBID FROM K1Archive_Link WHERE EDOCID=" & CStr(intID))

                If intDB > 0 Then
                    Dim objArchiveDT As DataTable = GetDataTableBySQL("SELECT ExternalID FROM K1Archive WHERE ID=" & CStr(intDB))
                    If objArchiveDT IsNot Nothing AndAlso objArchiveDT.Rows.Count = 1 Then
                        blnArchive = True

                        Dim objEncryption As New clsEncryption(True)
                        objConnection = New SqlConnection(objEncryption.Decrypt(CStr(objArchiveDT.Rows(0)(0))))
                        OpenConnectionWithContext(objConnection)

                        objTransaction = objConnection.BeginTransaction()

                        Dim cmdDB As SqlCommand = objConnection.CreateCommand()
                        cmdDB.CommandText = "SELECT TEXTPTR([" & strField & "]) POINTER, " &
                            "DATALENGTH([" & strField & "]) FILESIZE FROM [" & strTable & "] " &
                            "WHERE [" & clsDBConstants.Fields.cID & "] = " & intID
                        cmdDB.CommandType = CommandType.Text
                        cmdDB.Connection = objConnection
                        cmdDB.Transaction = objTransaction

                        Using objAdapter As New SqlDataAdapter() With {.SelectCommand = cmdDB}
                            objDT = New DataTable("DT")
                            objAdapter.Fill(objDT)
                        End Using
                    End If

                    '2000003607 - Ara Melkonian - Removed return, archived EDOC's were not being retrieved.
                    '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                    If objArchiveDT IsNot Nothing Then
                        objArchiveDT.Dispose()
                        objArchiveDT = Nothing
                    End If
                Else
                    Return Nothing
                End If
            Else
                objConnection = GetConnection(True)
                objTransaction = GetTransaction()

                objDT = GetDataTableBySQL("SELECT TEXTPTR([" & strField & "]) POINTER, " &
                "DATALENGTH([" & strField & "]) FILESIZE FROM [" & strTable & "] " &
                "WHERE [" & clsDBConstants.Fields.cID & "] = " & intID)
            End If

            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                If objDT.Columns.Contains("POINTER") AndAlso Not objDT.Rows(0)("POINTER") Is DBNull.Value Then
                    arrPointer = CType(objDT.Rows(0)("POINTER"), Byte())
                    If (IsDBNull(objDT.Rows(0)("FILESIZE"))) Then
                        intDataLength = 0
                    Else
                        intDataLength = CType(objDT.Rows(0)("FILESIZE"), Integer)
                    End If
                ElseIf objDT.Columns.Count = 2 AndAlso Not objDT.Rows(0)(0) Is DBNull.Value Then
                    '-- Sometime it cant find the column by it name
                    arrPointer = CType(objDT.Rows(0)(0), Byte())
                    intDataLength = CType(objDT.Rows(0)(0), Integer)
                End If
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            If intDataLength = 0 Then
                Return Nothing
            End If

            ' Set up READTEXT command, parameters, and open BinaryReader.
            Dim objCommand As New SqlClient.SqlCommand("READTEXT [" & strTable & "].[" &
                strField & "] @Pointer @Offset @Size HOLDLOCK", objConnection)
            objCommand.Transaction = objTransaction
            objCommand.CommandTimeout = m_intCommandTimeout

            objPointerParam = objCommand.Parameters.Add("@Pointer", SqlDbType.Binary, 16)
            objOffsetParam = objCommand.Parameters.Add("@Offset", SqlDbType.Int)
            objSizeParam = objCommand.Parameters.Add("@Size", SqlDbType.Int)

            ' Calculate the buffer size - may be less than BUFFER_LENGTH for the last block.
            If intOffSet + intChunkSize >= intDataLength Then
                intChunkSize = intDataLength - intOffSet
            End If
            ReDim arrBuffer(intChunkSize - 1)

            objPointerParam.Value = arrPointer
            objOffsetParam.Value = intOffSet
            objSizeParam.Value = intChunkSize

            objDR = objCommand.ExecuteReader(CommandBehavior.SingleResult)
            objDR.Read()
            objDR.GetBytes(0, 0, arrBuffer, 0, intChunkSize)
            objDR.Close()

            Return arrBuffer
        Catch ex As SqlException
            ProcessSQLExceptionExecute(ex)
            Return Nothing
        Finally
            If Not objDR Is Nothing Then
                objDR.Close()
                objDR = Nothing
            End If

            If blnArchive Then
                objConnection.Close()
                objConnection.Dispose()
                objConnection = Nothing
            Else
                CloseConnection(objConnection)
            End If
        End Try

    End Function

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public Overloads Overrides Sub WriteBLOB(ByVal objTable As clsTable,
    ByVal objField As clsField, ByVal intID As Integer, ByVal strFile As String, Optional ByVal blnShowTransferUI As Boolean = True)
        WriteBLOB(objTable.DatabaseName, objField.DatabaseName, objField.DataType,
            objField.Length, intID, strFile, blnShowTransferUI)
    End Sub

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public Overloads Overrides Sub WriteBLOB(ByVal strTable As String,
    ByVal strField As String, ByVal eDataType As SqlDbType,
    ByVal intDataLength As Integer, ByVal intID As Integer, ByVal strFile As String, Optional ByVal blnShowTransferUI As Boolean = True)
        m_blnCancel = False
        Dim objFS As FileStream = IO.File.OpenRead(strFile)
        Dim intFileLength As Integer = CType(objFS.Length, Integer)
        objFS.Close()

        If blnShowTransferUI Then
            RaiseFileTransferInit(enumDataAccessType.DB_DIRECT,
                CInt(Math.Ceiling(intFileLength / m_intBlobChunkSize)), enumTransferType.UPLOAD)
        End If

        Dim intOffSet As Integer = 0

        Try
            If Me.Session IsNot Nothing Then Me.Session.StartAutoUpdate()

            While intOffSet < intFileLength AndAlso Not m_blnCancel
                SetBLOB(strTable, strField, eDataType, intDataLength,
                    intID, strFile, m_intBlobChunkSize, intOffSet)
                If blnShowTransferUI Then
                    RaiseFileTransferStep()
                End If
            End While
        Catch ex As clsK1Exception
            If Not (ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE OrElse
            ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_FIELD) Then
                Throw
            End If
        Finally
            If Me.Session IsNot Nothing Then Me.Session.EndAutoUpdate()
        End Try

        If blnShowTransferUI Then
            RaiseFileTransferEnd()
        End If
    End Sub

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public Sub SetBLOB(ByVal strTable As String, ByVal strField As String,
    ByVal eDataType As SqlDbType, ByVal intDataLength As Integer, ByVal intID As Integer,
    ByVal strFile As String, ByVal intChunkSize As Integer, ByRef intOffSet As Integer)

        Dim objFS As IO.FileStream = Nothing
        Dim arrBuffer() As Byte

        Try
#If TRACE Then
            Trace.WriteVerbose("(SetBLOB) " & strTable & "." & strField, "")
#End If

            objFS = IO.File.OpenRead(strFile)

            If (eDataType = SqlDbType.Binary OrElse eDataType = SqlDbType.VarBinary) AndAlso
            intDataLength < (objFS.Length + 4) Then
                Throw New Exception("The size of the file is greater than the specified " &
                    "length of field '" & strField & "'.")
            End If

            If intOffSet + intChunkSize >= objFS.Length Then
                intChunkSize = CType((objFS.Length - intOffSet), Integer)
            End If
            ReDim arrBuffer(intChunkSize - 1)

            objFS.Position = intOffSet
            objFS.Read(arrBuffer, 0, intChunkSize)
            objFS.Close()

            intOffSet = UpdateBLOB(strTable, strField, intID, arrBuffer, intOffSet)
        Catch ex As SqlException
            ProcessSQLExceptionExecute(ex)
        Finally
            If Not objFS Is Nothing Then
                objFS.Close()
                objFS = Nothing
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public Function UpdateBLOB(ByVal strTable As String, ByVal strField As String, ByVal intID As Integer,
                               ByVal arrBuffer() As Byte, ByVal intOffSet As Long) As Integer
        Dim arrPointer(15) As Byte

        Dim objConnection As SqlConnection = Nothing
        Dim objPointerParam As SqlClient.SqlParameter
        Dim objBytesParam As SqlClient.SqlParameter
        Dim objOffsetParam As SqlClient.SqlParameter
        Dim objDeleteParam As SqlClient.SqlParameter

        Try
#If TRACE Then
            Trace.WriteVerbose("(UpdateBLOB) " & strTable & "." & strField, "")
#End If

            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@ID", intID))

            Dim strQuery As String = "SELECT TEXTPTR([{0}]) POINTER FROM [{1}] WHERE [{2}] = @ID"
            Dim objDT As DataTable = GetDataTableBySQL(String.Format(strQuery, strField, strTable, clsDBConstants.Fields.cID),
                                                       colParams)
            colParams.Dispose()

            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                If Not objDT.Rows(0)("POINTER") Is DBNull.Value Then
                    arrPointer = CType(objDT.Rows(0)("POINTER"), Byte())
                End If
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            objConnection = GetConnection(True)

            strQuery = "UPDATETEXT [{0}].[{1}] @Pointer @Offset @DeleteLength @Bytes"
            Dim objCommand As New SqlClient.SqlCommand(String.Format(strQuery, strTable, strField), objConnection)

            objCommand.CommandTimeout = m_intCommandTimeout
            objCommand.Transaction = GetTransaction()

            objPointerParam = objCommand.Parameters.Add("@Pointer", SqlDbType.Binary, 16)
            objPointerParam.Value = arrPointer
            objBytesParam = objCommand.Parameters.Add("@Bytes", SqlDbType.Image, arrBuffer.Length)
            objBytesParam.Value = arrBuffer
            objOffsetParam = objCommand.Parameters.Add("@Offset", SqlDbType.Int)
            objOffsetParam.Value = intOffSet
            objDeleteParam = objCommand.Parameters.Add("@DeleteLength", SqlDbType.Int)
            If intOffSet = 0 Then
                objDeleteParam.Value = DBNull.Value '-- Needs to be NULL to replace what is already in the column
            Else
                objDeleteParam.Value = 0 '-- Need to set this otherwise it wont append
            End If

            objCommand.ExecuteNonQuery()

            Return CType((intOffSet + arrBuffer.Length), Integer)
        Finally
            CloseConnection(objConnection)
        End Try
    End Function

#End Region

#Region " Training "

    ''' <summary>
    ''' Determines if the record limit has been exceeded for the specified table
    ''' </summary>
    Public Overrides Function RecordCountExceeded(ByVal strTable As String) As Boolean
        Dim objTable As clsTable = m_objSysInfo.Tables(strTable)

        Dim strSQL As String = "SELECT COUNT([" & clsDBConstants.Fields.cID &
            "]) AS CurrentTrainingCount FROM [" & objTable.DatabaseName & "]"

        Dim dtCount As DataTable = GetDataTableBySQL(strSQL)

        If dtCount Is Nothing OrElse Not dtCount.Rows.Count = 1 OrElse
        CInt(dtCount.Rows(0).Item("CurrentTrainingCount")) >= m_intRecordLimit Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " Bulk Insert "

    ''' <summary>
    ''' Bulk Inserts the records of the datatable to the destination
    ''' </summary>
    ''' <param name="objDT">A data table containing the records to bulk insert</param>
    ''' <param name="strDestTable">The table to bulk insert the records into</param>
    Public Overrides Sub BulkInsert(ByVal objDT As DataTable, ByVal strDestTable As String)
        Dim conDB As SqlConnection = Nothing
        Try
            If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
                Return
            End If

#If TRACE Then
            Trace.WriteVerbose("(BulkInsert) " & strDestTable & " Records:" & objDT.Rows.Count, "")
#End If

            'Get or Create the Connection
            conDB = GetConnection(True)

            Dim objBI As New SqlBulkCopy(conDB, SqlBulkCopyOptions.Default, GetTransaction())

            objBI.DestinationTableName = strDestTable
            objBI.WriteToServer(objDT)
        Catch ex As SqlException
            ProcessSQLExceptionExecute(ex)
        Catch ex As Exception
            Throw
        Finally
            CloseConnection(conDB)
        End Try
    End Sub
#End Region

#Region " SQL Notification "

    Public Overrides Sub AddSqlNotification(ByVal strTable As String,
                                            ByVal strField As String,
                                            ByVal objCallBack As OnChangeEventHandler)
        Try
            Dim strSql As String = "SELECT [{1}].[{0}] FROM [dbo].[{1}]"

            '-- Build SQL from table and field name
            strSql = String.Format(strSql, strField, strTable)

            AddSqlNotification(strSql, objCallBack)
        Catch ex As Exception

        End Try
    End Sub

    Public Overrides Sub AddSqlNotification(ByVal strSQL As String,
                                            ByVal objCallBack As OnChangeEventHandler)
        Try
#If TRACE Then
            Trace.WriteVerbose("(AddSqlNotification) " & strSQL, "")
#End If
            FillDataTable(strSQL, CommandType.Text, Nothing, objCallBack)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " DBStateChanged "

    Protected Overrides Sub DBStateChanged()
        MyBase.DBStateChanged()

        DisposeTransactions()
    End Sub
#End Region

    Public Overrides Sub InitializeServerObjects()
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Sub DisposeTransactions()
        Try
            If m_colTransactions IsNot Nothing Then
                For Each objTransaction As clsDBTransaction In m_colTransactions.Values
                    objTransaction.Dispose()
                Next

                m_colTransactions.Dispose()
                m_colTransactions = Nothing
                m_blnMultiThreaded = False
            End If
        Catch ex As Exception

        End Try

        Try
            If Not m_objDBTransaction Is Nothing Then
                m_objDBTransaction.Dispose()
                m_objDBTransaction = Nothing
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                MyBase.Dispose(blnDisposing)

                m_objSysInfo = Nothing

                '2019-08-20 -- Emmanuel -- Fix Memory leak in Scheduled task service #1900003560
                m_dbSystem = Nothing

                DisposeTransactions()
            End If
        End If

        '[Naing Begin] Fix for Bug 1300002478
        If (m_blnUseSqlDependency) Then
            clsSqlDependency.StopBrokerAndResetFlag(m_strConnectionString)
        End If

        m_blnDisposedValue = True
    End Sub

    Public Overrides Sub CreateStoredProcedure(sbProcName As String, sbProcedure As String)
        dbSystem.CreateStoredProcedure(sbProcName, sbProcedure)
    End Sub

    Public Overrides Function CheckStoredProcedureExists(sbProcName As String) As Boolean
        Return dbSystem.CheckStoredProcedureExists(sbProcName)
    End Function
#End Region

End Class
