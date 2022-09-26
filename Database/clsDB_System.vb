Imports System.Data.SqlClient

Public Class clsDB_System
    Inherits clsDB_Direct

#Region " Member Variables "

    'Variables used for the training version of the database
    Private m_intSecurityID As Integer = clsDBConstants.cintNULL
    Private m_objSecurity As clsSecurity
    Private m_strGroup As String
    Private m_strServer As String
    Private m_strDatabase As String
    Private m_strUserID As String

#End Region

#Region " Properties "

    ''' <summary>
    ''' Set this property if you want all objects created with this class to have a default security
    ''' </summary>
    Public Property Security() As clsSecurity
        Get
            Return m_objSecurity
        End Get
        Set(ByVal value As clsSecurity)
            m_objSecurity = m_objSecurity
        End Set
    End Property

#End Region

    Enum ConstraintAction As Integer
        NoAction = 0
        Null
        Cascade
    End Enum

#Region " Constructors "

    Public Sub New(ByVal strConnection As String)
        MyBase.New(strConnection)
        Me.GetDatabaseInfo(m_strServer, m_strDatabase, m_strUserID)
    End Sub

    Public Sub New(ByVal strConnection As String, ByVal eAppType As clsDBConstants.enumApplicationType)
        MyBase.New(strConnection, eAppType)
        Me.GetDatabaseInfo(m_strServer, m_strDatabase, m_strUserID)
    End Sub

    Public Sub New(ByVal strConnection As String,
    ByVal objSysInfo As clsSysInfo,
    ByVal objSecurity As clsSecurity)

        MyBase.New(strConnection, objSysInfo)
        'm_objSysInfo = objSysInfo
        m_objSecurity = objSecurity
        Me.GetDatabaseInfo(m_strServer, m_strDatabase, m_strUserID)

    End Sub

    Public Sub New(ByVal strConnection As String, ByVal strApplicationName As String)

        MyBase.New(strConnection, strApplicationName)
        Me.GetDatabaseInfo(m_strServer, m_strDatabase, m_strUserID)

    End Sub

#End Region

#Region " Methods "

#Region " Common "

    Public Overrides Function TryInitializeSqlDependency() As Boolean
        Me.GetDatabaseInfo(m_strServer, m_strDatabase, m_strUserID)
        If (IsBrokerEnabled()) Then
            Return MyBase.TryInitializeSqlDependency()
        End If
        Return False
    End Function

#Region " Backup Database "

    ''' <summary>
    ''' Backs up the database to the specified path
    ''' </summary>
    Public Sub BackupDatabase(ByVal strFullPath As String)
        ExecuteSQL("BACKUP DATABASE [" & m_strDatabase & "] TO DISK = '" &
            SQLString(strFullPath) & "' WITH INIT")
    End Sub

#End Region

#Region " Get Column By SQL "

    ''' <summary>
    ''' Returns the first column in the first row of the returned query
    ''' </summary>
    Public Function GetColumnBySQL(ByVal strSQL As String) As Object
        Return GetColumnBySQL(strSQL, Nothing, 0)
    End Function

    ''' <summary>
    ''' Returns the specified column in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="intColumnIndex">Index of the column for the value that will be returned</param>
    Public Function GetColumnBySQL(ByVal strSQL As String, ByVal intColumnIndex As Integer) As Object
        Return GetColumnBySQL(strSQL, Nothing, intColumnIndex, Nothing)
    End Function

    ''' <summary>
    ''' Returns the column specified by its name in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="strColumnName">Name of the column for the value that will be returned</param>
    Public Function GetColumnBySQL(ByVal strSQL As String, ByVal strColumnName As String) As Object
        Return GetColumnBySQL(strSQL, Nothing, Integer.MinValue, strColumnName)
    End Function

    ''' <summary>
    ''' Returns the first column in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="colParams">Dictionary of input parameters for the query</param>
    Public Function GetColumnBySQL(ByVal strSQL As String, ByVal colParams As clsDBParameterDictionary) As Object
        Return GetColumnBySQL(strSQL, colParams, 0)
    End Function

    ''' <summary>
    ''' Returns the specified column in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="colParams">Dictionary of input parameters for the query</param>
    ''' <param name="intColumnIndex">Index of the column for the value that will be returned</param>
    Public Function GetColumnBySQL(ByVal strSQL As String, ByVal colParams As clsDBParameterDictionary,
    ByVal intColumnIndex As Integer) As Object
        Return GetColumnBySQL(strSQL, colParams, intColumnIndex, Nothing)
    End Function

    ''' <summary>
    ''' Returns the column specified by its name in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="colParams">Dictionary of input parameters for the query</param>
    ''' <param name="strColumnName">Name of the column for the value that will be returned</param>
    Public Function GetColumnBySQL(ByVal strSQL As String, ByVal colParams As clsDBParameterDictionary,
    ByVal strColumnName As String) As Object
        Return GetColumnBySQL(strSQL, colParams, Integer.MinValue, strColumnName)
    End Function

    ''' <summary>
    ''' Returns the column specified by either its index or name in the first row of the returned query
    ''' </summary>
    ''' <param name="strSQL">Query to perform</param>
    ''' <param name="colParams">Dictionary of input parameters for the query</param>
    ''' <param name="intColumnIndex">Index of the column for the value that will be returned</param>
    ''' <param name="strColumnName">Name of the column for the value that will be returned</param>
    ''' <remarks>
    ''' If a Column Index and Column name is provided then as long as the index is greater than or equal to zero
    ''' it will be used over the column name. This is due to accessing a column value is faster using an index 
    ''' then column name.
    ''' </remarks>
    Private Function GetColumnBySQL(ByVal strSQL As String, ByVal colParams As clsDBParameterDictionary,
    ByVal intColumnIndex As Integer, ByVal strColumnName As String) As Object
        Dim objDT As DataTable = GetDataTableBySQL(strSQL, colParams)

        If objDT IsNot Nothing AndAlso objDT.Rows.Count > 0 Then
            If intColumnIndex >= 0 Then
                Return objDT.Rows(0)(intColumnIndex)
            Else
                Return objDT.Rows(0)(strColumnName)
            End If
        Else
            Return Nothing
        End If
    End Function


    Public Function GetColumn(Of IType)(ByVal strStoredProcedure As String, ByVal colParams As clsDBParameterDictionary,
                                        ByVal intColumnIndex As Integer) As IType
        Dim objDT As DataTable = Me.GetDataTable(strStoredProcedure, colParams)

        If objDT IsNot Nothing AndAlso objDT.Rows.Count > 0 Then
            Return CType(objDT.Rows(0)(intColumnIndex), IType)
        Else
            Return Nothing
        End If
    End Function

#End Region


    ''' <summary>
    ''' Checks if the Read Committed Snapshots (new isolation level for transaction) is enabled for the current database.
    ''' </summary>
    ''' <remarks>This is used by the database updater.</remarks>
    Public Function IsSnapshotEnabled() As Boolean
        Dim objDT As DataTable = GetDataTableBySQL("SELECT is_read_committed_snapshot_on FROM sys.databases WHERE name='" & m_strDatabase & "'")

        If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
            Return False
        End If

        Return CBool(objDT.Rows(0)(0))
    End Function

#Region " Enable Snapshot "

    ''' <summary>
    ''' Enables Read Committed Snapshots (new isolation level for transaction) for the current database.
    ''' </summary>
    ''' <remarks>This is used by the database updater.</remarks>
    Public Sub EnableSnapshot()
        ExecuteSQL("ALTER DATABASE " & m_strDatabase & " SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE")
    End Sub

#End Region

#Region " Enable Broker "

    ''' <summary>
    ''' Sets up the SQL Notification Broker for the current database given the required privileges to the current user.
    ''' </summary>
    ''' <remarks>This is used by the database updater.</remarks>
    Public Sub EnableBroker()
        Try
            Dim strQuery As String =
            "ALTER DATABASE [{DATASOURCE}] SET NEW_BROKER" & vbCrLf & vbCrLf &
            "ALTER DATABASE [{DATASOURCE}] SET ENABLE_BROKER" & vbCrLf & vbCrLf &
            "ALTER DATABASE [{DATASOURCE}] SET ANSI_NULLS ON, ANSI_PADDING ON, ANSI_WARNINGS ON, ARITHABORT ON, CONCAT_NULL_YIELDS_NULL ON, QUOTED_IDENTIFIER ON" & vbCrLf & vbCrLf &
            "--DBA creates a new role" & vbCrLf &
            "IF NOT EXISTS(SELECT uid FROM [dbo].[sysusers] WHERE [name] = 'recfind_dependency_subscriber')" & vbCrLf &
            vbTab & "EXEC sp_addrole 'recfind_dependency_subscriber'" & vbCrLf & vbCrLf &
            "--Permissions needed for startUser" & vbCrLf &
            "GRANT CREATE PROCEDURE TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT CREATE QUEUE TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT CREATE SERVICE TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT REFERENCES ON CONTRACT::[http://schemas.microsoft.com/SQL/Notifications/PostQueryNotification] to [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT VIEW DEFINITION TO [{K1USER}]" & vbCrLf & vbCrLf &
            "--Permissions needed for executeUser" & vbCrLf &
            "GRANT SELECT TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT SUBSCRIBE QUERY NOTIFICATIONS TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT RECEIVE ON QueryNotificationErrorsQueue TO [{K1USER}]" & vbCrLf & vbCrLf &
            "GRANT REFERENCES ON CONTRACT::[http://schemas.microsoft.com/SQL/Notifications/PostQueryNotification] to [{K1USER}]" & vbCrLf & vbCrLf &
            "EXEC sp_addrolemember 'recfind_dependency_subscriber', '{K1USER}'"

            '-- Replace tokens with dynamic data
            strQuery = strQuery.Replace("{DATASOURCE}", m_strDatabase)
            strQuery = strQuery.Replace("{K1USER}", m_strUserID)

            ClearConnectionPool()
            KillDatabaseConnections(m_strDatabase)

            '-- Run Setup Script
            Me.ExecuteSQL(strQuery)
            'Next
        Catch ex As Exception
            Throw
        End Try

    End Sub

#End Region

    ''' <summary>
    ''' Kills all active connections for the currently used database.
    ''' </summary>
    ''' <remarks>
    ''' If current connection is using the specified database then 
    ''' the current connection will not be killed.
    ''' </remarks>
    Public Sub KillDatabaseConnections()
        KillDatabaseConnections(m_strDatabase)
    End Sub

    ''' <summary>
    ''' Kills all active connections for a specified database.
    ''' </summary>
    ''' <param name="strDatabase">Name of database to kill connections for.</param>
    ''' <remarks>
    ''' If current connection is using the specified database then 
    ''' the current connection will not be killed.
    ''' </remarks>
    Public Sub KillDatabaseConnections(ByVal strDatabase As String)
        Dim strQuery As String = "DECLARE @KillQuery VARCHAR(8000); {1}" &
                                 "SET @KillQuery = '' {1}" &
                                 "SELECT @KillQuery = COALESCE(@KillQuery, ', ') + 'kill ' + CONVERT(VARCHAR, spid) + '; ' {1}" &
                                 "FROM master..sysprocesses " &
                                 "WHERE spid >=50 AND  (dbid=db_id('{0}') OR hostprocess=(SELECT hostprocess FROM master..sysprocesses WHERE spid=@@spid)) " &
                                 "AND NOT spid = @@spid {1}{1}" &
                                 "EXEC(@KillQuery)"

        Me.ExecuteSQL(String.Format(strQuery, strDatabase, Environment.NewLine))
    End Sub

    Public Sub ClearConnectionPool()
        Dim conDB As SqlConnection = Nothing
        Try
            conDB = GetConnection(True)

            SqlConnection.ClearPool(conDB)
        Catch ex As Exception
            Throw
        Finally
            CloseConnection(conDB)
        End Try

    End Sub

    'return a list of 
    Public Function GetLinkedFunctionsAndMethods() As String
        Dim strSQL As String = "SELECT [" & clsDBConstants.Tables.cDRMMethod & "].[" & clsDBConstants.Fields.DRMMethods.cMethodID & "], [" & clsDBConstants.Tables.cDRMMethod _
     & "].[" & clsDBConstants.Fields.cEXTERNALID & "], [" & clsDBConstants.Tables.cDRMFunction & "].[" & clsDBConstants.Fields.DRMFunctions.cUIID _
     & "], [" & clsDBConstants.Tables.cDRMMethod & "].[" & clsDBConstants.Fields.cID & "] FROM [" & clsDBConstants.Tables.cDRMMethod & "] INNER JOIN [" _
     & clsDBConstants.Tables.cDRMFunction & "] ON [" & clsDBConstants.Tables.cDRMFunction & "].[" & clsDBConstants.Fields.cID & "] = [" & clsDBConstants.Tables.cDRMMethod _
     & "].[" & clsDBConstants.Fields.DRMMethods.cDRMFunctionID & "]"
        Return strSQL
    End Function

    Public ReadOnly Property ServerName As String
        Get
            Return m_strServer
        End Get
    End Property

    Public ReadOnly Property DBName As String
        Get
            Return m_strDatabase
        End Get
    End Property
#End Region

#Region " Tables "

#Region " Create Table "

    ''' <summary>
    ''' Creates a table with an identity field
    ''' </summary>
    Public Sub CreateTable(ByVal strTable As String)
        Dim strSQL As String = "CREATE TABLE dbo.[{0}](" & vbCrLf &
                                    vbTab & "{1} INT IDENTITY NOT FOR REPLICATION, " & vbCrLf &
                                    vbTab & "CONSTRAINT PK_{0} PRIMARY KEY ({1})" & vbCrLf &
                                ")"
        ExecuteSQL(String.Format(strSQL, strTable, clsDBConstants.Fields.cID))
    End Sub

#End Region

#Region " Delete Table "

    ''' <summary>
    ''' Deleted the specified table
    ''' </summary>
    Public Sub DeleteTable(ByVal strTable As String)
        If TableExists(strTable) Then
            ExecuteSQL("DROP TABLE [" & strTable & "]")
        End If
    End Sub
#End Region

#Region " Table Exists "

    ''' <summary>
    ''' Checks if the table name exists in the database
    ''' </summary>
    ''' <param name="strTableName">Name of table we are looking for</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TableExists(ByVal strTableName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@Table", strTableName))
            Dim strSQL As String = "SELECT 1 FROM [INFORMATION_SCHEMA].[TABLES] " &
                "WHERE [TABLE_NAME] = @Table"
            Dim blnExists As Boolean = CBool(ExecuteScalar(strSQL, colParams))
            colParams.Dispose()

            If Not blnExists Then
                Return False
            Else
                Return blnExists
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Rename Table "

    ''' <summary>
    ''' Renames the specified table if the new table name does not already exist in the database
    ''' </summary>
    ''' <param name="strOldName">Name of the table we want to rename</param>
    ''' <param name="strNewName">New name of the table</param>
    ''' <remarks>
    ''' A NullReferenceException is thrown if strTable does not match the DatabaseName property case.
    ''' An Exception is thrown if the new table name already exist.
    ''' </remarks>
    Public Sub RenameTable(ByVal strOldName As String, ByVal strNewName As String)
        Dim strSQL As String
        Dim blnStartTransaction As Boolean = False
        Dim objTable As clsTable

        objTable = SysInfo.Tables(strOldName)

        If objTable Is Nothing Then
            Throw New NullReferenceException("Invalid table name or table does not exist. " &
                "Table name is case-sensitive make sure you have the correct case.")
        End If

        If TableExists(strNewName) Then
            Throw New Exception("The table '" & strNewName & "' already exists in the database.")
        End If

        Try
            If GetTransaction() Is Nothing Then
                blnStartTransaction = True
                BeginTransaction()
            End If

            '-- Rename indexes on this table
            clsTableIndex.RenameTableIndexes(Me, strOldName, strNewName)

            strSQL = "EXEC sp_rename 'dbo." & SQLString(strOldName) &
                "', '" & SQLString(strNewName) & "'" & vbCrLf
            ExecuteSQL(strSQL)

            '-- Rename foreign keys to or from this table
            RenameTableInForeignKey(strOldName, strNewName)

            strSQL = "UPDATE [Table] SET DatabaseName = '" & SQLString(strNewName) & "' " &
                "WHERE [ID] = " & objTable.ID

            ExecuteSQL(strSQL)

            Dim objNewTable As clsTable = clsTable.GetItem(objTable.ID, Me)

            '================================================================================
            'Copy Fields over to new object
            '================================================================================
            For Each objKeyPair As Generic.KeyValuePair(Of String, clsField) In objTable.Fields
                objNewTable.Fields(objKeyPair.Key) = objKeyPair.Value
            Next

            '================================================================================
            'Copy Field Links over to new object
            '================================================================================
            For Each objKeyPair As Generic.KeyValuePair(Of String, clsFieldLink) In objTable.FieldLinks
                objNewTable.FieldLinks(objKeyPair.Key) = objKeyPair.Value
            Next

            '-- Update Table in memory object
            SysInfo.DRMInsertUpdateTable(objNewTable)

            '-- Recreate Default Stored procedures for the table
            CreateStandardSPs(objNewTable)

            If blnStartTransaction Then
                EndTransaction(True)
            End If
        Catch ex As Exception
            If blnStartTransaction Then
                EndTransaction(False)
            End If

            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames the specified table if the new table name does not already exist in the database
    ''' </summary>
    ''' <param name="strOldName">Name of the table we want to rename</param>
    ''' <param name="strNewName">New name of the table</param>
    Public Sub RenameNonRecFindTable(ByVal strOldName As String, ByVal strNewName As String)

        Dim blnStartedTransaction As Boolean = False

        If TableExists(strNewName) Then
            Throw New Exception("The table '" & strNewName & "' already exists in the database.")
        End If

        Try
            If Not HasTransaction Then
                blnStartedTransaction = True
                BeginTransaction()
            End If

            clsTableIndex.RenameTableIndexes(Me, strOldName, strNewName)

            Dim strSQL As String = "EXEC sp_rename 'dbo." & SQLString(strOldName) &
                "', '" & SQLString(strNewName) & "'" & vbCrLf
            ExecuteSQL(strSQL)

            '-- Rename foreign keys to or from this table
            RenameTableInForeignKey(strOldName, strNewName)

            If blnStartedTransaction Then
                EndTransaction(True)
            End If
        Catch ex As Exception
            If blnStartedTransaction Then
                EndTransaction(False)
            End If

            Throw
        End Try
    End Sub
#End Region

#Region " Find Tables "

    ''' <summary>
    ''' Checks if the range of tables exist in the database
    ''' </summary>
    ''' <param name="arrTableNames">Name of the table that the field might exist on</param>
    ''' <returns>datatable of tables that do exist</returns>
    ''' <remarks></remarks>
    Public Function FindTables(ByVal arrTableNames() As String) As DataTable
        Try
            Dim strSQL As String =
            "SELECT [TABLE_NAME] FROM [INFORMATION_SCHEMA].[TABLES] " &
            "WHERE [TABLE_NAME] IN ('" & ImplodeArray(arrTableNames, "', '") & "')"

            Dim objDT As DataTable = GetDataTableBySQL(strSQL)

            Return objDT
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#End Region

#Region " Fields "

#Region " Create Field "

    ''' <summary>
    ''' Creates a field in the specified table
    ''' </summary>
    ''' <param name="strFieldDeclaration">The field to create (Ex. [Description] NVARCHAR(100) NULL)</param>
    Public Sub CreateField(ByVal strTable As String, ByVal strFieldDeclaration As String)
        ExecuteSQL("ALTER TABLE [" & strTable & "] ADD " & strFieldDeclaration)
    End Sub

    ''' <summary>
    ''' Creates a field in the specified table
    ''' </summary>
    Public Sub CreateField(ByVal objTable As clsTable, ByVal objField As clsField, ByVal strDefault As String)
        If Not objField.IsIdentityField Then
            CreateField(objTable.DatabaseName, GetFieldDeclaration(objField) &
                " " & GetFieldDefaultDeclaration(objTable, objField, strDefault))
        End If
    End Sub

#End Region

#Region " Delete Field "

    ''' <summary>
    ''' Deletes the field from the specified table
    ''' </summary>
    Public Sub DeleteField(ByVal strTable As String, ByVal strField As String)
        ExecuteSQL("ALTER TABLE [" & strTable & "] DROP COLUMN [" & strField & "]")
    End Sub

    ''' <summary>
    ''' Deletes the field from the specified table
    ''' </summary>
    Public Sub DeleteField(ByVal objField As clsField)
        DeleteField(objField.Table.DatabaseName, objField.DatabaseName)
    End Sub

#End Region

#Region " Update Field "

    ''' <summary>
    ''' Updates the field in the specified table (will fail if indexes are not removed first)
    ''' </summary>
    ''' <param name="strFieldDeclaration">The field to modify (Ex. [Description] NVARCHAR(200) NULL)</param>
    Public Sub UpdateField(ByVal strTable As String, ByVal strFieldDeclaration As String)
        ExecuteSQL("ALTER TABLE [" & strTable & "] ALTER COLUMN " & strFieldDeclaration)
    End Sub

    ''' <summary>
    ''' Deletes indexes and default constraint, updates the field, recreates indexes
    ''' </summary>
    Public Sub UpdateField(ByVal objField As clsField, ByVal strDefaultValue As String)
        Try
            If objField.DataType <> SqlDbType.NText AndAlso
            objField.DataType <> SqlDbType.Image AndAlso
            objField.DataType <> SqlDbType.Text Then
                Dim objTable As clsTable = objField.Table
                Dim colTableIndexes As List(Of clsTableIndex) = clsTableIndex.GetTableIndexes(
                    Me, objTable.DatabaseName, objField.DatabaseName)

                For Each objTableIndex As clsTableIndex In colTableIndexes
                    '-- Drop table indexes that are bound to the column
                    clsTableIndex.DropIndex(Me, objTableIndex.Table, objTableIndex.Name)
                Next

                '-- Drop Default Constraint
                DropDefaultConstraint(objField)

                ' Create the SQL Alter Query from the Table Object Information
                UpdateField(objTable.DatabaseName, GetFieldDeclaration(objField))

                '-- Add Default Constraint
                If Not String.IsNullOrEmpty(strDefaultValue) Then
                    CreateDefaultConstraint(objField, strDefaultValue)
                End If

                For Each objTableIndex As clsTableIndex In colTableIndexes
                    '-- Create Indexes that were dropped
                    clsTableIndex.Create(Me, objTableIndex)
                Next
            Else
                Throw New Exception("Can't modify fields declared as TEXT, NTEXT, or IMAGE.")
            End If
        Catch ex As SqlException
            If ex.Number = 8115 Then
                Throw New Exception("If you are changing the scale for a field, " &
                    "you should also increase the precision, otherwise current data will get truncated.")
            Else
                Throw
            End If
        End Try
    End Sub

#End Region

#Region " Get Field Declaration "

    ''' <summary>
    ''' takes a field class and converts it to a SQL column declaration
    ''' </summary>
    Public Function GetFieldDeclaration(ByVal objField As clsField) As String
        Return GetFieldDeclaration(objField.DatabaseName, objField.DataType, objField.Length, objField.Scale, objField.IsNullable)
    End Function

    ''' <summary>
    ''' Generates a SQL column declaration
    ''' </summary>
    Public Function GetFieldDeclaration(ByVal strFieldName As String, ByVal eDataType As Data.SqlDbType) As String
        Return GetFieldDeclaration(strFieldName, eDataType, 0, 0, True)
    End Function

    ''' <summary>
    ''' Generates a SQL column declaration
    ''' </summary>
    Public Function GetFieldDeclaration(ByVal strFieldName As String, ByVal eDataType As Data.SqlDbType,
    ByVal blnIsNullable As Boolean) As String
        Return GetFieldDeclaration(strFieldName, eDataType, 0, 0, blnIsNullable)
    End Function

    ''' <summary>
    ''' Generates a SQL column declaration
    ''' </summary>
    Public Function GetFieldDeclaration(ByVal strFieldName As String, ByVal eDataType As Data.SqlDbType,
    ByVal intLength As Integer, ByVal intScale As Integer, ByVal blnIsNullable As Boolean) As String
        Dim strFieldDeclaration As String

        strFieldDeclaration = "[" & strFieldName & "] " & eDataType.ToString.ToUpper

        If Not (eDataType = SqlDbType.Money OrElse eDataType = SqlDbType.SmallMoney) Then

            If intLength > 0 AndAlso intScale = clsDBConstants.cintNULL Then
                strFieldDeclaration &= "(" & intLength & ")"
            End If

            If intLength > 0 AndAlso intScale >= 0 Then
                strFieldDeclaration &= "(" & intLength & ", " & intScale & ")"
            End If

            If intLength <= 0 AndAlso (eDataType = SqlDbType.NVarChar OrElse
            eDataType = SqlDbType.VarChar) Then
                strFieldDeclaration &= "(MAX)"
            End If
        End If

        If eDataType = SqlDbType.Char OrElse eDataType = SqlDbType.NChar OrElse eDataType = SqlDbType.NText _
        OrElse eDataType = SqlDbType.NVarChar OrElse eDataType = SqlDbType.Text OrElse eDataType = SqlDbType.VarChar Then
            strFieldDeclaration &= " COLLATE database_default"
        End If

        If blnIsNullable Then
            strFieldDeclaration &= " NULL"
        Else
            strFieldDeclaration &= " NOT NULL"
        End If

        Return strFieldDeclaration
    End Function

#End Region

#Region " Field Exists "

    ''' <summary>
    ''' Checks if the field exists on the specified table
    ''' </summary>
    ''' <param name="strTableName">Name of the table that the field might exist on</param>
    ''' <param name="strFieldName">Name of the field we are looking for</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FieldExists(ByVal strTableName As String, ByVal strFieldName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@Table", strTableName))
            colParams.Add(New clsDBParameter("@Field", strFieldName))

            Dim strSQL As String = "SELECT 1 FROM [INFORMATION_SCHEMA].[COLUMNS] " &
                "WHERE [TABLE_NAME] = @Table AND [COLUMN_NAME] = @Field"

            Dim blnExists As Boolean = CBool(ExecuteScalar(strSQL, colParams))
            colParams.Dispose()

            If Not blnExists Then
                Return False
            Else
                Return blnExists
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Rename Field "

    ''' <summary>
    ''' Renames the specified RecFind field if the new field name does not already exist in a given table 
    ''' </summary>
    ''' <param name="strTableName">Name of the table that the field is in</param>
    ''' <param name="strFieldName">Name of the field we want to rename</param>
    ''' <param name="strNewName">New name of the field</param>
    ''' <remarks>
    ''' A NullReferenceException is thrown if strTable does not match the DatabaseName property case.
    ''' A NullReferenceException is thrown if strField does not match the DatabaseName property case.
    ''' An Exception is thrown if the new field name already exists in the table.
    ''' </remarks>
    Public Sub RenameRecFindField(ByVal strTableName As String, ByVal strFieldName As String,
    ByVal strNewName As String)
        Dim objTable As clsTable
        Dim objField As clsField
        Dim strSQL As String
        Dim blnStartedTransaction As Boolean = False

        If FieldExists(strTableName, strNewName) Then
            Throw New Exception("The field '" & strNewName & "' already exists in table '" & strTableName & "'.")
        End If

        objTable = SysInfo.Tables(strTableName)

        If objTable Is Nothing Then
            Throw New NullReferenceException("Invalid table name or table does not exist. " &
                "Table name is case-sensitive make sure you have the correct case.")
        End If

        objField = SysInfo.Fields(objTable.ID & "_" & strFieldName)

        If objField Is Nothing Then
            Throw New NullReferenceException("Invalid field name or field does not exist in the table. " &
                "Field name is case-sensitive make sure you have the correct case.")
        End If

        If objField.IsSystemLocked Then
            Throw New Exception("This field cannot be renamed as it is system locked.")
        End If

        'If FieldExists(strTableName, strNewName) Then
        '    Throw New Exception("The field '" & strNewName & "' already exists in table '" & strTableName & "'.")
        'End If

        Try
            If Not HasTransaction Then
                blnStartedTransaction = True
                BeginTransaction()
            End If

            strSQL = "UPDATE [Field] SET DatabaseName = '" & SQLString(strNewName) & "' " &
                "WHERE [ID] = " & objField.ID
            ExecuteSQL(strSQL)

            strSQL = "EXEC sp_rename 'dbo." & SQLString(objTable.DatabaseName) & "." &
                SQLString(objField.DatabaseName) & "', '" & SQLString(strNewName) & "'" & vbCrLf

            ExecuteSQL(strSQL)

            If objField.IsForeignKey Then
                '-- Rename all foreign keys the field appears in
                RenameFieldInForeignKey(strTableName, strFieldName, strNewName)
            End If

            '-- Rename all indexes the field appears in
            clsTableIndex.RenameFieldIndexes(Me, strTableName, strFieldName, strNewName)

            objField = clsField.GetItem(objField.ID, objField.Database)

            '-- Update Table and Field in memory object
            SysInfo.DRMInsertUpdateField(objTable, objField)

            '-- Recreate Default Stored procedures for the table
            CreateStandardSPs(objTable)

            If blnStartedTransaction Then
                EndTransaction(True)
            End If
        Catch ex As Exception
            If blnStartedTransaction Then
                EndTransaction(False)
            End If

            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames the specified field if the new field name does not already exist in a given table 
    ''' </summary>
    ''' <param name="strTableName">Name of the table that the field is in</param>
    ''' <param name="strFieldName">Name of the field we want to rename</param>
    ''' <param name="strNewName">New name of the field</param>
    Public Sub RenameField(ByVal strTableName As String, ByVal strFieldName As String,
                           ByVal strNewName As String)

        Dim blnStartedTransaction As Boolean = False

        If FieldExists(strTableName, strNewName) Then
            Throw New Exception("The field '" & strNewName & "' already exists in table '" & strTableName & "'.")
        End If

        Try
            If Not HasTransaction Then
                blnStartedTransaction = True
                BeginTransaction()
            End If

            If FindForeignKeys(strTableName, strFieldName, Nothing).Rows.Count > 0 Then
                '-- Rename all foreign keys the field appears in
                RenameFieldInForeignKey(strTableName, strFieldName, strNewName)
            End If

            '-- Rename all indexes the field appears in
            clsTableIndex.RenameFieldIndexes(Me, strTableName, strFieldName, strNewName)

            ExecuteSQL(String.Format("EXEC sp_rename 'dbo.{0}.{1}', '{2}'", SQLString(strTableName),
                                                                            SQLString(strFieldName),
                                                                            SQLString(strNewName)))

            If blnStartedTransaction Then
                EndTransaction(True)
            End If
        Catch ex As Exception
            If blnStartedTransaction Then
                EndTransaction(False)
            End If

            Throw
        End Try
    End Sub


#End Region

#Region " Find Fields "

    ''' <summary>
    ''' Checks if the field exists on the specified table
    ''' </summary>
    ''' <param name="strTableName">Name of the table that the field might exist on</param>
    ''' <param name="arrFieldNames">Names of the field we are looking for</param>
    ''' <returns>Datatable of fields found</returns>
    ''' <remarks></remarks>
    Public Function FindFields(ByVal strTableName As String, ByVal arrFieldNames() As String) As DataTable
        Try
            Dim strSQL As String =
            "SELECT [COLUMN_NAME] FROM [INFORMATION_SCHEMA].[COLUMNS] WHERE [TABLE_NAME] = '" & SQLString(strTableName) & "' " &
            "AND [COLUMN_NAME] IN ('" & ImplodeArray(arrFieldNames, "', '") & "')"

            Dim objDT As DataTable = GetDataTableBySQL(strSQL)

            Return objDT
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#End Region

#Region " Stored Procedures "

    ''' <summary>
    ''' Remove the stored procedure from the database
    ''' </summary>
    Public Sub DeleteStoredProcedure(ByVal strSPName As String)
        ExecuteSQL("If OBJECT_ID('dbo.[" & strSPName & "]') IS NOT NULL " &
            "DROP PROCEDURE dbo.[" & strSPName & "]")
    End Sub

    ''' <summary>
    ''' Checks whether a stored procedure already exists in the database with the specified name
    ''' </summary>
    Public Overrides Function CheckStoredProcedureExists(ByVal strSPName As String) As Boolean
        Return CBool(ExecuteScalar("If OBJECT_ID('dbo.[" & strSPName & "]') IS NOT NULL " &
            "SELECT 1 AS [Exists] ELSE SELECT 0 AS [Exists]"))
    End Function

    ''' <summary>
    ''' Creates the specified stored procedure
    ''' </summary>
    Public Overrides Sub CreateStoredProcedure(ByVal strName As String, ByVal strBody As String)
        ExecuteSQL("CREATE PROCEDURE dbo.[" & strName & "] AS" & vbCrLf & vbCrLf & strBody)
    End Sub

    ''' <summary>
    ''' Creates the specified stored procedure
    ''' </summary>
    Public Overloads Sub CreateStoredProcedure(ByVal strName As String, ByVal strParams As String, ByVal strBody As String)
        ExecuteSQL("CREATE PROCEDURE dbo.[" & strName & "] (" & vbCrLf &
            vbTab & strParams & vbCrLf & ") AS" & vbCrLf & vbCrLf & strBody)
    End Sub

    ''' <summary>
    ''' Starting declaration for a stored procedure
    ''' </summary>
    Private Function StandardSPDeclaration(ByVal strSP As String, ByVal strParams As String) As String
        Dim strDeclare As String = "CREATE PROCEDURE dbo.[" & strSP & "]"

        If strParams IsNot Nothing AndAlso strParams.Length > 0 Then
            strDeclare &= "(" & strParams & vbCrLf & ") AS" & vbCrLf & vbCrLf
        End If

        Return strDeclare
    End Function

    ''' <summary>
    ''' Returns the name of all stored procedures whose name matches the given criteria
    ''' </summary>
    ''' <param name="strCriteria"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindStoredProcedures(ByVal strCriteria As String) As DataTable
        Return Me.GetDataTableBySQL("SELECT [ROUTINE_NAME] FROM [INFORMATION_SCHEMA].[Routines] " &
         "WHERE [ROUTINE_TYPE] = 'PROCEDURE' AND [ROUTINE_NAME] LIKE '" & strCriteria & "'")
    End Function

#Region " Stored Procedure Exists "

    ''' <summary>
    ''' Checks if the stored procedure exists.
    ''' </summary>
    ''' <param name="strSPName">Name of the Stored Procedure that we are looking for.</param>
    ''' <returns>Returns True if the stored procedure exists;Otherwise False.</returns>
    ''' <remarks></remarks>
    Public Function StoredProcedureExists(ByVal strSPName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@SP", strSPName))

            Dim strSQL As String = "SELECT 1 FROM [INFORMATION_SCHEMA].[ROUTINES] " &
                                   "WHERE [ROUTINE_TYPE] = 'PROCEDURE' AND [ROUTINE_NAME] = @SP"

            Dim blnExists As Boolean = CBool(ExecuteScalar(strSQL, colParams))
            colParams.Dispose()

            Return blnExists
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

    Public Function GetStoredProcedureText(ByVal strSPName As String) As String
        Dim sbSpText As New Text.StringBuilder
        Dim dtSPText As DataTable = GetDataTableBySQL("Exec sp_helptext '" & strSPName & "'")

        If dtSPText IsNot Nothing AndAlso dtSPText.Rows.Count > 0 Then
            For Each objRow As DataRow In dtSPText.Rows
                sbSpText.Append(objRow(0))
            Next
        Else
            Throw New Exception("Could not find stored procedure '" & strSPName & "'")
        End If

        Return sbSpText.ToString
    End Function

#End Region

#Region " Functions "

    ''' <summary>
    ''' Remove the function from the database
    ''' </summary>
    Public Sub DeleteFunction(ByVal strFunctionName As String)
        ExecuteSQL("If OBJECT_ID('dbo.[" & strFunctionName & "]') IS NOT NULL " &
            "DROP FUNCTION dbo.[" & strFunctionName & "]")
    End Sub

    '''' <summary>
    '''' Creates the specified function
    '''' </summary>
    'Public Sub CreateFunction(ByVal strName As String, ByVal strBody As String)
    '    ExecuteSQL("CREATE PROCEDURE dbo.[" & strName & "] AS" & vbCrLf & vbCrLf & strBody)
    'End Sub

    '''' <summary>
    '''' Creates the specified function
    '''' </summary>
    'Public Sub CreateFunction(ByVal strName As String, ByVal strParams As String, ByVal strBody As String)
    '    ExecuteSQL("CREATE PROCEDURE dbo.[" & strName & "] (" & vbCrLf & _
    '        vbTab & strParams & vbCrLf & ") AS" & vbCrLf & vbCrLf & strBody)
    'End Sub

    ''' <summary>
    ''' Returns the name of all functions whose name matches the given criteria
    ''' </summary>
    ''' <param name="strCriteria"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindFunctions(ByVal strCriteria As String) As DataTable
        Return Me.GetDataTableBySQL("SELECT [ROUTINE_NAME] FROM [INFORMATION_SCHEMA].[ROUTINES] " &
         "WHERE [ROUTINE_TYPE] = 'Function' AND [ROUTINE_NAME] LIKE '" & strCriteria & "'")
    End Function

#Region " Stored Procedure Exists "

    ''' <summary>
    ''' Checks if the function exists.
    ''' </summary>
    ''' <param name="strFunctionName">Name of the function that we are looking for.</param>
    ''' <returns>Returns True if the funtion exists;Otherwise False.</returns>
    ''' <remarks></remarks>
    Public Function FunctionExists(ByVal strFunctionName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@name", strFunctionName))

            Dim strSQL As String = "SELECT 1 FROM [INFORMATION_SCHEMA].[ROUTINES] " &
                                   "WHERE [ROUTINE_TYPE] = 'Function' AND [ROUTINE_NAME] = @name"

            Dim blnExists As Boolean = CBool(ExecuteScalar(strSQL, colParams))
            colParams.Dispose()

            Return blnExists
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

    Public Function GetFunctionText(ByVal strFunctionName As String) As String
        Dim sbSpText As New Text.StringBuilder
        Dim dtSPText As DataTable = GetDataTableBySQL("Exec sp_helptext '" & strFunctionName & "'")

        If dtSPText IsNot Nothing AndAlso dtSPText.Rows.Count > 0 Then
            For Each objRow As DataRow In dtSPText.Rows
                sbSpText.Append(objRow(0))
            Next
        Else
            Throw New Exception("Could not find the function '" & strFunctionName & "'")
        End If

        Return sbSpText.ToString
    End Function

#End Region

#Region " Triggers "

#Region " Create Trigger "

#Region " Using a clsMaskFieldDictionary "

    ''' <summary>
    ''' Creates a trigger in the database
    ''' </summary>
    ''' <param name="colMaskObjs">clsMaskFieldDictionary of the Trigger table record for the trigger</param>
    ''' <remarks></remarks>
    Public Sub CreateTrigger(ByVal colMaskObjs As clsMaskFieldDictionary)
        Try
            Dim blnOnInsert As Boolean = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cONINSERT, False), Boolean)
            Dim blnOnUpdate As Boolean = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cONUPDATE, False), Boolean)
            Dim blnOnDelete As Boolean = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cONDELETE, False), Boolean)
            Dim strName As String = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cDATABASENAME, ""), String)
            Dim strTriggerAction As String = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cTRIGGERACTION, ""), String)
            Dim strBody As String = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cSQL, ""), String)
            Dim intTableID As Integer = CType(colMaskObjs.GetMaskValue(clsDBConstants.Fields.Trigger.cTABLEID, clsDBConstants.cintNULL), Integer)
            Dim objTable As clsTable = m_objSysInfo.Tables(intTableID)

            CreateTrigger(objTable.DatabaseName, strName, strBody, strTriggerAction, blnOnInsert, blnOnUpdate, blnOnDelete)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Non-K1 triggers "

    ''' <summary>
    ''' Creates a trigger in the database
    ''' </summary>
    ''' <param name="strTableName">Name of the table that the trigger is binding to</param>
    ''' <param name="strTriggerName">Name of the trigger</param>
    ''' <param name="strBody">SQL body of the trigger</param>
    ''' <param name="strTriggerAction">Action trigger performs (possible values I or F or A)</param>
    ''' <param name="blnOnInsert">Does the trigger execute on an Insert?</param>
    ''' <param name="blnOnUpdate">Does the trigger execute on an Update?</param>
    ''' <param name="blnOnDelete">Does the trigger execute on an Delete?</param>
    ''' <remarks></remarks>
    Public Sub CreateTrigger(ByVal strTableName As String, ByVal strTriggerName As String,
          ByVal strBody As String, ByVal strTriggerAction As String, ByVal blnOnInsert As Boolean,
          ByVal blnOnUpdate As Boolean, ByVal blnOnDelete As Boolean)
        Try
            Dim strSQL As String
            Dim strAction As String = ""
            Dim strType As String = ""

            DropTrigger(strTriggerName)

            Select Case strTriggerAction.Trim
                Case "I"
                    strAction = "INSTEAD OF "
                Case "F", "A" '- For, After
                    strAction = "FOR "
            End Select

            If blnOnInsert Then
                AppendToCommaString(strType, "INSERT")
            End If

            If blnOnUpdate Then
                AppendToCommaString(strType, "UPDATE")
            End If

            If blnOnDelete Then
                AppendToCommaString(strType, "DELETE")
            End If

            strSQL = "CREATE TRIGGER [" & strTriggerName & "] ON [dbo].[" & strTableName & "]" & vbCrLf &
                strAction & strType & vbCrLf & "AS" & vbCrLf & vbCrLf & strBody

            ExecuteSQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#End Region

#Region " Drop Trigger "

    ''' <summary>
    ''' Drops the specified trigger from the database
    ''' </summary>
    ''' <param name="strTriggerName">name of the trigger we want to drop</param>
    ''' <remarks></remarks>
    Public Sub DropTrigger(ByVal strTriggerName As String)
        ExecuteSQL("If OBJECT_ID('dbo.[" & strTriggerName & "]') IS NOT NULL " &
            "DROP TRIGGER dbo.[" & strTriggerName & "]")
    End Sub

#End Region

#Region " Rename Trigger "

    ''' <summary>
    ''' Renames a trigger in the database
    ''' </summary>
    ''' <param name="strOldName">Current name of the trigger in the database</param>
    ''' <param name="strNewName">New name of the trigger</param>
    ''' <remarks>
    ''' sp_rename does not work with triggers (as it does not update syscomments) 
    ''' so they must be dropped and recreated when renaming.
    ''' </remarks>
    Public Sub RenameTrigger(ByVal strOldName As String, ByVal strNewName As String)
        Try
            Dim strTriggerText As String = GetTriggerText(strOldName)

            '-- Rename trigger in sql
            strTriggerText.Replace(strOldName, strNewName)

            '-- drop old trigger
            DropTrigger(strOldName)

            '-- recreate trigger in db
            ExecuteSQL(strTriggerText)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Trigger Exists "

    ''' <summary>
    ''' Checks if the trigger exists
    ''' </summary>
    ''' <param name="strTriggerName">Name of the trigger we are looking for</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TriggerExists(ByVal strTriggerName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@Name", strTriggerName))
            Dim strSQL As String = "SELECT 1 FROM sys.triggers WHERE [name] = @Name"

            Dim blnExists As Boolean = CBool(ExecuteScalar(strSQL, colParams))
            colParams.Dispose()

            If Not blnExists Then
                Return False
            Else
                Return blnExists
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

    Public Function GetTriggerText(ByVal strTiggerName As String) As String
        Dim sbTriggerText As New Text.StringBuilder
        Dim dtTriggerText As DataTable = GetDataTableBySQL("Exec sp_helptext '" & strTiggerName & "'")

        If dtTriggerText IsNot Nothing AndAlso dtTriggerText.Rows.Count > 0 Then
            For Each objRow As DataRow In dtTriggerText.Rows
                sbTriggerText.Append(objRow(0))
            Next
        Else
            Throw New Exception("Could not find trigger '" & strTiggerName & "'")
        End If

        Return sbTriggerText.ToString
    End Function


#End Region

#Region " Views "

    ''' <summary>
    ''' Remove the view from the database
    ''' </summary>
    Public Sub DeleteView(ByVal strViewName As String)
        ExecuteSQL(String.Format("IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[{0}]')) " &
                                 "DROP VIEW [{0}]", strViewName))
    End Sub

    ''' <summary>
    ''' Checks if the view exists in the database
    ''' </summary>
    Public Function ViewExists(ByVal strViewName As String) As Boolean
        Dim intResult As Integer = ExecuteScalar(String.Format("SELECT 1 FROM sys.views WHERE object_id = OBJECT_ID(N'[{0}]') ", strViewName))

        Return CBool(intResult)
    End Function

#End Region

#Region " Constraints & Foreign Keys "

    Public Function GetFieldDefaultDeclaration(ByVal objTable As clsTable,
    ByVal objField As clsField, ByVal strDefault As String) As String
        Return GetFieldDefaultDeclaration(objTable.DatabaseName, objField.DatabaseName, strDefault)
    End Function


    Public Function GetFieldDefaultDeclaration(ByVal strTableName As String, ByVal strFieldName As String,
    ByVal strDefault As String) As String
        If String.IsNullOrEmpty(strDefault) Then
            'strDefault = "''"
            Return ""
        End If

        Return String.Format("Constraint DF_{0}_{1} DEFAULT {2} WITH VALUES", strTableName, strFieldName, strDefault)
    End Function

#Region " Delete "

    ''' <summary>
    ''' Removes the default value constraint from a field if one exists
    ''' </summary>
    Public Sub DropDefaultConstraint(ByVal objField As clsField)
        DropDefaultConstraint(objField.Table.DatabaseName, objField.DatabaseName)
    End Sub

    ''' <summary>
    ''' Removes the default value constraint from a field if one exists
    ''' </summary>
    Public Sub DropDefaultConstraint(ByVal strTable As String, ByVal strField As String)
        Try
            Dim strSQL As String = "SELECT object_name(cdefault) FROM syscolumns " &
                "WHERE id = object_id('{0}') AND name = '{1}'"
            Dim strConstraintName As Object = GetColumnBySQL(String.Format(strSQL, strTable, strField))

            If strConstraintName IsNot System.DBNull.Value AndAlso strConstraintName IsNot Nothing Then
                ExecuteSQL(String.Format("ALTER TABLE [{0}] DROP CONSTRAINT [{1}]", strTable, strConstraintName))
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Removes a foreign key relationship
    ''' </summary>
    Public Sub DeleteForeignKey(ByVal objField As clsField, ByVal objRefTable As clsTable)
        DeleteForeignKey(objField.Table.DatabaseName, objField.DatabaseName, objRefTable.DatabaseName)
    End Sub

    ''' <summary>
    ''' Removes a foreign key relationship
    ''' </summary>
    Public Sub DeleteForeignKey(ByVal strTable As String, ByVal strField As String,
    ByVal strRefTable As String)
        ExecuteSQL("ALTER TABLE [" & strTable & "] DROP CONSTRAINT " &
            "[FK_" & strField & "_" & strTable & "_" & strRefTable & "]")
    End Sub

#End Region

#Region " Create "

    Public Sub CreateDefaultConstraint(ByVal objField As clsField, ByVal strDefault As String)
        CreateDefaultConstraint(objField.Table.DatabaseName, objField.DatabaseName, strDefault)
    End Sub

    Public Sub CreateDefaultConstraint(ByVal strTableName As String, ByVal strFieldName As String,
    ByVal strDefault As String)
        Dim strSQL As String = "ALTER TABLE [{0}] ADD CONSTRAINT DF_{0}_{1} DEFAULT {2} FOR [{1}] WITH VALUES"

        '-- Drop existing default constraint first
        DropDefaultConstraint(strTableName, strFieldName)

        '-- Add the default constraint for the column
        ExecuteSQL(String.Format(strSQL, strTableName, strFieldName, strDefault))
    End Sub

    ''' <summary>
    ''' Creates a foreign key relationship
    ''' </summary>
    Public Sub CreateForeignKey(ByVal objField As clsField, ByVal objIdentityTable As clsTable)
        CreateForeignKey(objField.Table.DatabaseName, objField.DatabaseName, objIdentityTable.DatabaseName)
    End Sub

    Public Sub CreateForeignKey(ByVal strTable As String, ByVal strField As String, ByVal strIdentityTable As String)
        CreateForeignKey(strTable, strField, strIdentityTable, ConstraintAction.NoAction)
    End Sub

    ''' <summary>
    ''' Creates a foreign key relationship
    ''' </summary>
    Public Sub CreateForeignKey(ByVal strTable As String, ByVal strField As String,
    ByVal strIdentityTable As String, ByVal eOnDelete As ConstraintAction)
        Dim strOnDelete As String

        Select Case eOnDelete
            Case ConstraintAction.Null
                strOnDelete = " ON DELETE SET NULL"

            Case ConstraintAction.Cascade
                strOnDelete = " ON DELETE CASCADE"

            Case Else
                strOnDelete = ""
        End Select

        ExecuteSQL("ALTER TABLE [" & strTable & "] ADD " &
        "CONSTRAINT [FK_" & strField & "_" & strTable & "_" & strIdentityTable & "] " &
        "FOREIGN KEY ([" & strField & "]) " &
        "REFERENCES [" & strIdentityTable & "] ([" & clsDBConstants.Fields.cID & "])" & strOnDelete)
    End Sub

#End Region

#Region " Rename "

    ''' <summary>
    ''' Renames foreign keys that are bound to a specified table
    ''' </summary>
    ''' <param name="strTableName">Name of table that FKs are bound to</param>
    ''' <param name="strNewTableName">New name of table</param>
    ''' <remarks></remarks>
    Public Sub RenameTableInForeignKey(ByVal strTableName As String, ByVal strNewTableName As String)
        Try
            Dim dtForeignKeys As DataTable = FindAllForeignKeys(strTableName)
            If dtForeignKeys IsNot Nothing AndAlso dtForeignKeys.Rows.Count > 0 Then
                Dim blnUpdate As Boolean = False
                For Each objRow As DataRow In dtForeignKeys.Rows
                    blnUpdate = False
                    Dim arrParts() As String = CStr(objRow(0)).Split(Chr(95))
                    For intIndex As Integer = 1 To arrParts.GetUpperBound(0)
                        If arrParts(intIndex) = strTableName Then
                            arrParts(intIndex) = strNewTableName
                            blnUpdate = True
                        End If
                    Next

                    If blnUpdate Then RenameForeignKey(CType(objRow(0), String), ImplodeArray(arrParts, "_"))
                Next
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames foreign keys that are bound to a specified table and field
    ''' </summary>
    ''' <param name="strTableName">Name of table that FKs are bound to</param>
    ''' <param name="strFieldName">Name of the field that the FKs are bound to</param>
    ''' <param name="strNewFieldName">New name of the field</param>
    ''' <remarks></remarks>
    Public Sub RenameFieldInForeignKey(ByVal strTableName As String, ByVal strFieldName As String,
    ByVal strNewFieldName As String)
        Try
            Dim dtForeignKeys As DataTable = FindForeignKeys(strTableName, strFieldName, Nothing)
            If dtForeignKeys IsNot Nothing AndAlso dtForeignKeys.Rows.Count > 0 Then
                Dim arrForeignKeys() As DataRow = dtForeignKeys.Select("CONSTRAINT_TYPE = 'FOREIGN KEY'")
                For Each objRow As DataRow In arrForeignKeys
                    Dim strNewName As String = CStr(objRow(0)).Replace("_" & strFieldName & "_", "_" & strNewFieldName & "_")
                    RenameForeignKey(CType(objRow(0), String), strNewName)
                Next
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames the specified foreign key
    ''' </summary>
    ''' <param name="strOldName">Current name of the foreign key (format FK_[Foreign Field]_[Foreign Table]_[Identity Table])</param>
    ''' <param name="strNewName">New name of the foreign key (format FK_[Foreign Field]_[Foreign Table]_[Identity Table])</param>
    ''' <remarks></remarks>
    Private Sub RenameForeignKey(ByVal strOldName As String, ByVal strNewName As String)
        Try
            Dim strSQL As String = "EXEC sp_rename 'dbo." & SQLString(strOldName) & "', '" & SQLString(strNewName) & "'"

            ExecuteSQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Search "

    ''' <summary>
    ''' Finds all Foreign Keys for the specified Table
    ''' </summary>
    ''' <param name="strTableName">Name of the table (If nothing or empty will be treated as '%')</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindAllForeignKeys(ByVal strTableName As String) As DataTable
        Try
            Dim strWhere As String = ""
            Dim strSQL As String =
              "SELECT [CONSTRAINT_NAME], [CONSTRAINT_TYPE] FROM [INFORMATION_SCHEMA].[TABLE_CONSTRAINTS] WHERE "


            If Not String.IsNullOrEmpty(strTableName) Then
                strWhere &= "[CONSTRAINT_NAME] LIKE 'FK%~_" & SQLString(strTableName) & "' ESCAPE '~'"
            End If

            If strWhere.Length > 0 Then
                strWhere &= " OR "
            End If
            strWhere &= "[CONSTRAINT_NAME] LIKE 'FK~_%~_" & SQLString(strTableName) & "~_%' ESCAPE '~'"

            Return GetDataTableBySQL(strSQL & strWhere)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Finds all Foreign Keys for the specified Identity Table or Foreign Table and Field
    ''' </summary>
    ''' <param name="strForeignTableName">Name of the foreign table (If nothing or empty will be treated as '%')</param>
    ''' <param name="strForeignFieldName">Name of the foreign field (If nothing or empty will be treated as '%')</param>
    ''' <param name="IdentityTableName">Name of the identity table (If nothing or empty will be ignored)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindForeignKeys(ByVal strForeignTableName As String, ByVal strForeignFieldName As String,
    ByVal IdentityTableName As String) As DataTable
        Try
            Dim strWhere As String = ""
            Dim strSQL As String =
              "SELECT [CONSTRAINT_NAME], [CONSTRAINT_TYPE] FROM [INFORMATION_SCHEMA].[TABLE_CONSTRAINTS] WHERE "


            If Not String.IsNullOrEmpty(IdentityTableName) Then
                strWhere &= "[CONSTRAINT_NAME] LIKE 'FK%[_]" & SQLString(IdentityTableName) & "' "
            End If

            If String.IsNullOrEmpty(strForeignTableName) Then
                strForeignTableName = "%"
            End If

            If String.IsNullOrEmpty(strForeignFieldName) Then
                strForeignFieldName = "%"
            End If

            If Not strForeignTableName = "%" AndAlso Not strForeignFieldName = "%" Then
                If strWhere.Length > 0 Then
                    strWhere &= "OR "
                End If
                strWhere &= "[CONSTRAINT_NAME] LIKE 'FK[_]" & SQLString(strForeignFieldName) & "[_]" & SQLString(strForeignTableName) & "[_]%'"
            End If

            Return GetDataTableBySQL(strSQL & strWhere)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetForeignKeyInfo(ByVal strForeignTableName As String,
                                      ByVal strForeignFieldName As String,
                                      ByVal IdentityTableName As String) As DataTable
        Dim strWhere As String = ""
        Dim strSQL As String =
          "SELECT " &
            "fk.[CONSTRAINT_NAME] AS FOREIGN_KEY_NAME, " &
            "fk.[COLUMN_NAME] as FOREIGN_KEY, " &
            "fk.[TABLE_NAME] as FOREIGN_KEY_TABLE, " &
            "pk.[CONSTRAINT_NAME] as PRIMARY_KEY_NAME, " &
            "pk.[COLUMN_NAME] as PRIMARY_KEY, " &
            "pk.[TABLE_NAME] as PRIMARY_KEY_TABLE " &
        "FROM [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE] as fk " &
        "JOIN [INFORMATION_SCHEMA].[REFERENTIAL_CONSTRAINTS] as rc ON fk.[CONSTRAINT_NAME] = rc.[CONSTRAINT_NAME] " &
        "JOIN [INFORMATION_SCHEMA].[CONSTRAINT_COLUMN_USAGE] as pk ON rc.[UNIQUE_CONSTRAINT_NAME] = pk.[CONSTRAINT_NAME] "


        If Not String.IsNullOrEmpty(IdentityTableName) Then
            strWhere &= "pk.[TABLE_NAME] = '" & SQLString(IdentityTableName) & "' "
        End If

        If Not String.IsNullOrEmpty(strForeignFieldName) Then
            If strWhere.Length > 0 Then
                strWhere &= "OR "
            End If
            strWhere &= "fk.[COLUMN_NAME] = '" & SQLString(strForeignFieldName) & "' "
        End If

        If Not String.IsNullOrEmpty(strForeignTableName) Then
            If strWhere.Length > 0 AndAlso String.IsNullOrEmpty(strForeignFieldName) Then
                strWhere &= "OR "
            ElseIf strWhere.Length > 0 Then
                strWhere &= "AND "
            End If
            strWhere &= "fk.[TABLE_NAME] = '" & SQLString(strForeignTableName) & "' "
        End If

        Return GetDataTableBySQL(strSQL & CStr(IIf(strWhere.Length > 0, "WHERE " & strWhere, "")))
    End Function
#End Region

#End Region

#Region " Table Stored Procedure Creation Methods "

    ''' <summary>
    ''' Creates all the standard stored procedures for a table depending on the table class
    ''' </summary>
    Public Sub CreateStandardSPs(ByVal objTable As clsTable)
        If objTable.IsLinkTable Then
            CreateLinkSPInsert(objTable)
            CreateLinkSPDelete(objTable)
            CreateLinkSPGetList(objTable)
        Else
            CreateSPInsert(objTable)
            CreateSPUpdate(objTable)
            CreateSPDelete(objTable)
            CreateSPGetItem(objTable)
            CreateSPGetList(objTable)
        End If
    End Sub

    ''' <summary>
    ''' Deletes all the standard stored procedures for a table depending on the table class
    ''' </summary>
    ''' <param name="objTable"></param>
    ''' <remarks></remarks>
    Public Sub DeleteStandardSPs(ByVal objTable As clsTable)
        Dim strTableName As String = objTable.DatabaseName
        Dim eTableClass As clsDBConstants.enumTableClass = objTable.ClassType

        DeleteStoredProcedure(strTableName & clsDBConstants.StoredProcedures.cDELETE)
        DeleteStoredProcedure(strTableName & clsDBConstants.StoredProcedures.cGETLIST)
        DeleteStoredProcedure(strTableName & clsDBConstants.StoredProcedures.cINSERT)

        If Not eTableClass = clsDBConstants.enumTableClass.LINK_TABLE _
        AndAlso Not eTableClass = clsDBConstants.enumTableClass.LINK_TABLE_ESSENTIAL Then
            DeleteStoredProcedure(strTableName & clsDBConstants.StoredProcedures.cUPDATE)
            DeleteStoredProcedure(strTableName & clsDBConstants.StoredProcedures.cGETITEM)
        End If
    End Sub

    '''' <summary>
    '''' Will drop existing standard stored procedures and create new ones
    '''' </summary>
    '''' <param name="objTable"></param>
    '''' <remarks></remarks>
    'Public Sub RecreateStandardSPs(ByVal objTable As clsTable)
    '    DeleteStandardSPs(objTable)
    '    CreateStandardSPs(objTable)
    'End Sub

#Region " Common Stored Procedure Methods "

    ''' <summary>
    ''' Determines whether the datatype is included in the GetList stored procedure
    ''' </summary>
    Private Function IsValidGetListType(ByVal eDataType As SqlDbType) As Boolean
        Select Case eDataType
            Case SqlDbType.Binary, SqlDbType.Image, SqlDbType.Text, SqlDbType.NText,
            SqlDbType.Text, SqlDbType.UniqueIdentifier, SqlDbType.VarBinary, SqlDbType.Variant
                Return False
            Case Else
                Return True
        End Select
    End Function

    ''' <summary>
    ''' Determines whether the datatype uses 'LIKE' or '=' for equality operations
    ''' </summary>
    Private Function DataTypeUsesLike(ByVal eDataType As SqlDbType) As Boolean
        Select Case eDataType
            Case SqlDbType.Char, SqlDbType.NText, SqlDbType.NVarChar, SqlDbType.VarChar
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Function SPFieldDeclaration(ByVal objField As clsField) As String
        Dim strFieldDeclaration As String
        Dim intLength As Integer

        strFieldDeclaration = objField.DataType.ToString.ToUpper

        If Not (objField.DataType = SqlDbType.Money OrElse objField.DataType = SqlDbType.SmallMoney) Then
            intLength = Me.GetFieldLength(objField)

            If objField.Length > 0 AndAlso objField.Scale = clsDBConstants.cintNULL Then
                strFieldDeclaration &= "(" & intLength & ")"
            End If

            If objField.Length > 0 AndAlso objField.Scale >= 0 Then
                strFieldDeclaration &= "(" & intLength & ", " & objField.Scale & ")"
            End If

            If intLength <= 0 AndAlso (objField.DataType = SqlDbType.NVarChar OrElse
            objField.DataType = SqlDbType.VarChar) Then
                strFieldDeclaration &= "(MAX)"
            End If
        End If

        Return strFieldDeclaration
    End Function
#End Region

#Region " Standard Table Stored Procedures "

    ''' <summary>
    ''' Creates the Insert Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateSPInsert(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strFields As String = ""
        Dim strValues As String = ""
        Dim strNextID As String = ""
        Dim strParams As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cINSERT
        DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                SPFieldDeclaration(objField) & " = NULL")

            If objField.IsIdentityField Then
                strParams &= " OUTPUT" 'need to return the id
            Else
                AppendToCommaString(strFields, vbCrLf & vbTab & "[" & objField.DatabaseName & "]")
                AppendToCommaString(strValues, vbCrLf & vbTab & ParamName(objField.DatabaseName))
            End If
        Next

        strSQL = StandardSPDeclaration(strSP, strParams) &
            "INSERT INTO " & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "(" & strFields & vbCrLf & ")" & vbCrLf & "VALUES" & vbCrLf &
            "(" & strValues & vbCrLf & ")" & vbCrLf & vbCrLf &
            "SET @ID = SCOPE_IDENTITY()" & vbCrLf & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the Update Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateSPUpdate(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strUpdate As String = ""
        Dim strParams As String = ""
        Dim strIDParam As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cUPDATE
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            If Not (objField.DataType = SqlDbType.Image) Then
                If objField.IsIdentityField Then
                    strIDParam = ParamName(objField.DatabaseName)
                Else
                    AppendToCommaString(strUpdate, vbCrLf & vbTab & "[" & objField.DatabaseName & "] = " &
                        ParamName(objField.DatabaseName))
                End If

                AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                    SPFieldDeclaration(objField) & " = NULL")
            End If
        Next

        strSQL = StandardSPDeclaration(strSP, strParams) &
            "UPDATE " & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "SET" & strUpdate & vbCrLf &
            "WHERE" & vbCrLf & vbTab & "[" & clsDBConstants.Fields.cID & "] = " & strIDParam & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the GetItem Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateSPGetItem(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strSP As String
        Dim strParams As String = ""
        Dim strWhere As String = ""
        Dim strList As String = ""

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cGETITEM
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            If Not (objField.DataType = SqlDbType.Image) Then
                If objField.IsIdentityField Then
                    strParams = vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " & SPFieldDeclaration(objField)
                    strWhere = "[" & objTable.DatabaseName & "].[" & clsDBConstants.Fields.cID & "] = " &
                        ParamName(objField.DatabaseName)
                End If

                AppendToCommaString(strList, vbCrLf & vbTab & "[" & objTable.DatabaseName & "]." &
                    "[" & objField.DatabaseName & "]")
            End If
        Next

        strSQL = StandardSPDeclaration(strSP, strParams) &
            "SELECT" & strList & vbCrLf &
            "FROM" & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "WHERE" & vbCrLf & vbTab & strWhere & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the GetList Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateSPGetList(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strSelect As String = ""
        Dim strWhere As String = ""
        Dim strParams As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cGETLIST
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            If IsValidGetListType(objField.DataType) Then
                AppendToCommaString(strSelect, vbCrLf & vbTab & "[" & objTable.DatabaseName & "].[" &
                    objField.DatabaseName & "]")
                AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                    SPFieldDeclaration(objField) & " = NULL")
                AppendToString(strWhere, vbTab & "(" & ParamName(objField.DatabaseName) &
                    " IS NULL OR [" & objTable.DatabaseName & "].[" & objField.DatabaseName & "]",
                    vbCrLf & "AND" & vbCrLf)

                If DataTypeUsesLike(objField.DataType) Then
                    strWhere &= " LIKE " & ParamName(objField.DatabaseName) & ")"
                Else
                    strWhere &= " = " & ParamName(objField.DatabaseName) & ")"
                End If
            End If
        Next

        strSQL = StandardSPDeclaration(strSP, strParams) &
            "SELECT" & strSelect & vbCrLf &
            "FROM" & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "WHERE" & vbCrLf & strWhere & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the Delete Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateSPDelete(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cDELETE
        Call DeleteStoredProcedure(strSP)

        strSQL = StandardSPDeclaration(strSP, vbCrLf & vbTab & "@ID INT") &
            "DELETE FROM " & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "WHERE" & vbCrLf & vbTab & "[" & clsDBConstants.Fields.cID & "] = @ID" & vbCrLf

        ExecuteSQL(strSQL)
    End Sub
#End Region

#Region " Link Table Stored Procedures "

    ''' <summary>
    ''' Creates the Insert Stored Procedure for a link table
    ''' </summary>
    Private Sub CreateLinkSPInsert(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strParams As String = Nothing
        Dim strFields As String = ""
        Dim strValues As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cINSERT
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            If Not objField.IsIdentityField Then
                AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                    SPFieldDeclaration(objField) & " = NULL")
                AppendToCommaString(strFields, vbCrLf & vbTab & "[" & objField.DatabaseName & "]")
                AppendToCommaString(strValues, vbCrLf & vbTab & ParamName(objField.DatabaseName))
            End If
        Next

        strSQL = StandardSPDeclaration(strSP, strParams) &
            " INSERT INTO " & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "(" & strFields & vbCrLf & ")" & vbCrLf & " VALUES" & vbCrLf &
            "(" & strValues & vbCrLf & ")" & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the GetList Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateLinkSPGetList(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strSelect As String = ""
        Dim strWhere As String = ""
        Dim strSP As String
        Dim strParams As String = Nothing
        Dim strTables As String = ""
        Dim strAlias As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cGETLIST
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            If IsValidGetListType(objField.DataType) Then

                AppendToCommaString(strSelect, vbCrLf & vbTab & "[" & objTable.DatabaseName & "]." &
                    "[" & objField.DatabaseName & "]")

                If Not objField.IsIdentityField Then

                    AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                        SPFieldDeclaration(objField) & " = NULL")

                    If DataTypeUsesLike(objField.DataType) Then
                        AppendToString(strWhere, vbTab & "(" & ParamName(objField.DatabaseName) &
                            " IS NULL OR [" & objTable.DatabaseName & "].[" & objField.DatabaseName &
                            "] LIKE " & ParamName(objField.DatabaseName) & ")", vbCrLf & "AND" & vbCrLf)
                    Else
                        AppendToString(strWhere, vbTab & "(" & ParamName(objField.DatabaseName) &
                            " IS NULL OR [" & objTable.DatabaseName & "].[" & objField.DatabaseName &
                            "] = " & ParamName(objField.DatabaseName) & ")", vbCrLf & "AND" & vbCrLf)
                    End If
                End If
            End If
        Next

        'For quick access to security and externalIDs of linked table records
        For Each objFieldLink As clsFieldLink In objTable.FieldLinks.Values
            strAlias = objFieldLink.ForeignKeyField.DatabaseName

            AppendToCommaString(strSelect, vbCrLf & vbTab & "[" & strAlias & "].[" & clsDBConstants.Fields.cSECURITYID &
                "] AS " & strAlias & "_" & clsDBConstants.Fields.cSECURITYID)
            AppendToCommaString(strSelect, vbCrLf & vbTab & "[" & strAlias & "].[" & clsDBConstants.Fields.cEXTERNALID &
                "] AS " & strAlias & "_" & clsDBConstants.Fields.cEXTERNALID)
            AppendToString(strWhere, vbTab & "[" & objTable.DatabaseName & "].[" & strAlias &
                "] = [" & strAlias & "].[" & clsDBConstants.Fields.cID & "]", vbCrLf & " AND " & vbCrLf)
            strTables &= ", " & vbCrLf & vbTab & "[" & objFieldLink.IdentityTable.DatabaseName & "] AS [" & strAlias & "]"
        Next

        strSQL &= StandardSPDeclaration(strSP, strParams) &
            "SELECT" & strSelect & vbCrLf &
            "FROM" & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & strTables & vbCrLf &
            "WHERE" & vbCrLf & strWhere & vbCrLf

        ExecuteSQL(strSQL)
    End Sub

    ''' <summary>
    ''' Creates the Delete Stored Procedure for a standard table
    ''' </summary>
    Private Sub CreateLinkSPDelete(ByVal objTable As clsTable)
        Dim strSQL As String = ""
        Dim strParams As String = Nothing
        Dim strWhere As String = ""
        Dim strSP As String

        strSP = objTable.DatabaseName & clsDBConstants.StoredProcedures.cDELETE
        Call DeleteStoredProcedure(strSP)

        For Each objField As clsField In objTable.Fields.Values
            AppendToCommaString(strParams, vbCrLf & vbTab & ParamName(objField.DatabaseName) & " " &
                SPFieldDeclaration(objField) & " = NULL ")
            AppendToString(strWhere, vbTab & "(" & ParamName(objField.DatabaseName) & " IS NULL " &
                "OR [" & objField.DatabaseName & "] = " & ParamName(objField.DatabaseName) & ")" & vbCrLf,
                " AND" & vbCrLf)
        Next

        strSQL &= StandardSPDeclaration(strSP, strParams) &
            "DELETE FROM " & vbCrLf & vbTab & "[" & objTable.DatabaseName & "]" & vbCrLf &
            "WHERE" & vbCrLf & strWhere

        ExecuteSQL(strSQL)
    End Sub
#End Region

#End Region

#Region " Records "

    Public Overloads Sub DeleteRecord(ByVal strTable As String, ByVal strField As String, ByVal intID As Integer)
        Try
            ExecuteSQL("DELETE FROM [" & strTable & "] " &
                    "WHERE [" & strField & "] = (" & intID & ")")
        Catch
            Throw
        End Try
    End Sub

    Public Overrides Sub DeleteRecord(ByVal strTable As String, ByVal intID As Integer)
        Try
            DeleteRecord(strTable, clsDBConstants.Fields.cID, intID)
        Catch
            Throw
        End Try
    End Sub

    Public Sub DeleteRecordRange(ByVal strTable As String,
                                 ByVal strField As String,
                                 ByVal arrIDs() As Integer)
        If arrIDs.GetUpperBound(0) >= 0 Then
            Dim strIDs As String = ""

            For intIndex As Integer = 0 To arrIDs.GetUpperBound(0)
                modGlobal.AppendToCommaString(strIDs, arrIDs(intIndex).ToString)
            Next

            DeleteRecordRange(strTable, strField, strIDs)
        End If

    End Sub

    Public Sub DeleteRecordRange(ByVal strTable As String,
                                 ByVal strField As String,
                                 ByVal strIDs As String,
                                 Optional blnCreateAuditTrailRecord As Boolean = True)

        Dim blnCreatedTransaction As Boolean = False
        Try
            If strIDs IsNot Nothing AndAlso strIDs.Length > 0 Then
                If Not HasTransaction Then
                    BeginTransaction()
                    blnCreatedTransaction = True
                End If

                ExecuteSQL("DELETE FROM [" & strTable & "] " & "WHERE [" & strField & "] IN (" & strIDs & ")")

                Dim objTable As clsTable = Me.SysInfo.Tables(strTable)

                If blnCreateAuditTrailRecord AndAlso objTable IsNot Nothing Then

                    Dim objDataTable As DataTable = CreateDataTableFromIDString(strIDs)

                    clsAuditTrail.CreateAuditTrailRecords(Me, clsMethod.enumMethods.cDELETE,
                                                          objTable,
                                                          objDataTable)

                End If

            End If
            If blnCreatedTransaction Then
                EndTransaction(True)
            End If
        Catch
            If blnCreatedTransaction Then
                EndTransaction(False)
            End If
            Throw
        End Try

    End Sub

    Public Sub ClearRecordKey(ByVal strTable As String,
                              ByVal strClearField As String,
                              ByVal strIDField As String,
                              ByVal strIDs As String,
                              Optional blnCreateAuditTrailRecord As Boolean = True)
        Try
            If strIDs IsNot Nothing AndAlso strIDs.Length > 0 Then
                ExecuteSQL("UPDATE [" & strTable & "] " &
                    "SET [" & strClearField & "] = NULL " &
                    "WHERE [" & strIDField & "] IN (" & strIDs & ")")

                Dim objTable As clsTable = Me.SysInfo.Tables(strTable)

                If blnCreateAuditTrailRecord AndAlso objTable IsNot Nothing Then
                    Dim objDT As DataTable = CreateDataTableFromIDString(strIDs)
                    clsAuditTrail.CreateAuditTrailRecords(Me,
                                                          clsMethod.enumMethods.cMODIFY,
                                                          objTable,
                                                          objDT)
                End If

            End If
        Catch
            Throw
        End Try

    End Sub
#End Region

#Region " Database Information Methods "

    ''' <summary>
    ''' Returns information about the current database.
    ''' Returned Columns:
    ''' ServerName - Name of the Server the database is on
    ''' DatabaseName - Name of database
    ''' CreationDate - Date the database was created
    ''' SpaceUsedMB - Amount of space in MegaBytes that the database is using on the file system
    ''' FreeSpaceMB - Amount of space in MegaBytes unused by the database
    ''' DataFileName - File name of the data file for the database
    ''' LastBackupDate - The date the database was last backed up
    ''' BackupSizeMB - The size of the backup file it when it was last backed up
    ''' DurationInSec -  The total time in seconds it took to back up the database
    ''' BackupFileName - The physical device name of the last backup file
    ''' </summary>
    Public Function GetDatabaseDetails() As DataTable
        Try
            Dim strSQL As String
            Const cDATABASE_NAME As String = "DatabaseName"
            Const cSERVER_NAME As String = "ServerName"

            strSQL = "SELECT " & vbCrLf &
                        vbTab & "sd.name AS DatabaseName, " & vbCrLf &
                        vbTab & "database_creation_date AS CreationDate, " & vbCrLf &
                        vbTab & "CAST(FILEPROPERTY(sf.name, 'SpaceUsed')AS int)/128.0 AS SpaceUsedMB, " & vbCrLf &
                        vbTab & "sf.size/128.0 - CAST(FILEPROPERTY(sf.name, 'SpaceUsed') AS int)/128.0 AS FreeSpaceMB, " & vbCrLf &
                        vbTab & "sd.filename AS DataFileName, " & vbCrLf &
                        vbTab & "backup_finish_date AS LastBackupDate, " & vbCrLf &
                        vbTab & "backup_size/1024/1024 AS BackupSizeMB, " & vbCrLf &
                        vbTab & "DATEDIFF(s, backup_start_date, backup_finish_date) AS DurationInSec, " & vbCrLf &
                        vbTab & "bm.physical_device_name AS BackupFileName, " & vbCrLf &
                        vbTab & "sf.name as LogicalName " & vbCrLf &
                      "FROM " & vbCrLf &
                        vbTab & "master.dbo.sysdatabases as sd " & vbCrLf &
                      "LEFT JOIN " & vbCrLf &
                        vbTab & "msdb.dbo.backupset as bs1 " & vbCrLf &
                        vbTab & vbTab & "ON " & vbCrLf &
                        vbTab & vbTab & vbTab & "bs1.database_name = sd.name " & vbCrLf &
                        vbTab & vbTab & "AND " & vbCrLf &
                        vbTab & vbTab & vbTab & "bs1.backup_finish_date = (SELECT MAX(backup_finish_date) FROM msdb.dbo.backupset as bs2 WHERE bs1.database_name = bs2.database_name) " & vbCrLf &
                      "LEFT JOIN " & vbCrLf &
                        vbTab & "msdb.dbo.backupmediafamily bm ON bm.media_set_id = bs1.media_set_id " & vbCrLf &
                      "LEFT JOIN " & vbCrLf &
                        vbTab & "dbo.SYSFILES as sf ON sd.filename COLLATE Latin1_General_CI_AI = sf.filename COLLATE Latin1_General_CI_AI " & vbCrLf &
                      "WHERE " & vbCrLf &
                        vbTab & "sd.name = @DatabaseName "

            '-- Set the DatabaseName Parameter
            Dim colDBParameters As New clsDBParameterDictionary
            colDBParameters.Add(New clsDBParameter(ParamName(cDATABASE_NAME), m_strDatabase, ParameterDirection.Input, SqlDbType.NVarChar))

            '-- Get the details
            Dim objDTDetails As DataTable = GetDataTableBySQL(strSQL, colDBParameters)

            If objDTDetails IsNot Nothing AndAlso objDTDetails.Rows.Count > 0 Then
                objDTDetails.Columns.Add(cSERVER_NAME, GetType(String))
                objDTDetails.Rows(0)(cSERVER_NAME) = m_strServer

                Return objDTDetails
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Checks if the SQL Notification Broker is enabled for the database.
    ''' </summary>
    Public Function IsBrokerEnabled() As Boolean
        Try
            Me.BeginTransaction(IsolationLevel.ReadUncommitted)

            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@Name", m_strDatabase))

            Dim strSQL As String = "SELECT is_broker_enabled FROM sys.databases WHERE name = @Name"
            Dim blnEnabled As Boolean = CBool(Me.GetColumnBySQL(strSQL, colParams))
            colParams.Dispose()

            Me.EndTransaction(True)

            Return blnEnabled
        Catch ex As Exception
            Me.EndTransaction(False)

            Throw
        End Try
    End Function

#End Region

#Region " Convert To Stored Procedures "

    Protected Function GetFieldLength(ByVal objField As clsField) As Integer
        Dim intLength As Integer = objField.Length

        If objField.IsEncrypted Then
            Dim objEncryption As New clsEncryption(True)
            Dim sPassword As String = objEncryption.Encrypt("".PadRight(intLength, "z"c))
            intLength = sPassword.Length
        End If

        Return intLength
    End Function

    Public Function GetNextErrorMessageUIID() As Integer
        Dim strSQL As String = "SELECT ISNULL(Max([{1}]), 0) FROM [{0}] WHERE [{1}] < 10000 OR [{1}] > 20000"

        Dim intUIID As Integer = ExecuteScalar(String.Format(strSQL, clsDBConstants.Tables.cERRORMESSAGE,
                                                             clsDBConstants.Fields.ErrorMessage.cUIID))
        intUIID += 1

        Return intUIID
    End Function

    Public Function GetUserProfileCount(ByVal objSG As clsSecurityGroup) As Integer
        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@ID", objSG.ID))

        '2016-08-02 -- Peter Melisi -- Bug fix for #1600003149 -- Use correct table for new multiple security groups.
        Dim strSQL As String = "SELECT Count([" & clsDBConstants.Fields.cID &
            "]) As NumUsers FROM [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "] WHERE " &
            clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID & " = @ID"

        Dim objDT As DataTable = GetDataTableBySQL(strSQL, colParams)
        colParams.Dispose()

        If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
            Return 0
        Else
            Return CType(objDT.Rows(0)(0), Integer)
        End If
    End Function

#Region " Language Functions "

    Private Function GetLanguageSQLBase(ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = "SELECT" & vbCrLf &
            "LS1.[" & clsDBConstants.Fields.cID & "] as LanguageString1ID," & vbCrLf &
            "CAST(" & clsDBConstants.Fields.LanguageString.cSTRING & " AS NVARCHAR(100)) AS LanguageText1," & vbCrLf &
            "(SELECT LS2.[" & clsDBConstants.Fields.cID & "] " &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS2 " &
            "WHERE LS2." & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage2ID & " " &
            "AND LS2." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & ") " &
            "AS LanguageString2ID, " & vbCrLf &
            "(SELECT CAST(" & clsDBConstants.Fields.LanguageString.cSTRING & " AS NVARCHAR(100)) " &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS2 " &
            "WHERE LS2." & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage2ID & " " &
            "AND LS2." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & ") " &
            "AS LanguageText2"

        Return strSQL
    End Function

    Private Function GetLanguageSQLForButton(ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = GetLanguageSQLBase(intLanguage1ID, intLanguage2ID)

        strSQL &= "," & vbCrLf & "'" & clsDBConstants.enumLanguageFilters.Buttons.ToString & "' AS LanguageFilter," & vbCrLf &
            clsDBConstants.Tables.cBUTTON & "." & clsDBConstants.Fields.cEXTERNALID & " AS Item" & vbCrLf &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS1, " &
            clsDBConstants.Tables.cSTRING & ", " &
            clsDBConstants.Tables.cBUTTON & vbCrLf &
            "WHERE" & vbCrLf &
            "LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & clsDBConstants.Tables.cBUTTON & "." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND " & vbCrLf & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage1ID

        Return strSQL
    End Function

    Private Function GetLanguageSQLForDefaultNoCaption(ByVal strTable As String,
    ByVal eFilter As clsDBConstants.enumLanguageFilters,
    ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = GetLanguageSQLBase(intLanguage1ID, intLanguage2ID)

        strSQL &= "," & vbCrLf & "'" & eFilter.ToString & "' AS LanguageFilter," & vbCrLf &
            strTable & "." & clsDBConstants.Fields.cEXTERNALID & " AS Item" & vbCrLf &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS1, " &
            clsDBConstants.Tables.cSTRING & ", [" & strTable & "]" & vbCrLf &
            "WHERE" & vbCrLf &
            "LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & strTable & "." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND " & vbCrLf & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage1ID

        Return strSQL
    End Function

    Private Function GetLanguageSQLForDefault(ByVal strTable As String,
    ByVal eFilter As clsDBConstants.enumLanguageFilters,
    ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = GetLanguageSQLBase(intLanguage1ID, intLanguage2ID)

        strSQL &= "," & vbCrLf & "'" & eFilter.ToString & "' AS LanguageFilter," & vbCrLf &
            "[" & strTable & "]." & clsDBConstants.Fields.cEXTERNALID & " AS Item" & vbCrLf &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS1, " &
            clsDBConstants.Tables.cSTRING & ", " &
            clsDBConstants.Tables.cCAPTION & ", [" &
            strTable & "]" & vbCrLf &
            "WHERE" & vbCrLf &
            "LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & clsDBConstants.Tables.cCAPTION & "." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & "[" & strTable & "]." & clsDBConstants.Fields.String.cCAPTIONID & " = " &
            clsDBConstants.Tables.cCAPTION & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND " & vbCrLf & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage1ID

        Return strSQL
    End Function

    Private Function GetLanguageSQLForFieldColumns(ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = GetLanguageSQLBase(intLanguage1ID, intLanguage2ID)

        strSQL &= "," & vbCrLf & "'" & clsDBConstants.enumLanguageFilters.FieldColumnHeadings.ToString & "' AS LanguageFilter," & vbCrLf &
            clsDBConstants.Tables.cLISTCOLUMN & "." & clsDBConstants.Fields.cEXTERNALID & " AS Item" & vbCrLf &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS1, " &
            clsDBConstants.Tables.cSTRING & ", " &
            clsDBConstants.Tables.cCAPTION & ", " &
            clsDBConstants.Tables.cLISTCOLUMN & vbCrLf &
            "WHERE" & vbCrLf &
            "LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & clsDBConstants.Tables.cCAPTION & "." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND" & vbCrLf & clsDBConstants.Tables.cLISTCOLUMN & "." & clsDBConstants.Fields.ListColumn.cCAPTIONID & " = " &
            clsDBConstants.Tables.cCAPTION & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "AND " & vbCrLf & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage1ID

        Return strSQL
    End Function

    Private Function GetLanguageSQLForLabels(ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = GetLanguageSQLBase(intLanguage1ID, intLanguage2ID)

        strSQL &= "," & vbCrLf & "'" & clsDBConstants.enumLanguageFilters.LabelsAndMessages.ToString & "' AS LanguageFilter," & vbCrLf &
            clsDBConstants.Tables.cCAPTION & "." & clsDBConstants.Fields.cEXTERNALID & " AS Item" & vbCrLf &
            "FROM " & clsDBConstants.Tables.cLANGUAGESTRING & " LS1" & vbCrLf &
            "INNER JOIN " & clsDBConstants.Tables.cSTRING & " ON " &
            "LS1." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "INNER JOIN " & clsDBConstants.Tables.cCAPTION & " ON " &
            clsDBConstants.Tables.cCAPTION & "." & clsDBConstants.Fields.LanguageString.cSTRINGID & " = " &
            clsDBConstants.Tables.cSTRING & ".[" & clsDBConstants.Fields.cID & "]" & vbCrLf &
            "WHERE" & vbCrLf & clsDBConstants.Fields.LanguageString.cLANGUAGEID & " = " & intLanguage1ID

        Return strSQL
    End Function

    Public Function GetLanguageSQL(ByVal eFilter As clsDBConstants.enumLanguageFilters,
    ByVal intLanguage1ID As Integer, ByVal intLanguage2ID As Integer) As String
        Dim strSQL As String = Nothing

        Select Case eFilter
            Case clsDBConstants.enumLanguageFilters.NoFilter
                strSQL = GetLanguageSQLForButton(intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefaultNoCaption(clsDBConstants.Tables.cERRORMESSAGE,
                        clsDBConstants.enumLanguageFilters.ErrorMessages, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefault(clsDBConstants.Tables.cFIELD,
                        clsDBConstants.enumLanguageFilters.Fields, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForFieldColumns(intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefault(clsDBConstants.Tables.cFIELDLINK,
                        clsDBConstants.enumLanguageFilters.FieldLinks, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefault(clsDBConstants.Tables.cTABLE,
                        clsDBConstants.enumLanguageFilters.Tables, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefault(clsDBConstants.Tables.cTYPEFIELDINFO,
                        clsDBConstants.enumLanguageFilters.TypeDependentFields, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefault(clsDBConstants.Tables.cTYPEFIELDLINKINFO,
                        clsDBConstants.enumLanguageFilters.TypeDependentFieldLinks, intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForLabels(intLanguage1ID, intLanguage2ID) & vbCrLf &
                    "UNION" & vbCrLf & GetLanguageSQLForDefaultNoCaption(clsDBConstants.Tables.cAPPLICATIONMETHOD,
                        clsDBConstants.enumLanguageFilters.ApplicationMethods, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.Buttons
                strSQL = GetLanguageSQLForButton(intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.ErrorMessages
                strSQL = GetLanguageSQLForDefaultNoCaption(clsDBConstants.Tables.cERRORMESSAGE,
                    clsDBConstants.enumLanguageFilters.ErrorMessages, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.Fields
                strSQL = GetLanguageSQLForDefault(clsDBConstants.Tables.cFIELD,
                    clsDBConstants.enumLanguageFilters.Fields, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.FieldColumnHeadings
                strSQL = GetLanguageSQLForFieldColumns(intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.FieldLinks
                strSQL = GetLanguageSQLForDefault(clsDBConstants.Tables.cFIELDLINK,
                    clsDBConstants.enumLanguageFilters.FieldLinks, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.Tables
                strSQL = GetLanguageSQLForDefault(clsDBConstants.Tables.cTABLE,
                    clsDBConstants.enumLanguageFilters.Tables, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.TypeDependentFields
                strSQL = GetLanguageSQLForDefault(clsDBConstants.Tables.cTYPEFIELDINFO,
                    clsDBConstants.enumLanguageFilters.TypeDependentFields, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.TypeDependentFieldLinks
                strSQL = GetLanguageSQLForDefault(clsDBConstants.Tables.cTYPEFIELDLINKINFO,
                    clsDBConstants.enumLanguageFilters.TypeDependentFieldLinks, intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.LabelsAndMessages
                strSQL = GetLanguageSQLForLabels(intLanguage1ID, intLanguage2ID)
            Case clsDBConstants.enumLanguageFilters.ApplicationMethods
                strSQL = GetLanguageSQLForDefaultNoCaption(clsDBConstants.Tables.cAPPLICATIONMETHOD,
                clsDBConstants.enumLanguageFilters.ApplicationMethods, intLanguage1ID, intLanguage2ID)
        End Select

        Return strSQL
    End Function
#End Region

    Public Function GetProfileEmailsDataTable(ByVal colProfiles As Hashtable) As DataTable
        Try
            Dim strIDs As String = ""
            For Each intID As Integer In colProfiles.Values
                If strIDs.Length > 0 Then
                    strIDs &= ","
                End If
                strIDs &= intID
            Next

            Dim strSQL As String = "SELECT " & clsDBConstants.Tables.cPERSON & "." &
                clsDBConstants.Fields.Person.cWORKEMAIL & " FROM " &
                clsDBConstants.Tables.cPERSON & ", " & clsDBConstants.Tables.cUSERPROFILE & " " &
                "WHERE " & clsDBConstants.Tables.cUSERPROFILE & "." & clsDBConstants.Fields.UserProfile.cPERSONID & " = " &
                clsDBConstants.Tables.cPERSON & ".[" & clsDBConstants.Fields.cID & "] AND " &
                clsDBConstants.Tables.cUSERPROFILE & ".[" & clsDBConstants.Fields.cID & "] IN " &
                "(" & strIDs & ")"

            Return GetDataTableBySQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Full text Indexing "

    Public Sub FullTextIndexAddColumn(ByVal strTable As String, ByVal strField As String)
        'Dim colParams As New clsDBParameterDictionary

        'colParams.Add(New clsDBParameter("@tabname", strTable))
        'colParams.Add(New clsDBParameter("@colname", strField))
        'colParams.Add(New clsDBParameter("@action", "add"))
        'colParams.Add(New clsDBParameter("@language", 0))

        'ExecuteProcedure("sp_fulltext_column", colParams)
        ExecuteSQL(String.Format("ALTER FULLTEXT INDEX ON [dbo].[{0}] ADD ([{1}] LANGUAGE 0)",
                                  strTable, strField))
    End Sub

    Public Sub FullTextIndexDeleteColumn(ByVal strTable As String, ByVal strField As String)
        'Dim colParams As New clsDBParameterDictionary

        'colParams.Add(New clsDBParameter("@tabname", strTable))
        'colParams.Add(New clsDBParameter("@colname", strField))
        'colParams.Add(New clsDBParameter("@action", "drop"))

        'ExecuteProcedure("sp_fulltext_column", colParams)
        ExecuteSQL(String.Format("ALTER FULLTEXT INDEX ON [dbo].[{0}] DROP ([{1}])",
                                  strTable, strField))
    End Sub

    Public Sub FullTextReIndex(ByVal strIndexName As String)
        'Dim colParams As New clsDBParameterDictionary

        'colParams.Add(New clsDBParameter("@ftcat", strIndexName))
        'colParams.Add(New clsDBParameter("@action", "rebuild"))

        'ExecuteProcedure("sp_fulltext_catalog", colParams)
        ExecuteSQL(String.Format("ALTER FULLTEXT CATALOG [{0}] REBUILD", strIndexName))
    End Sub

    Public Function FullTextIndexHasColumn(ByVal strTable As String, ByVal strField As String) As Boolean
        Dim strSQL As String = "SELECT 1 FROM sys.fulltext_index_columns  as fic" & vbCrLf &
                               "INNER JOIN sys.columns as c ON fic.object_id = c.object_id AND fic.column_id = c.column_id" & vbCrLf &
                               "WHERE object_name(fic.object_id) = '{0}' AND c.name = '{1}'"

        Return CBool(ExecuteScalar(String.Format(strSQL, strTable, strField)))
    End Function

    Public Sub RebuildEdocView(Optional ByVal arrExcludeFields As FrameworkCollections.K1Collection(Of String) = Nothing)
        Dim objTable As clsTable = Me.SysInfo.Tables(clsDBConstants.Tables.cEDOC)
        Dim strFields As String = Nothing
        Dim strSQL As String = Nothing

        If arrExcludeFields Is Nothing Then
            arrExcludeFields = New FrameworkCollections.K1Collection(Of String)
        End If

        For Each objField As clsField In objTable.Fields.Values
            If (objField.IsTextType OrElse objField.IsIdentityField _
                OrElse objField.DatabaseName = clsDBConstants.Fields.cSECURITYID) _
            AndAlso Not arrExcludeFields.Contains(objField.DatabaseName) Then
                AppendToCommaString(strFields, String.Format("[{0}]", objField.DatabaseName))
            End If
        Next

        If ViewExists(clsDBConstants.Views.cEDOCIndexView) Then
            strSQL = "ALTER VIEW [dbo].[{0}]" & vbCrLf &
                     "WITH SCHEMABINDING" & vbCrLf &
                     "AS" & vbCrLf &
                     "SELECT {1}" & vbCrLf &
                     "FROM [dbo].[{2}]"
        Else
            strSQL = "CREATE VIEW [dbo].[{0}]" & vbCrLf &
                     "WITH SCHEMABINDING" & vbCrLf &
                     "AS" & vbCrLf &
                     "SELECT {1}" & vbCrLf &
                     "FROM [dbo].[{2}]"
        End If

        ExecuteSQL(String.Format(strSQL, clsDBConstants.Views.cEDOCIndexView, strFields,
                                         clsDBConstants.Tables.cEDOC))

        '--
        strSQL = "CREATE UNIQUE CLUSTERED INDEX [PK_{0}] ON [dbo].[{0}](" & vbCrLf &
                 "[ID] Asc" & vbCrLf &
                 ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"

        ExecuteSQL(String.Format(strSQL, clsDBConstants.Views.cEDOCIndexView))

        '--
        strSQL = "CREATE FULLTEXT INDEX ON [dbo].[{0}] KEY INDEX [PK_{0}] ON [RecFind6Index] WITH CHANGE_TRACKING AUTO"

        ExecuteSQL(String.Format(strSQL, clsDBConstants.Views.cEDOCIndexView))

        '--
        strFields = Nothing
        For Each objField As clsField In objTable.Fields.Values
            If objField.IsTextType AndAlso Not arrExcludeFields.Contains(objField.DatabaseName) Then
                AppendToCommaString(strFields, String.Format(vbCrLf & "[{0}] LANGUAGE 0", objField.DatabaseName))
            End If
        Next

        strSQL = "ALTER FULLTEXT INDEX ON [dbo].[{0}] ADD (" & vbCrLf &
                 strFields & vbCrLf &
                 ")"

        ExecuteSQL(String.Format(strSQL, clsDBConstants.Views.cEDOCIndexView))

        '--
        strSQL = "ALTER FULLTEXT INDEX ON [dbo].[{0}] ENABLE"

        ExecuteSQL(String.Format(strSQL, clsDBConstants.Views.cEDOCIndexView))
    End Sub

    Public Function GetFullTextInfo(ByVal strTableName As String) As Object
        Dim strSQL As String = "SELECT OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFullTextBackgroundUpdateIndexOn') as BackgroundUpdateIndexOn," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextPendingChanges') as PendingChanges," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextPopulateStatus') as PopulateStatus," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableHasActiveFulltextIndex') as HasActiveFulltextIndex," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextCatalogId') as CatalogId," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFullTextChangeTrackingOn') as ChangeTrackingOn," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextDocsProcessed') as DocsProcessed," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextFailCount') as FailCount," &
                                "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextItemCount') as ItemCount"

        Dim objDT As DataTable = GetDataTableBySQL(String.Format(strSQL, strTableName))
        Dim info = New With {.Name = ""}

        Return info
    End Function
#End Region

#Region " Collation "

    Public Function GetFieldsWithInvalidCollation() As DataTable
        Dim strQuery As String = "SELECT OBJECT_NAME(c.object_ID) as tableName, c.name, c.max_length, " &
                                    "c.scale, c.is_nullable, dc.definition, t.name as type " &
                                 "FROM sys.columns as c " &
                                 "LEFT JOIN sys.default_constraints as dc ON dc.object_id = c.default_object_id " &
                                 "INNER JOIN sys.types as t ON t.user_type_id = c.user_type_id " &
                                 "WHERE c.collation_name is not NULL AND " &
                                    "NOT c.collation_name = (SELECT CAST(DATABASEPROPERTYEX('{0}', 'Collation') AS NVARCHAR)) AND " &
                                    "c.object_id IN (SELECT object_ID FROM sys.tables where is_ms_shipped=0 AND type=N'U') " &
                                 "ORDER BY tablename, c.name"

        Return Me.GetDataTableBySQL(String.Format(strQuery, m_strDatabase))
    End Function

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                MyBase.Dispose(blnDisposing)

                m_objSecurity = Nothing
            End If
        End If
    End Sub
#End Region

End Class

