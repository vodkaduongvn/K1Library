Imports System.Data.SqlClient
''' <summary>
''' Represents a connection to the data layer (either direct database or via web services)
''' </summary>
Public MustInherit Class clsDB
    Implements IDisposable

#Region " Members "

    Protected m_eDataAccess As enumDataAccessType
    Protected m_objProfile As clsUserProfile
    Protected m_objSysInfo As clsSysInfo
    Protected m_intRecordLimit As Integer = clsDBConstants.cintNULL
    Protected m_blnDisposedValue As Boolean = False     ' To detect redundant calls
    Protected m_intBlobChunkSize As Integer = 524288    ' 512K (Matches our default WS chunk size)
    Protected m_blnCancel As Boolean = False
    Protected m_colNotifications As New K1Library.FrameworkCollections.K1Dictionary(Of SqlClient.OnChangeEventHandler)
    Protected m_objSQLDependency As clsSqlDependency
    Protected m_objSysInfoReload As clsSysInfo
    Protected m_dtLastRefreshed As DateTime
    Protected m_eState As enumState
    Protected m_blnMultiThreaded As Boolean = True
    Protected m_objSession As clsSession
    Protected m_IdentityProvider As clsIdentityProvider
#End Region

#Region " Enumerations "

    Public Enum enumDataAccessType
        DB_DIRECT = 1
        WEB_SERVICE = 2
        WEB_SESSION = 3
        AUDIT_TRAIL = 4
    End Enum

    Public Enum enumTransferType
        UPLOAD = 0
        DOWNLOAD = 1
    End Enum

    Public Enum enumState
        NORMAL = 0
        DRM_LOCKED_SESSION = 1
    End Enum

#End Region

#Region " Properties "

    ''' <summary>
    ''' Each data access object should be linked to exactly one user profile
    ''' </summary>
    Public Overridable Property Profile() As clsUserProfile
        Get
            Return m_objProfile
        End Get
        Set(ByVal value As clsUserProfile)
            m_objProfile = value
        End Set
    End Property

    ''' <summary>
    ''' Holds in memory information regarding database objects
    ''' </summary>
    Public Overridable Property SysInfo() As clsSysInfo
        Get
            If m_objSysInfoReload IsNot Nothing Then
                m_objSysInfo = m_objSysInfoReload
                m_objSysInfoReload = Nothing
            End If
            Return m_objSysInfo
        End Get
        Set(ByVal value As clsSysInfo)
            m_objSysInfo = value
        End Set
    End Property

    ''' <summary>
    ''' If K1 is setup as a training version, this is record limit associated
    ''' </summary>
    Public Overridable Property RecordLimit() As Integer
        Get
            Return m_intRecordLimit
        End Get
        Set(ByVal value As Integer)
            m_intRecordLimit = value
        End Set
    End Property

    ''' <summary>
    ''' Flags whether K1 is licensed as a training version only
    ''' </summary>
    Public Overridable ReadOnly Property IsTrainingVersion() As Boolean
        Get
            If RecordLimit = clsDBConstants.cintNULL Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    ''' <summary>
    ''' This flags which type of Data Access Type (Direct or Web Service) this object is
    ''' </summary>
    Public Overridable ReadOnly Property DataAccessType() As enumDataAccessType
        Get
            Return m_eDataAccess
        End Get
    End Property

    ''' <summary>
    ''' When reading/writing bytes to or from a blob field, this is the size of the chunks used
    ''' </summary>
    Public Overridable Property BlobChunkSize() As Integer
        Get
            Return m_intBlobChunkSize
        End Get
        Set(ByVal value As Integer)
            m_intBlobChunkSize = value
        End Set
    End Property

    Public Overridable Property ThreadedOperationCancelled() As Boolean
        Get
            Return m_blnCancel
        End Get
        Set(ByVal value As Boolean)
            m_blnCancel = value
        End Set
    End Property

    Public Overridable Property SqlDependency() As clsSqlDependency
        Get
            Return m_objSQLDependency
        End Get
        Set(ByVal value As clsSqlDependency)
            m_objSQLDependency = value
        End Set
    End Property

    '[Naing] We should use this to allow applications to track multiple sql data notifications
    Protected objSqlNotifications As IList(Of clsSqlNotifier)
    Public Overridable ReadOnly Property SqlNotifications() As IList(Of clsSqlNotifier)
        Get
            If (objSqlNotifications Is Nothing) Then
                objSqlNotifications = New List(Of clsSqlNotifier)
            End If
            Return objSqlNotifications
        End Get
    End Property

    Public Overridable Property LastRefresh() As DateTime
        Get
            Return m_dtLastRefreshed
        End Get
        Set(ByVal value As DateTime)
            m_dtLastRefreshed = value
        End Set
    End Property

    Public Overridable Property Session() As clsSession
        Get
            Return m_objSession
        End Get
        Set(ByVal value As clsSession)
            m_objSession = value
        End Set
    End Property

    Public Overridable Property State() As enumState
        Get
            Return m_eState
        End Get
        Set(ByVal value As enumState)
            m_eState = value
            'RaiseEvent StateChanged()
            DBStateChanged()
        End Set
    End Property

    Public Overridable ReadOnly Property IdentityProvider As clsIdentityProvider
        Get
            If m_IdentityProvider Is Nothing Then
                m_IdentityProvider = New clsIdentityProvider()
            End If
            Return m_IdentityProvider
        End Get
    End Property

    Public MustOverride ReadOnly Property HasTransaction() As Boolean
#End Region

#Region " Events "

    Public Event SessionUpdated()
    Public Event FileTransferInit(ByVal eType As enumDataAccessType,
        ByVal intTotalChunks As Integer, ByVal eTransferType As enumTransferType)
    Public Event FileTransferStep()
    Public Event FileTransferEnd()

#End Region

#Region " Methods "

#Region " Transactions "

    ''' <summary>
    ''' Creates a new transaction for the connection using the default isolation type (Snapshot)
    ''' </summary>
    Public MustOverride Sub BeginTransaction()

    ''' <summary>
    ''' Creates a new transaction for the connection using the isolation type specified
    ''' </summary>
    Public MustOverride Sub BeginTransaction(ByVal eIsolationLevel As System.Data.IsolationLevel)

    ''' <summary>
    ''' Ends any transactions which exist on open connections
    ''' </summary>
    ''' <param name="blnCommit">True - Commit transaction, False - Rollback transaction</param>
    Public MustOverride Sub EndTransaction(ByVal blnCommit As Boolean)
#End Region

#Region " General Methods "

    ''' <summary>
    ''' Returns a single item record for an object
    ''' </summary>
    Public Shared Function DataRowValue(ByVal objDR As DataRow,
    ByVal strKey As String, ByVal objNullValue As Object) As Object
        Dim objReturn As Object

        If objDR.Table.Columns.Item(strKey) Is Nothing Then
            Return objNullValue
        Else
            objReturn = objDR.Item(strKey)
        End If

        Return NullValue(objReturn, objNullValue)
    End Function

    ''' <summary>
    ''' Returns a single TType value from the record.
    ''' </summary>
    Public Shared Function DataRowValue(Of TType)(ByVal objDR As DataRow,
    ByVal strKey As String, ByVal objNullValue As Object) As TType
        Return CType(DataRowValue(objDR, strKey, objNullValue), TType)
    End Function

    ''' <summary>
    ''' If objValue is null, objNull is returned, else objValue is returned
    ''' </summary>
    Public Shared Function NullValue(ByVal objValue As Object, ByVal objNull As Object) As Object
        If IsDBNull(objValue) Then
            Return objNull
        Else
            Return objValue
        End If
    End Function

    ''' <summary>
    ''' Returns a null value if the object is equal to its null equivalent
    ''' </summary>
    Public Shared Function ParamNull(ByVal objValue As Object, ByVal objCompare As Object) As Object
        If objValue.Equals(objCompare) Then
            Return System.DBNull.Value
        Else
            Return objValue
        End If
    End Function

    ''' <summary>
    ''' Replaces all string marker characters in a Sql Text Value
    ''' </summary>
    Public Shared Function SQLString(ByVal strText As String) As String
        Return strText.Replace("'", "''")
    End Function

    ''' <summary>
    ''' Returns the Database Version of the K1 database
    ''' </summary>
    Public Function GetDatabaseVersion() As Double
        Dim strSQL As String = "SELECT [" & clsDBConstants.Fields.K1Configuration.cDATABASEVERSION &
            "] FROM [" & clsDBConstants.Tables.cK1CONFIGURATION & "]"

        Dim objDT As DataTable = GetDataTableBySQL(strSQL)
        If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
            Return CDbl(objDT.Rows(0)(0))
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' 
    Public Function CanWorkWithDatabaseVersion(ByVal dblExpectedVersion As Double,
    ByVal blnAnyDbVersionHigher As Boolean) As Boolean
        Dim dblActualVersion As Double
        Dim intMajorActual As Integer
        Dim intMajorExpected As Integer

        dblActualVersion = GetDatabaseVersion()
        intMajorActual = CInt(Math.Floor(dblActualVersion))
        intMajorExpected = CInt(Math.Floor(dblExpectedVersion))

        If (blnAnyDbVersionHigher OrElse intMajorActual = intMajorExpected) AndAlso
        dblExpectedVersion <= dblActualVersion Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Returns the Availability Group Name if one exists
    ''' </summary>
    Public Function GetAvailabilityGroupName() As String
        Const strSQL =
            "select sys.availability_group_listeners.dns_name from sys.availability_group_listeners 
                join sys.availability_databases_cluster on sys.availability_group_listeners.group_id = sys.availability_databases_cluster.group_id
                where sys.availability_databases_cluster.database_name = db_name()"
        Dim strResponse As String = String.Empty

        Try
            Dim objDT As DataTable = GetDataTableBySQL(strSQL)
            If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
                strResponse = CStr(objDT.Rows(0)(0))
            End If
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine($"GetAvailabilityGroupName Exception {ex.ToString}")
        Finally
#If TRACE Then
            Trace.WriteVerbose($"(ALWAYS ON GROUP QUERY) {strSQL} Response : {strResponse}", "")
#End If
        End Try

        Return strResponse
    End Function

    ''' <summary>
    ''' Converts a field name to a SQL Server Stored Procedure parameter name
    ''' </summary>
    Public Shared Function ParamName(ByVal strField As String) As String
        Return "@" & strField
    End Function

    ''' <summary>
    ''' When using the SQL LIKE statement certain characters need to be prefixed with an escape character
    ''' because they mean something specific to the LIKE function
    ''' </summary>
    Public Shared Function EscapeWildcardChars(ByVal strValue As String,
                                               ByVal blnPagingFilter As Boolean) As String
        Dim blnFoundEscapeChar As Boolean = False
        Dim arrEscapeChar() As Char = {"#"c, "~"c, "!"c, "@"c, "$"c, "^"c, "&"c, "-"c}
        Dim intLoop As Integer

        For intLoop = 0 To arrEscapeChar.GetUpperBound(0)
            If strValue.IndexOf(arrEscapeChar(intLoop)) < 0 Then
                blnFoundEscapeChar = True
                Exit For
            End If
        Next

        If blnFoundEscapeChar = False Then
            'In the unlikely case that the search string contains all the possible escape characters
            'I just replace all instances of the '~' character and use that as the escape char
            strValue.Replace("~", "")
            intLoop = 1
        End If

        strValue = strValue.Replace("%", arrEscapeChar(intLoop) & "%")
        strValue = strValue.Replace("[", arrEscapeChar(intLoop) & "[")
        strValue = strValue.Replace("]", arrEscapeChar(intLoop) & "]")
        strValue = strValue.Replace("_", arrEscapeChar(intLoop) & "_")
        If Not blnPagingFilter Then
            If strValue.IndexOf("*"c) < 0 Then
                strValue = "%" & strValue & "%"
            End If

            strValue = strValue.Replace("*", "%")
        End If
        Dim strTempValue As String = strValue

        strValue = clsDB.SQLString(strValue)
        strValue = "'" & strValue & "'"

        If strTempValue.IndexOf(arrEscapeChar(intLoop)) >= 0 Then
            strValue &= " ESCAPE '" & arrEscapeChar(intLoop) & "'"
        End If

        Return strValue
    End Function

    Public MustOverride Sub GetDatabaseInfo(ByRef strServer As String, ByRef strDatabase As String,
        ByRef strUserID As String, Optional ByVal dblVersion As Double = 0,
        Optional ByVal strAppName As String = Nothing, Optional ByVal dblMinDBVersion As Double = 0)


    Public MustOverride Sub GetDatabaseGroupInfo(ByRef strGroup As String, ByRef strServer As String, ByRef strDatabase As String,
                                            ByRef strUserID As String, Optional ByVal dblVersion As Double = 0,
                                            Optional ByVal strAppName As String = Nothing, Optional ByVal dblMinDBVersion As Double = 0)

    Public MustOverride Function CheckHostName(ByVal strServer As String) As String

    Public MustOverride Function CheckSQLServerName() As String

    Public MustOverride Sub InitializeServerObjects()

#End Region

#Region " GetItem, GetList Methods "

    ''' <summary>
    ''' Gets the record with the given ID from the specified table
    ''' </summary>
    Public MustOverride Function GetItem(ByVal strTable As String, ByVal intID As Integer) As DataTable

    ''' <summary>
    ''' Gets a list of records from the specified table where strByField matches objValue
    ''' </summary>
    Public MustOverride Function GetList(ByVal strTable As String, ByVal strByField As String,
    ByVal objValue As Object) As DataTable
#End Region

#Region " GetDataTable Methods "

    ''' <summary>
    ''' Returns all records for the specified table
    ''' </summary>
    Public MustOverride Function GetDataTable(ByVal objTable As clsTable) As DataTable

    ''' <summary>
    ''' Returns all records selected in the stored procedure
    ''' </summary>
    Public MustOverride Function GetDataTable(ByVal strStoredProcedure As String) As DataTable

    ''' <summary>
    ''' Returns all records selected in the stored procedure matching the designated SP's parameters
    ''' </summary>
    Public MustOverride Function GetDataTable(ByVal strStoredProcedure As String,
                                              ByVal colParams As clsDBParameterDictionary) As DataTable

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement
    ''' </summary>
    Public MustOverride Function GetDataTableBySQL(ByVal strSQL As String) As DataTable

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement with the designated SP Parameters
    ''' </summary>
    Public MustOverride Function GetDataTableBySQL(ByVal strSQL As String,
                                                   ByVal colParams As clsDBParameterDictionary) As DataTable

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement with the designated parameters.
    ''' Replaces all the format items with the string equivalent value in the corresponding array.
    ''' </summary>
    Public MustOverride Function GetDataTableBySQL(ByVal strSQLFormat As String,
                                                   ByVal colParams As clsDBParameterDictionary,
                                                   ByVal ParamArray args() As Object) As DataTable

    ''' <summary>
    ''' Returns all records selected in the SQL SELECT statement with the designated parameters.
    ''' Replaces all the format items with the string equivalent value in the corresponding array.
    ''' </summary>
    Public MustOverride Function GetDataTableBySQL(ByVal strSQLFormat As String,
                                                   ByVal colParams As clsDBParameterDictionary,
                                                   ByVal blnValidateType As Boolean) As DataTable

    ''' <summary>
    ''' Returns all records for the specified table where strField matches objValue
    ''' </summary>
    Public MustOverride Function GetDataTableByField(ByVal strTable As String,
    ByVal strField As String, ByVal objValue As Object) As DataTable

    ''' <summary>
    ''' Returns all records for the specified table where strField matches objValue
    ''' </summary>
    Public MustOverride Function GetDataTableByField(ByVal strTable As String,
    ByVal colParams As clsDBParameterDictionary) As DataTable

#End Region

#Region " Execute\Delete Methods "

    ''' <summary>
    ''' Executes the specified stored procedure
    ''' </summary>
    Public MustOverride Sub ExecuteProcedure(ByVal strStoredProcedure As String)

    ''' <summary>
    ''' Executes the specified stored procedure with the designated SP Parameters
    ''' </summary>
    Public MustOverride Sub ExecuteProcedure(ByVal strStoredProcedure As String,
    ByRef colParams As clsDBParameterDictionary)

    ''' <summary>
    ''' Executes the specified SQL statement
    ''' </summary>
    Public MustOverride Function ExecuteSQL(ByVal strSQL As String) As Integer

    ''' <summary>
    ''' Executes the specified SQL statement with the designated SP Parameters
    ''' </summary>
    Public MustOverride Function ExecuteSQL(ByVal strSQL As String,
    ByRef colParams As clsDBParameterDictionary) As Integer

    ''' <summary>
    ''' Executes the specified SQL SELECT statement and returns first column (should be integer)
    ''' </summary>
    Public MustOverride Function ExecuteScalar(ByVal strSQL As String) As Integer

    ''' <summary>
    ''' Executes the specified SQL SELECT statement with the designated SP Parameters, and returns first column (should be integer)
    ''' </summary>
    Public MustOverride Function ExecuteScalar(ByVal strSQL As String,
    ByRef colParams As clsDBParameterDictionary) As Integer

    ''' <summary>
    ''' Deletes the record with the designated ID from the specified table
    ''' </summary>
    Public MustOverride Sub DeleteRecord(ByVal strTable As String, ByVal intID As Integer)

    Public Sub DeleteRecord(ByVal objTable As clsTable, ByVal intID As Integer)
        Dim blnCreatedTransaction As Boolean = False
        Dim strSql As String
        Dim objOriginalRecord As clsTableMask = Nothing
        Dim strExternalId As String

        Try
            If HasTransaction = False Then
                BeginTransaction()
                blnCreatedTransaction = True
            End If

            If clsAuditTrail.AuditTableMethodData(Me,
            clsMethod.enumMethods.cDELETE, objTable) Then
                'They are auditing delete data
                objOriginalRecord = New clsTableMask(objTable,
                    clsTableMask.enumMaskType.VIEW, intID)

                'Load all the link table data (as we will be deleting it after this)
                For Each objMaskFieldLink As clsMaskFieldLink In objOriginalRecord.MaskManyToManyCollection.Values
                    objMaskFieldLink.LoadValues(intID)
                Next

                'download the file temporarily so that it can be attached to the audit trail record after deletion
                For Each objMaskField As clsMaskField In objOriginalRecord.MaskFieldCollection.Values

                    If objMaskField.Field.IsBinaryType Then
                        Dim strFile As String = Path.Combine(Path.GetTempPath(), String.Format("Knowledgeone\Client\{0}", Path.GetRandomFileName()))

                        objMaskField.Database.ReadBLOB(objMaskField.Field.Table.DatabaseName,
                            objMaskField.Field.DatabaseName, objMaskField.MaskFieldCollection.ID, strFile)

                        objMaskField.Value1.FileName = strFile

                    End If

                Next

                strExternalId = objOriginalRecord.MaskFieldCollection.ExternalID
            Else
                strExternalId = GetRecordExternalID(objTable, intID)
            End If

            'Remove all the foreign keys to this records
            For Each objFieldLink As clsFieldLink In objTable.OneToManyLinks.Values
                Try
                    strSql = "UPDATE [" & objFieldLink.ForeignKeyTable.DatabaseName & "] " &
                        "SET [" & objFieldLink.ForeignKeyField.DatabaseName & "] = NULL " &
                        "WHERE [" & objFieldLink.ForeignKeyField.DatabaseName & "] = " & intID
                    ExecuteSQL(strSql)
                Catch ex As clsK1Exception
                    If Not (ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_FIELD OrElse
                    ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE) Then
                        Throw New Exception("A problem occurred removing the related link field '" &
                            objFieldLink.ForeignKeyTable.DatabaseName &
                            "." & objFieldLink.ForeignKeyField.DatabaseName & "': " & ex.Message)
                    End If
                End Try
            Next

            'Remove all the link table records
            For Each objFieldLink As clsFieldLink In objTable.ManyToManyLinks.Values
                Try
                    strSql = "DELETE FROM [" & objFieldLink.ForeignKeyTable.DatabaseName & "] " &
                        "WHERE [" & objFieldLink.ForeignKeyField.DatabaseName & "] = " & intID
                    ExecuteSQL(strSql)
                Catch ex As clsK1Exception
                    If Not (ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_FIELD OrElse
                    ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE) Then
                        Throw New Exception("A problem occurred removing the link table '" &
                            objFieldLink.ForeignKeyTable.CaptionText & "' records: " & ex.Message)
                    End If
                End Try
            Next

            If objTable.DatabaseName = clsDBConstants.Tables.cEDOC Then
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

            'Remove the actual record
            Try
                strSql = "DELETE FROM [" & objTable.DatabaseName & "] " &
                    "WHERE [" & clsDBConstants.Fields.cID & "] = " & intID
                ExecuteSQL(strSql)
            Catch ex As clsK1Exception
                If Not ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE Then
                    Throw
                End If
            End Try

            'do audit trail stuff
            Dim objATD As clsAuditTrailRecordData = Nothing

            If objOriginalRecord IsNot Nothing Then
                objATD = New clsAuditTrailRecordData(clsMethod.enumMethods.cDELETE,
                    objOriginalRecord, Nothing)
            End If

            clsAuditTrail.CreateTableMethodRecord(Me, clsMethod.enumMethods.cDELETE,
                objTable, intID, strExternalId, objATD)

            If blnCreatedTransaction Then
                EndTransaction(True)
            End If
        Catch ex As Exception
            If blnCreatedTransaction Then
                EndTransaction(False)
            End If

            Throw
        End Try
    End Sub
#End Region

#Region " Link Table Methods "

    ''' <summary>
    ''' Creates a Link Table record
    ''' </summary>
    Public MustOverride Sub InsertLink(ByVal strLinkTable As String, ByVal strLinkField1 As String,
    ByVal intLinkID1 As Integer, ByVal strLinkField2 As String, ByVal intLinkID2 As Integer)

    ''' <summary>
    ''' Deletes a Link Table record
    ''' </summary>
    Public MustOverride Sub DeleteLink(ByVal strLinkTable As String, ByVal strLinkField1 As String,
    ByVal intLinkID1 As Integer, ByVal strLinkField2 As String, ByVal intLinkID2 As Integer)
#End Region

#Region " Blob Methods "

    ''' <summary>
    ''' Retrieves the size in bytes of the blob field
    ''' </summary>
    Public MustOverride Function GetBLOBSize(ByVal strTableName As String,
    ByVal strFieldName As String, ByVal intID As Integer) As Integer

    ''' <summary>
    ''' Saves the blob to the specified file
    ''' </summary>
    Public MustOverride Sub ReadBLOB(ByVal strTableName As String,
    ByVal strFieldName As String, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)

    ''' <summary>
    ''' Saves the blob to memory
    ''' </summary>
    Public MustOverride Function ReadBLOBToMemory(ByVal strTableName As String,
    ByVal strFieldName As String, ByVal intID As Integer) As Byte()

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public MustOverride Overloads Sub WriteBLOB(ByVal objTable As clsTable,
    ByVal objField As clsField, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)

    ''' <summary>
    ''' Inserts the file into the blob field of the designated record
    ''' </summary>
    Public MustOverride Overloads Sub WriteBLOB(ByVal strTable As String,
    ByVal strField As String, ByVal eDataType As SqlDbType,
    ByVal intDataLength As Integer, ByVal intID As Integer, ByVal strFileName As String, Optional ByVal blnShowTransferUI As Boolean = True)
#End Region

#Region " Training "

    ''' <summary>
    ''' Determines if the record limit has been exceeded for the specified table
    ''' </summary>
    Public MustOverride Function RecordCountExceeded(ByVal strTable As String) As Boolean
#End Region

#Region " Event Methods "

    Public Sub SessionUpdatedMethod()
        RaiseEvent SessionUpdated()
    End Sub

    Public Sub RaiseFileTransferInit(ByVal eType As enumDataAccessType,
    ByVal intTotalChunks As Integer, ByVal eTransferType As enumTransferType)
        RaiseEvent FileTransferInit(eType, intTotalChunks, eTransferType)
    End Sub

    Public Sub RaiseFileTransferStep()
        RaiseEvent FileTransferStep()
    End Sub

    Public Sub RaiseFileTransferEnd()
        RaiseEvent FileTransferEnd()
    End Sub
#End Region

#Region " Bulk Insert "

    ''' <summary>
    ''' Bulk Inserts the records of the datatable to the destination
    ''' </summary>
    ''' <param name="objDT">A data table containing the records to bulk insert</param>
    ''' <param name="strDestTable">The table to bulk insert the records into</param>
    Public MustOverride Sub BulkInsert(ByVal objDT As DataTable, ByVal strDestTable As String)
#End Region

#Region " SQL Methods "

    Public Function GetRecordExternalID(ByVal objTable As clsTable,
    ByVal intID As Integer) As String
        Dim strExternalID As String = Nothing

        Try
            Dim strSQL As String = "SELECT [" & clsDBConstants.Fields.cEXTERNALID & "] " &
                "FROM [" & objTable.DatabaseName & "] " &
                "WHERE [" & clsDBConstants.Fields.cID & "] = @ID"

            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@ID", intID))

            Dim objDT As DataTable = GetDataTableBySQL(strSQL, colParams)

            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                strExternalID = CStr(objDT.Rows(0)(0))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If
        Catch ex As Exception
            strExternalID = Nothing
        End Try

        Return strExternalID
    End Function

    Public Overridable Function GetCurrentTime() As Date
        Try
            Dim dtNow As Date

            Dim objDT As DataTable = GetDataTableBySQL("SELECT GETDATE()")
            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                dtNow = CDate(objDT.Rows(0)(0))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return dtNow
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.ToString)
            Throw
        End Try
    End Function

    Public Function GetServerTimeDifference() As Double 'localtimedate
        Try
            Dim dblNow As Double

            Dim objDT As DataTable = GetDataTableBySQL("Select DATEDIFF(minute, GETUTCDATE(), GETDATE())")
            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                dblNow = CDbl(objDT.Rows(0)(0))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Dim currentOffset As TimeSpan = TimeZone.CurrentTimeZone.GetUtcOffset(Now)

            dblNow = (dblNow / 60) - currentOffset.TotalHours

            Return dblNow
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.ToString)
            Throw
        End Try
    End Function

    Public Function GetServerTimeZone() As TimeZoneInfo
        ' 2000003598  - Updated to attempt to get the timezone and on fail return web server timezone  
        Dim objTimeZoneResponse = GetCurrentTimeZone()
        If ValidateTimeZoneRespsonse(objTimeZoneResponse) Then
            Return objTimeZoneResponse.TimeZoneInfo
        Else
            objTimeZoneResponse = GetCurrentTimeZoneRegRead()
            If ValidateTimeZoneRespsonse(objTimeZoneResponse) Then
                Return objTimeZoneResponse.TimeZoneInfo
            Else
                Return TimeZoneInfo.Local
            End If
        End If
    End Function

    Private Function ValidateTimeZoneRespsonse(objTimeZoneResponse As (Success As Boolean, TimeZoneInfo As TimeZoneInfo)) As Boolean
        Return objTimeZoneResponse.Success AndAlso objTimeZoneResponse.TimeZoneInfo IsNot Nothing
    End Function

    Private Function GetCurrentTimeZoneRegRead() As (Success As Boolean, TimeZoneInfo As TimeZoneInfo)
        Try
            Dim strSQL As String = $"BEGIN
                                        SET NOCOUNT ON;
	                                    DECLARE @TimeZone VARCHAR(50)
	                                    EXEC MASTER.dbo.xp_regread 'HKEY_LOCAL_MACHINE',
	                                    'SYSTEM\CurrentControlSet\Control\TimeZoneInformation',
	                                    'TimeZoneKeyName',@TimeZone OUT
	                                    SELECT @TimeZone
                                    END"

            Dim objDT As DataTable = GetDataTableBySQL(strSQL)
            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                Dim strName As String = CStr(objDT.Rows(0)(0))
                Dim objInfo = TimeZoneInfo.FindSystemTimeZoneById(strName)
                Return (True, objInfo)
            Else
                Return (False, Nothing)
            End If
        Catch ex As Exception
            LogException(ex)
            Return (False, Nothing)
        End Try
    End Function

    Private Function GetCurrentTimeZone() As (Success As Boolean, TimeZoneInfo As TimeZoneInfo)
        Try
            Dim strSQL As String = $"BEGIN 
                                        SELECT CURRENT_TIMEZONE()
                                     END"

            ' This stored procedure returns the display name of the timezone
            Dim objDT As DataTable = GetDataTableBySQL(strSQL)

            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                Dim strDisplayName As String = CStr(objDT.Rows(0)(0))
                Dim objInfo = TimeZoneInfo.GetSystemTimeZones() _
                            .SingleOrDefault(
                                Function(TimeZoneInfo)
                                    Return TimeZoneInfo.DisplayName = strDisplayName
                                End Function)
                Return (True, objInfo)
            Else
                Return (False, Nothing)
            End If
        Catch ex As Exception
            LogException(ex)
            Return (False, Nothing)
        End Try
    End Function

    Private Sub LogException(objException As Exception)
#If DEBUG Then
        Using objLogger = clsK1Logger.Instance
            objLogger.strApplicationName = NameOf(clsDB)
            If TypeOf objException Is SqlException Then
                Dim objSqlException = TryCast(objException, SqlException)
                objLogger?.Log($"{objSqlException?.ErrorCode} - {objSqlException?.Message}")
            Else
                objLogger?.Log($"{objException?.Message}")
            End If
        End Using
#End If
    End Sub

    Public MustOverride Sub CreateStoredProcedure(sbProcName As String, sbProcedure As String)

    Public MustOverride Function CheckStoredProcedureExists(sbProcName As String) As Boolean

    '2017-08-24 -- Peter Melisi -- Changes for Timezones for User Profiles
    Public Function GetServerTimeDifference(ByVal strTimezone As String) As Double 'localtimedate
        Try
            Dim dblNow As Double

            Dim objDT As DataTable = GetDataTableBySQL("Select DATEDIFF(minute, GETUTCDATE(), GETDATE())")
            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                dblNow = CDbl(objDT.Rows(0)(0))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Dim currentOffset As TimeSpan

            If Not strTimezone.Equals(String.Empty) Then
                Dim systemTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(strTimezone)

                currentOffset = systemTimeZone.BaseUtcOffset

                '2017-10-12 -- Peter Melisi -- Bug fix for incorrect calculation during Daylight Saving Time.
                If systemTimeZone.IsDaylightSavingTime(DateTime.Now) Then
                    currentOffset = currentOffset.Add(New TimeSpan(1, 0, 0))
                End If
            Else
                currentOffset = TimeZone.CurrentTimeZone.GetUtcOffset(Now)
            End If

            dblNow = (dblNow / 60) - currentOffset.TotalHours

            Return dblNow
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.ToString)
            Throw
        End Try
    End Function
#End Region

#Region " SQL Dependency Methods "

    Public MustOverride Sub AddSqlNotification(ByVal strTable As String,
                                               ByVal strField As String,
                                               ByVal objCallBack As SqlClient.OnChangeEventHandler)

    Public MustOverride Sub AddSqlNotification(ByVal strSQL As String,
                                               ByVal objCallBack As SqlClient.OnChangeEventHandler)

    Protected Sub RegisterSqlNotificationCallBack(ByVal strID As String,
                                                  ByVal objCallBack As SqlClient.OnChangeEventHandler)
#If TRACE Then
        Trace.WriteVerbose($"(RegisterSqlNotificationCallBack) Register Id - {strID} with {objCallBack.Method.Name}", "")
#End If

        m_colNotifications.Add(strID, objCallBack)

    End Sub

    Private Function RemoveSqlNotification(ByVal objSqlDependency As SqlClient.SqlDependency) As System.Data.SqlClient.OnChangeEventHandler
        Try
            If m_colNotifications.ContainsKey(objSqlDependency.Id) Then
                Dim objCallBack As System.Data.SqlClient.OnChangeEventHandler

                objCallBack = m_colNotifications(objSqlDependency.Id)

                '-- notifications are only a one shot deal so remove the existing event handler
                '-- so a new one can be added next time
                RemoveHandler objSqlDependency.OnChange, objCallBack

                Return objCallBack
            End If

            Return Nothing
        Catch ex As Exception
            Throw
        End Try
    End Function

    Protected Sub OnSqlNotification(ByVal sender As Object, ByVal e As SqlClient.SqlNotificationEventArgs)
        Try
            Dim objSqlDependency As SqlClient.SqlDependency = CType(sender, SqlClient.SqlDependency)

#If TRACE Then
            Trace.WriteVerbose("(OnSqlNotification) Id - " & objSqlDependency.Id, "")
            Trace.WriteVerbose("(OnSqlNotification) Info - " & e.Info.ToString, "")
            Trace.WriteVerbose("(OnSqlNotification) Source - " & e.Source.ToString, "")
            Trace.WriteVerbose("(OnSqlNotification) Type - " & e.Type.ToString, "")
#End If

            If m_colNotifications.ContainsKey(objSqlDependency.Id) Then
                Dim objCallBack As System.Data.SqlClient.OnChangeEventHandler

                objCallBack = m_colNotifications(objSqlDependency.Id)

                If objCallBack IsNot Nothing Then
                    '-- Notifications are only a one shot deal so remove the existing event handler
                    '-- so a new one can be added next time
                    RemoveHandler objSqlDependency.OnChange,
                        New SqlClient.OnChangeEventHandler(AddressOf OnSqlNotification)

                    objCallBack.Invoke(sender, e)
                    'sender.Invoke(sender, e)
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " System Information Refresh "

    Public Overridable Sub RefreshSysInfo()
        Dim objSysInfoReload As New clsSysInfo(Me)
        m_objSysInfoReload = objSysInfoReload
    End Sub
#End Region

#Region " Session Functionality "

    Public Overridable Sub SessionStart(ByVal eAppType As clsDBConstants.enumApplicationType)
        m_objSession = New clsSession(Me, eAppType)
    End Sub

    Public Overridable Sub SessionStop()
        If Not m_objSession Is Nothing Then
            m_objSession.Delete()
            m_objSession.Dispose()
            m_objSession = Nothing
        End If
    End Sub

    Public Overridable Sub SessionUpdate()
        If Not m_objSession Is Nothing Then
            If m_blnMultiThreaded Then
                Dim objSessionThread As New Threading.Thread(AddressOf SessionUpdateThread)
                objSessionThread.IsBackground = True
                objSessionThread.Name = "Session Updater"
                objSessionThread.Start()
            Else
                If Not m_objSession Is Nothing AndAlso Not Me.Session.Expired Then
                    m_objSession.Update()
                End If
            End If
        End If
    End Sub

    Protected Overridable Sub SessionUpdateThread()
        Try
            If Not m_objSession Is Nothing AndAlso Not Me.Session.Expired Then
                m_objSession.Update()
            ElseIf Me.Session.Expired Then
                modEvents.RaiseSessionExpired()
            End If
        Catch ex As Exception
            '-- ignore
        End Try

    End Sub
#End Region

#Region " StateChanged "

    Protected Overridable Sub DBStateChanged()
        If m_eState = enumState.DRM_LOCKED_SESSION Then
            m_blnCancel = True
        End If
    End Sub

#End Region

    Public Sub CheckVersion(ByVal dblVersion As Double, ByVal strAppName As String, ByVal dblMinDBVersion As Double)

        If strAppName = "ConnectionManagerUtility" _
        OrElse strAppName = "Database Updater" _
        OrElse strAppName = "Installer" _
        OrElse strAppName = "API Library" _
        OrElse strAppName = "VERS Database Updater" _
        OrElse strAppName = "VEO Export Tool" Then
            Return
        End If

        Dim intDBVer As Double = Me.GetDatabaseVersion()

        If dblVersion >= 2.04 Then
            If (intDBVer >= dblMinDBVersion) Then
                Dim colParams As New clsDBParameterDictionary
                colParams.Add(New clsDBParameter("@AppName", strAppName, ParameterDirection.Input, SqlDbType.NVarChar))
                colParams.Add(New clsDBParameter("@AppVersion", dblVersion, ParameterDirection.Input, SqlDbType.Decimal))
                Me.ExecuteProcedure(clsDBConstants.StoredProcedures.cUI_CHECK_VERSION, colParams)
            Else
                Throw New clsK1Exception(ErrorNumber.Old_Version, "Could not connect to the database. The database is not compatible with this version.")
            End If
        ElseIf dblVersion >= 2.03 AndAlso dblVersion < 2.04 Then
            If (intDBVer >= 11.05 AndAlso intDBVer <= 11.06) Then
                Return
            ElseIf intDBVer < 11.05 Then
                Throw New clsK1Exception(ErrorNumber.Old_Version, "Could not connect to the database. The database version is out of date. If you wish to connect to an old database please install " & strAppName & " version 2.2.")
            Else
                Throw New clsK1Exception(ErrorNumber.Old_Version, "Your " & strAppName & " version is out of date. Please install the latest version before proceeding.")
            End If
        Else
            If (intDBVer <= 11.04) Then
                Return
            Else
                Throw New clsK1Exception(ErrorNumber.Old_Version, "Your " & strAppName & " version is out of date. Please install the latest version before proceeding.")
            End If
        End If
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then

                If (objSqlNotifications IsNot Nothing) Then
                    For Each notifier As clsSqlNotifier In objSqlNotifications
                        notifier.Dispose()
                    Next
                    objSqlNotifications.Clear()
                    objSqlNotifications = Nothing
                End If

                If m_objSession IsNot Nothing Then
                    If m_objSession.Type = clsSession.enumSessionType.MAIN Then
                        Try
                            SessionStop()
                        Catch ex As Exception
                        End Try
                    End If
                    m_objSession = Nothing
                End If

                If m_objSQLDependency IsNot Nothing Then
                    m_objSQLDependency.Dispose()
                    m_objSQLDependency = Nothing
                End If

                m_objSysInfo = Nothing
                m_objProfile = Nothing
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
