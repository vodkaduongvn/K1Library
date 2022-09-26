#Region " File Information "

'=====================================================================
' Contains information regarding a database table index
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date        Description
'---------------------------------------------------------------------
' KSD       16/01/2007  Implemented.
'=====================================================================

#End Region

#End Region

Public Class clsTableIndex
    Implements IDisposable

#Region " Members "

    Private m_strTable As String
    Private m_strName As String
    Private m_blnClustered As Boolean
    Private m_blnUnique As Boolean
    Private m_blnUniqueConstraint As Boolean
    Private m_blnPrimaryKey As Boolean
    Private m_colFields As New FrameworkCollections.K1Collection(Of String)
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls

#End Region

#Region " Constructors "

    Public Sub New(ByVal strTable As String, ByVal colFields As FrameworkCollections.K1Collection(Of String), _
                   ByVal blnClustered As Boolean, ByVal blnUnique As Boolean, ByVal blnPrimaryKey As Boolean, _
                   ByVal blnUniqueConstraint As Boolean)
        MyBase.New()
        m_strTable = strTable
        m_strName = GenerateIndexName(strTable, colFields)
        m_colFields = colFields
        m_blnClustered = blnClustered
        m_blnUnique = blnUnique
        m_blnPrimaryKey = blnPrimaryKey
        m_blnUniqueConstraint = blnUniqueConstraint
    End Sub

    Public Sub New(ByVal strTable As String, ByVal objDataRow As DataRow)
        MyBase.New()
        m_strTable = strTable
        m_strName = CStr(objDataRow("index_name"))
        m_colFields.AddRange(CStr(objDataRow("index_keys")).Replace(", ", " ").Split(Chr(32)))
        FormatDescription(CStr(objDataRow("index_description")), m_blnUnique, m_blnClustered, m_blnPrimaryKey, m_blnUniqueConstraint)
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property Table() As String
        Get
            Return m_strTable
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return m_strName
        End Get
    End Property

    Public ReadOnly Property PrimaryKey() As Boolean
        Get
            Return m_blnPrimaryKey
        End Get
    End Property

    Public ReadOnly Property Clustered() As Boolean
        Get
            Return m_blnClustered
        End Get
    End Property

    Public ReadOnly Property Unique() As Boolean
        Get
            Return m_blnUnique
        End Get
    End Property

    Public ReadOnly Property UniqueConstraint() As Boolean
        Get
            Return m_blnUniqueConstraint
        End Get
    End Property

    Public Property Fields() As FrameworkCollections.K1Collection(Of String)
        Get
            Return m_colFields
        End Get
        Set(ByVal Value As FrameworkCollections.K1Collection(Of String))
            m_colFields = Value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Get Table Indexes "

    ''' <summary>
    ''' Retrieves a list of indexes pertaining to a particular table
    ''' </summary>
    Public Shared Function GetTableIndexes(ByVal objDB As clsDB_System, _
    ByVal strTable As String) As List(Of clsTableIndex)
        Try
            Return GetTableIndexes(objDB, strTable, Nothing)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Retrieves a list of indexes pertaining to a particular table and field
    ''' </summary>
    Public Shared Function GetTableIndexes(ByVal objDB As clsDB_System, _
    ByVal strTable As String, ByVal strField As String) As List(Of clsTableIndex)
        Try
            Dim objDT As DataTable
            Dim colParams As New clsDBParameterDictionary
            Dim colTableIndexes As List(Of clsTableIndex) = New List(Of clsTableIndex)

            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName("objname"), strTable))
            objDT = objDB.GetDataTable("sp_helpindex", colParams)

            For Each objDR As DataRow In objDT.Rows
                Dim objTableIndex As New clsTableIndex(strTable, objDR)

                If String.IsNullOrEmpty(strField) OrElse objTableIndex.Fields.Contains(strField) Then
                    colTableIndexes.Add(objTableIndex)
                End If
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colTableIndexes
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Format Description "

    ''' <summary>
    ''' Breaks down the "index_description" returned from sp_helpindex to find the indexes
    ''' attributes
    ''' </summary>
    Private Sub FormatDescription(ByVal strDescription As String, ByRef blnUnique As Boolean, _
                                  ByRef blnClustered As Boolean, ByRef blnPrimaryKey As Boolean, _
                                  ByRef blnUniqueConstraint As Boolean)
        Dim intEndIndex As Integer = strDescription.IndexOf("located") - 1
        Dim arrDescription() As String = strDescription.Substring(0, intEndIndex).Split(","c)

        For Each strDescription In arrDescription
            Select Case strDescription.Trim
                Case "unique"
                    blnUnique = True

                Case "clustered"
                    blnClustered = True

                Case "primary key"
                    blnPrimaryKey = True

                Case "unique key"
                    blnUniqueConstraint = True
            End Select
        Next
    End Sub

#End Region

#Region " Generate Index Name "

    Private Shared Function GenerateIndexName(ByVal strTablename As String, _
        ByVal colFields As FrameworkCollections.K1Collection(Of String)) As String
        Try
            Dim strName As String = ""

            For Each strField As String In colFields
                strName &= "_" & strField.Trim
            Next

            If strName.Length > 0 Then
                strName = "IX_" & strTablename & strName
            End If

            Return strName
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                If Not m_colFields Is Nothing Then
                    m_colFields.Clear()
                    m_colFields = Nothing
                End If
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region " Create Index "

    ''' <summary>
    ''' Creates the index
    ''' </summary>
    Public Sub Create(ByVal objDB As clsDB_System)
        Create(objDB, Me)
    End Sub

    ''' <summary>
    ''' Creates an index using the tableindex object provided
    ''' </summary>
    Private Shared Sub CreateIndex(ByVal objDB As clsDB_System, ByVal objTableIndex As clsTableIndex)
        Dim strSQL As String = "CREATE "

        If objTableIndex.Unique Then
            strSQL &= "UNIQUE "
        End If

        If objTableIndex.Clustered Then
            strSQL &= "CLUSTERED"
        Else
            strSQL &= "NONCLUSTERED"
        End If

        strSQL &= " INDEX [" & clsDB_System.SQLString(objTableIndex.Name) & "] ON [" & _
            clsDB_System.SQLString(objTableIndex.Table) & "]("

        Dim strFields As String = ""
        For Each strField As String In objTableIndex.Fields
            AppendToCommaString(strFields, "[" & clsDB_System.SQLString(strField) & "]")
        Next

        strSQL &= clsDB_System.SQLString(strFields) & ")"

        objDB.ExecuteSQL(strSQL)
    End Sub

    Public Shared Sub Create(ByVal objDB As clsDB_System, ByVal objTableIndex As clsTableIndex)
        If objTableIndex.PrimaryKey Then
            CreatePrimaryKey(objDB, objTableIndex)
        ElseIf objTableIndex.UniqueConstraint Then
            CreateUniqueConstraint(objDB, objTableIndex)
        Else
            CreateIndex(objDB, objTableIndex)
        End If
    End Sub

    Private Shared Sub CreatePrimaryKey(ByVal objDB As clsDB_System, ByVal objTableIndex As clsTableIndex)
        Try
            Dim strSQL As String = "ALTER TABLE {0} WITH NOCHECK" & vbCrLf & _
                        "ADD CONSTRAINT PK_{0}_{1} PRIMARY KEY CLUSTERED ({1})"

            strSQL = String.Format(strSQL, objTableIndex.Table, objTableIndex.Fields(0))

            objDB.ExecuteSQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Shared Sub CreateUniqueConstraint(ByVal objDB As clsDB_System, ByVal objTableIndex As clsTableIndex)
        Try
            Dim strCluster As String = ""
            Dim strSQL As String = "ALTER TABLE {0} WITH NOCHECK" & vbCrLf & _
                                  "ADD CONSTRAINT IX_{0}_{1} UNIQUE {2} ({3})"

            Dim arrField() As String = objTableIndex.Fields.ToArray
            Dim strFields As String = String.Join(",", arrField)

            If objTableIndex.Clustered Then
                strCluster = "CLUSTERED"
            Else
                strCluster = "NONCLUSTERED"
            End If

            strSQL = String.Format(strSQL, objTableIndex.Table, strFields, strCluster, _
                                   String.Format("[{0}]", strFields.Replace(",", "],[")))

            objDB.ExecuteSQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Drop Index "

    ''' <summary>
    ''' Removes an index from the database
    ''' </summary>
    Public Sub DropIndex(ByVal objDB As clsDB_System)
        If Me.PrimaryKey OrElse Me.UniqueConstraint Then
            DropConstraintIndex(objDB, m_strTable, m_strName)
        Else
            DropIndex(objDB, m_strTable, m_strName)
        End If
    End Sub

    ''' <summary>
    ''' Removes an index from the database
    ''' </summary>
    Public Shared Sub DropIndex(ByVal objDB As clsDB_System, ByVal strTable As String, ByVal strName As String)
        objDB.ExecuteSQL("DROP INDEX [" & strTable & "].[" & strName & "]")
    End Sub

    ''' <summary>
    ''' Removes a unique index from the table
    ''' </summary>
    Public Sub DropConstraintIndex(ByVal objDB As clsDB_System)
        DropConstraintIndex(objDB, m_strTable, m_strName)
    End Sub

    ''' <summary>
    ''' Removes the Primary Key from the table
    ''' </summary>
    Public Shared Sub DropConstraintIndex(ByVal objDB As clsDB_System, ByVal strTable As String, ByVal strName As String)
        objDB.ExecuteSQL("ALTER TABLE [" & strTable & "] DROP CONSTRAINT " & strName)
    End Sub

#End Region

#Region " Drop Indexes "

    ''' <summary>
    ''' Removes all indexes pertaining to a particular table
    ''' </summary>
    Public Shared Sub DropIndexes(ByVal objDB As clsDB_System, ByVal objField As clsField)
        DropIndexes(objDB, objField.Table.DatabaseName, objField.DatabaseName)
    End Sub

    ''' <summary>
    ''' Removes all indexes pertaining to a particular table (and field)
    ''' </summary>
    Public Shared Sub DropIndexes(ByVal objDB As clsDB_System, ByVal strTable As String, ByVal strField As String)
        Dim colIndexes As List(Of clsTableIndex) = clsTableIndex.GetTableIndexes(objDB, strTable)
        For Each objIndex As clsTableIndex In colIndexes
            If strField Is Nothing OrElse objIndex.Fields.Contains(strField) Then
                If objIndex.PrimaryKey OrElse objIndex.UniqueConstraint Then
                    objIndex.DropConstraintIndex(objDB)
                Else
                    objIndex.DropIndex(objDB)
                End If
            End If
        Next
    End Sub

#End Region

#Region " Rename Index "

    ''' <summary>
    ''' Renames an index in the database
    ''' </summary>
    ''' <param name="strTableName">Name of the table the index is on</param>
    ''' <param name="strFieldName">Name of the field the index is bound to</param>
    ''' <param name="strNewTableName">New Name of the table the index is on (If Null or Empty then same as strTableName)</param>
    ''' <param name="strNewFieldName">New Name of the field the index is bound to (If Null or Empty then same as strFieldName)</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameIndex(ByVal objDB As clsDB_System, ByVal strTableName As String, _
    ByVal strFieldName As String, ByVal strNewTableName As String, ByVal strNewFieldName As String)
        Try
            If String.IsNullOrEmpty(strNewTableName) Then
                strNewTableName = strTableName
            End If

            If String.IsNullOrEmpty(strNewFieldName) Then
                strNewFieldName = strFieldName
            End If

            RenameIndex(objDB, strTableName, "IX_" & strTableName & "_" & strFieldName, _
                "IX_" & strNewTableName & "_" & strNewFieldName)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames an index in the database
    ''' </summary>
    ''' <param name="strTableName">Name of the table the index is on</param>
    ''' <param name="objIndex">The existing index</param>
    ''' <param name="strNewTableName">New Name of the table the index is on (If Null or Empty then same as strTableName)</param>
    ''' <param name="strNewFieldName">New Name of the field the index is bound to (If Null or Empty then same as strFieldName)</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameIndex(ByVal objDB As clsDB_System, ByVal strTableName As String, _
    ByVal objIndex As clsTableIndex, ByVal strNewTableName As String, ByVal strNewFieldName As String)
        Try
            If String.IsNullOrEmpty(strNewTableName) Then
                Return
            End If

            RenameIndex(objDB, strTableName, objIndex.Name, _
                "IX_" & strNewTableName & "_" & strNewFieldName)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Renames an index in the database
    ''' </summary>
    ''' <param name="strTableName">Name of the table the index is on</param>
    ''' <param name="strOldName">Current name of the index (In the format IX_[Table Name]_[Field Name]</param>
    ''' <param name="strNewName">New name for the index (In the format IX_[Table Name]_[Field Name]</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameIndex(ByVal objDB As clsDB_System, ByVal strTableName As String, _
    ByVal strOldName As String, ByVal strNewName As String)
        Try
            Dim strSQL As String

            strSQL = "EXEC sp_rename N'[dbo].[" & clsDB_System.SQLString(strTableName) & "]." & _
                clsDB_System.SQLString(strOldName) & "', N'" & clsDB_System.SQLString(strNewName) & "', N'INDEX'"

            objDB.ExecuteSQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Rename Table Indexes "

    ''' <summary>
    ''' Renames all the indexes for a specified table
    ''' i.e. replace strTableName with strNewName in the index name
    ''' </summary>
    ''' <param name="strTableName">Name of the table the index is on</param>
    ''' <param name="strNewName">New name for the table the index is on</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameTableIndexes(ByVal objDB As clsDB_System, ByVal strTableName As String, _
    ByVal strNewName As String)
        Try
            Dim colIndexes As List(Of clsTableIndex) = GetTableIndexes(objDB, strTableName)

            If colIndexes Is Nothing Then
                Return
            End If

            For Each objIndex As clsTableIndex In colIndexes
                Dim strFields As String = ""
                For Each strField In objIndex.Fields
                    strFields &= strField
                Next
                RenameIndex(objDB, strTableName, objIndex, strNewName, strFields)
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Rename Field Indexes "

    ''' <summary>
    ''' Renames all the indexes for a specified table and field
    ''' i.e. replace strFieldName with strNewFieldName in the index name
    ''' </summary>
    ''' <param name="strTableName">Name of the table the index is on</param>
    ''' <param name="strFieldName">Name of the field the index is bound to</param>
    ''' <param name="strNewFieldName">New name for the field the index is bound to</param>
    ''' <remarks></remarks>
    Public Shared Sub RenameFieldIndexes(ByVal objDB As clsDB_System, ByVal strTableName As String, _
    ByVal strFieldName As String, ByVal strNewFieldName As String)
        Try
            Dim colIndexes As List(Of clsTableIndex) = GetTableIndexes(objDB, strTableName)

            If colIndexes Is Nothing OrElse colIndexes.Count = 0 Then
                Return
            End If

            Dim blnUpdate As Boolean = False
            For Each objIndex As clsTableIndex In colIndexes
                Dim intIndex As Integer = objIndex.Fields.IndexOf(strFieldName)
                If intIndex >= 0 Then
                    objIndex.Fields(intIndex) = strNewFieldName
                    RenameIndex(objDB, strTableName, objIndex.Name, GenerateIndexName(strTableName, objIndex.Fields))
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region


    ''' <summary>
    ''' Indexes a foreign key field for faster access
    ''' </summary>
    Public Shared Sub CreateForeignKeyIndex(ByVal objDB As clsDB_System, ByVal objField As clsField)
        objDB.ExecuteSQL("CREATE NONCLUSTERED INDEX [IX_" & objField.Table.DatabaseName & "_" & _
            objField.DatabaseName & "] ON " & _
            "[" & objField.Table.DatabaseName & "]([" & objField.DatabaseName & "])")
    End Sub

    ''' <summary>
    ''' Indexes a foreign key field for faster access
    ''' </summary>
    Public Shared Sub CreateForeignKeyIndex(ByVal objDB As clsDB_System, ByVal strTable As String, ByVal strField As String)
        objDB.ExecuteSQL("CREATE NONCLUSTERED INDEX [IX_" & strTable & "_" & _
            strField & "] ON [" & strTable & "]([" & strField & "])")
    End Sub

    Public Shared Function IndexExist(ByVal objDB As clsDB_System, ByVal strIndexName As String) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@Name", strIndexName))

            Dim strSQL As String = "SELECT 1 FROM sys.indexes WHERE name=@Name"
            Dim blnResult As Boolean = CBool(objDB.GetColumnBySQL(strSQL, colParams))
            colParams.Dispose()

            Return blnResult
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function Exists(ByVal objDB As clsDB_System) As Boolean
        Return IndexExist(objDB, Me.Name)
    End Function
    

End Class
