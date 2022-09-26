#Region " File Information "

'==============================================================================
' This class is used to build a sql select statement and return a dataset
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       21/05/2005  Implemented
' KSD       16/01/2007  Added BuilD SQL Methods
'==============================================================================

#End Region

#End Region

Public Class clsSelectInfo
    Implements IDisposable

#Region " Enumerations "

    Public Enum enumSelectType
        COUNT = 0
        PAGE = 1
        COUNT_TO_REC = 2
        EXPORT = 3
        SELECTION = 4
        FIND = 5
        SELECT_FIELDS = 6
        GOTO_PAGE = 7
    End Enum

    Public Enum enumDatePartType
        Year = 1
        Month = 2
        Day = 3
        Hour = 4
        Minute = 5
        Second = 6
        Millesecond = 7
    End Enum

#End Region

#Region " Members "

    'Initialization Information
    Private m_eSelectType As enumSelectType
    Private m_objDB As clsDB
    Private m_objTable As clsTable
    Private m_objSearchFilter As clsSearchFilter
    Private m_intNumRecords As Integer
    Private m_colSorts As clsSortCollection
    Private m_blnForward As Boolean = True
    Private m_colRecord As clsMaskFieldDictionary
    Private m_colRecordEnd As clsMaskFieldDictionary
    Private m_blnInclRefRecord As Boolean
    Private m_intRecordNumber As Integer
    Private m_blnExcludeBinary As Boolean
    Private m_colOnlyFields As New Hashtable 'Used with the SELECT_FIELDS Enum to specify which fields get selected
    Private m_intTypeID As Integer = clsDBConstants.cintNULL
    Private m_blnUseParameters As Boolean = True

    Private m_blnIgnoreDataSecurity As Boolean

    '[Naing] Bug FIx: 1400002590. So that security set on fields will be ignored when using this to query the database.
    Private m_blnIgnoreFieldSecurity As Boolean

    Private m_blnIgnoreWildCards As Boolean 'added in 2.5 for smart and web client URL string searches
    'Used during the Build SQL Process
    Private m_blnUseFKeyExternalIDs As Boolean = True
    Private m_colFields As New Collection 'Hold field names and their respective SQL declaration
    Private m_colFilters As New List(Of String) 'Will contain all the filters
    Private m_colTableDeclarations As New Collection 'Will contain the tables in join order
    Private m_colTablesUsages As New Hashtable 'Holds incremental values of how many times a tables has been aliased
    Private m_colTableAliases As New Hashtable 'holds all the unique tables in query
    Private m_colTempTables As New Hashtable 'The key will be the temp table name, and value will be string of IDs
    Private m_intTotalIDs As Integer 'determines whether we need to use temp tables
    Private m_strSQL As String
    Private m_strOrderBy As String
    Private m_colParams As clsDBParameterDictionary

    'Returned from running the constructed SQL Query
    Private m_objDT As DataTable

    Private m_blnDisposedValue As Boolean


#End Region

#Region " Constructors "

#Region " enumSelectType.COUNT "

    ''' <summary>
    ''' (enumSelectType.COUNT) Returns a count of the records in the table
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal objSearchFilter As clsSearchFilter)
        m_eSelectType = enumSelectType.COUNT
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_objSearchFilter = objSearchFilter
    End Sub

    ''' <summary>
    ''' (enumSelectType.COUNT) Returns a count of the records in the table
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal objSearchFilter As clsSearchFilter, ByVal blnIgnoreDataSecurity As Boolean)
        Me.New(objTable, objSearchFilter)
        m_blnIgnoreDataSecurity = blnIgnoreDataSecurity
    End Sub


#End Region

#Region " enumSelectType.EXPORT "

    ''' <summary>
    ''' (enumSelectType.EXPORT) Returns all records for the table
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable,
                   ByVal colSorts As clsSortCollection,
                   ByVal objSearchFilter As clsSearchFilter,
                   ByVal blnExcludeBinary As Boolean,
                   ByVal blnUseFKeyExternalIDs As Boolean,
                   ByVal intTypeID As Integer,
                   Optional ByVal arrFields() As String = Nothing,
                   Optional ByVal blnIgnoreFieldSecurity As Boolean = False,
                   Optional ByVal blnIsDropDownRequest As Boolean = False)

        m_eSelectType = enumSelectType.EXPORT
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_colSorts = colSorts
        m_objSearchFilter = objSearchFilter
        m_blnExcludeBinary = blnExcludeBinary
        m_blnUseFKeyExternalIDs = blnUseFKeyExternalIDs
        m_intTypeID = intTypeID

        If arrFields IsNot Nothing Then
            For Each strField As String In arrFields
                m_colOnlyFields.Add(strField, strField)
            Next
        End If

        m_blnIgnoreFieldSecurity = blnIgnoreFieldSecurity
        Me.blnIsDropDownRequest = blnIsDropDownRequest

    End Sub

#End Region

#Region " enumSelectType.PAGE, enumSelectType.COUNT_TO_REC "

    ''' <summary>
    ''' (enumSelectType.PAGE) Retrieves a Page of Data (Only Columns or Sorted Fields) starting from colRecord
    ''' (enumSelectType.COUNT_TO_REC) Counts the number of records up to colRecord
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, _
    ByVal intNumberRecords As Integer, _
    ByVal blnForward As Boolean, _
    ByVal colRecord As clsMaskFieldDictionary, _
    ByVal blnIncludeReferencedRecord As Boolean, _
    ByVal eSelectType As enumSelectType, _
    ByVal intTypeID As Integer, _
    Optional ByVal arrFields() As String = Nothing)
        m_eSelectType = eSelectType
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_colSorts = colSorts
        m_objSearchFilter = objSearchFilter
        m_intNumRecords = intNumberRecords
        m_blnForward = blnForward
        m_colRecord = colRecord
        m_blnInclRefRecord = blnIncludeReferencedRecord
        m_intTypeID = intTypeID
        If arrFields IsNot Nothing Then
            For Each strField As String In arrFields
                m_colOnlyFields.Add(strField, strField)
            Next
        End If
    End Sub


    ''' <summary>
    ''' (enumSelectType.PAGE) Retrieves a Page of Data (Only Columns or Sorted Fields) starting from colRecord
    ''' (enumSelectType.COUNT_TO_REC) Counts the number of records up to colRecord
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, _
    ByVal intNumberRecords As Integer, _
    ByVal blnForward As Boolean, _
    ByVal colRecord As clsMaskFieldDictionary, _
    ByVal blnIncludeReferencedRecord As Boolean, _
    ByVal eSelectType As enumSelectType, _
    ByVal intTypeID As Integer,
    ByVal blnIgnoreDataSecurity As Boolean, _
    Optional ByVal arrFields() As String = Nothing,
    Optional ByVal blnIgnoreFieldSecurity As Boolean = False)
        Me.New(objTable, colSorts, objSearchFilter, intNumberRecords, blnForward, colRecord, _
               blnIncludeReferencedRecord, eSelectType, intTypeID, arrFields)

        m_blnIgnoreDataSecurity = blnIgnoreDataSecurity
        m_blnIgnoreFieldSecurity = blnIgnoreFieldSecurity
    End Sub
#End Region

#Region " enumSelectType.GOTO_PAGE "

    ''' <summary>
    ''' (enumSelectType.GOTO_PAGE) Will retrieve page of data after intRecordNumber
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, _
    ByVal intNumberRecords As Integer, _
    ByVal blnForward As Boolean, _
    ByVal intRecordNumber As Integer, _
    ByVal intTypeID As Integer)
        m_eSelectType = enumSelectType.GOTO_PAGE
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_colSorts = colSorts
        m_objSearchFilter = objSearchFilter
        m_intNumRecords = intNumberRecords
        m_blnForward = blnForward
        m_intRecordNumber = intRecordNumber
        m_intTypeID = intTypeID
    End Sub

    ''' <summary>
    ''' (enumSelectType.GOTO_PAGE) Will retrieve page of data after intRecordNumber
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, ByVal colSorts As clsSortCollection, _
                   ByVal objSearchFilter As clsSearchFilter, ByVal intNumberRecords As Integer, _
                   ByVal blnForward As Boolean, ByVal intRecordNumber As Integer, ByVal intTypeID As Integer, _
                   ByVal blnIgnoreDataSecurity As Boolean)
        Me.New(objTable, colSorts, objSearchFilter, intNumberRecords, blnForward, intRecordNumber, intTypeID)
        m_blnIgnoreDataSecurity = blnIgnoreDataSecurity
    End Sub

#End Region

#Region " enumSelectType.FIND "

    ''' <summary>
    ''' (enumSelectType.FIND) Will retrieve list of IDs that match objSearchElement
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, _
    ByVal objSearchElement As clsSearchElement)
        m_eSelectType = enumSelectType.FIND
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_colSorts = colSorts
        m_objSearchFilter = objSearchFilter
        m_blnForward = True
        If objSearchElement IsNot Nothing Then AppendToFilter(objSearchElement)
    End Sub
#End Region

#Region " enumSelectType.SELECTION "

    ''' <summary>
    ''' (enumSelectType.SELECTION) Will retrieve a list of IDs between (and including)
    ''' the two records specified (colRecordStart and colRecordEnd)
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, _
    Optional ByVal colRecordStart As clsMaskFieldDictionary = Nothing, _
    Optional ByVal colRecordEnd As clsMaskFieldDictionary = Nothing)
        m_eSelectType = enumSelectType.SELECTION
        m_objTable = objTable
        m_objDB = m_objTable.Database
        m_colSorts = colSorts
        m_objSearchFilter = objSearchFilter
        m_colRecord = colRecordStart
        m_colRecordEnd = colRecordEnd
        m_blnInclRefRecord = True
        m_blnForward = True
    End Sub
#End Region

#Region " Special enumSelectType.SELECTION "

    ''' <summary>
    ''' (enumSelectType.SELECTION) Will retrieve a list of IDs between (and including)
    ''' the two records specified (colRecordStart and colRecordEnd) but you can ignore 
    ''' security
    ''' </summary>
    Public Sub New(ByVal objTable As clsTable, _
    ByVal colSorts As clsSortCollection, _
    ByVal objSearchFilter As clsSearchFilter, ByVal blnIgnoreDataSecurity As Boolean, _
    Optional ByVal colRecordStart As clsMaskFieldDictionary = Nothing, _
    Optional ByVal colRecordEnd As clsMaskFieldDictionary = Nothing)
        Me.new(objTable, colSorts, objSearchFilter, colRecordStart, colRecordEnd)
        m_blnIgnoreDataSecurity = blnIgnoreDataSecurity
    End Sub
#End Region

#End Region

#Region " Properties "

    Friend Property SelectFields As String()
        Get
            Return CType(m_colOnlyFields.Values, String())
        End Get
        Set(value As String())
            For Each strFieldName As String In value
                If (m_colOnlyFields Is Nothing) Then
                    m_colOnlyFields = New Hashtable()
                End If
                m_colOnlyFields.Clear()
                m_colOnlyFields.Add(strFieldName, strFieldName)
            Next
        End Set
    End Property

    ''' <summary>
    ''' This will contain the final recordset
    ''' </summary>
    Public ReadOnly Property DataTable() As DataTable
        Get
            If m_strSQL Is Nothing Then
                Initialize()
            End If

            'Return a datatable from the constructed SQL
            ProcessSQL()

            Return m_objDT
        End Get
    End Property

    ''' <summary>
    ''' This will create the SQL query that can be used to get the final recordset
    ''' </summary>
    Public ReadOnly Property SQL() As String
        Get
            m_colFields.Clear()
            m_colFilters.Clear()
            m_colTableDeclarations.Clear()
            m_colTablesUsages.Clear()
            m_colTableAliases.Clear()
            m_colTempTables.Clear()
            m_intTotalIDs = 0
            m_objDT = Nothing

            If m_objDT Is Nothing Then
                Initialize()
            End If
            Return m_strSQL
        End Get
    End Property

    ''' <summary>
    ''' Should parameters be used in the generated query
    ''' </summary>
    Public Property UseParameters() As Boolean
        Get
            Return m_blnUseParameters
        End Get
        Set(ByVal value As Boolean)
            m_blnUseParameters = value
        End Set
    End Property
    ''' <summary>
    ''' Should ignore wildcards for string searches
    ''' </summary>
    Public Property IgnoreWildCards As Boolean
        Get
            Return m_blnIgnoreWildCards
        End Get
        Set(ByVal value As Boolean)
            m_blnIgnoreWildCards = value
        End Set
    End Property

    Public ReadOnly Property RootTable() As clsTable
        Get
            Return m_objTable
        End Get
    End Property
#End Region

#Region " Methods "

#Region " Build SQL "

    ''' <summary>
    ''' This function builds collection of objects used for constructing the final SQL Select Statement
    ''' </summary>
    Private Sub Initialize()
        'Initialize collection using the base table 
        CreateTableAlias(m_objTable.DatabaseName, m_objTable.DatabaseName, m_objTable.DatabaseName, Nothing, m_colTableDeclarations)

        'Create the fields collection, along with join tables
        InitializeFields()

        'Add the security filter
        AddSecurityFilter()

        'Parse the Main filter and get tables involved in query
        ParseSearchFilter()

        'Add any paging filters to the filters collection
        HandlePagingRecords()

        'Get the order by sql clause
        GetOrderBy()

        'Lastly, build the SQL statement using the constructed collections
        AssembleSQL()

        'Return a datatable from the constructed SQL
        'ProcessSQL()
    End Sub
#End Region

#Region " SQL Construction "
    Private Property blnIsDropDownRequest As Boolean = False

    Private Function GetColFields() As IEnumerable
        If blnIsDropDownRequest Then
            Return m_colFields.Cast(Of String).Where(Function(strField) strField.Contains($"[{clsDBConstants.Fields.cEXTERNALID}]") OrElse strField.Contains($"[{clsDBConstants.Fields.cID}]") OrElse strField.Contains($"[{clsDBConstants.Fields.cSECURITYID}]"))
        End If
        Return m_colFields
    End Function


    ''' <summary>
    ''' Constructs the SQL SELECT Statement using the collections
    ''' </summary>
    Private Sub AssembleSQL()
        'Start with the SELECT statement
        Select Case m_eSelectType
            Case enumSelectType.COUNT, _
            enumSelectType.COUNT_TO_REC
                m_strSQL = "SELECT" & vbCrLf & vbTab & _
                    "COUNT([" & m_objTable.DatabaseName & "].[" & _
                    clsDBConstants.Fields.cID & "]) AS Count" & vbCrLf
                m_colFields.Clear()

            Case enumSelectType.EXPORT, _
            enumSelectType.GOTO_PAGE, _
            enumSelectType.SELECT_FIELDS
                m_strSQL = "SELECT " & vbCrLf & vbTab

            Case enumSelectType.PAGE
                m_strSQL = "SELECT TOP " & m_intNumRecords & vbCrLf & vbTab

            Case enumSelectType.SELECTION, _
            enumSelectType.FIND
                m_strSQL = "SELECT" & vbCrLf & vbTab & _
                    "[" & m_objTable.DatabaseName & "].[" & _
                    clsDBConstants.Fields.cID & "]" & vbCrLf

        End Select

        'Add the Fields if we have a collection of field to add to the SQL
        If m_colFields.Count > 0 Then
            Dim strFields As String = ""
            For Each strField As String In GetColFields()
                AppendToString(strFields, strField, "," & vbCrLf & vbTab)
            Next
            m_strSQL &= strFields & vbCrLf
        End If

        m_strSQL &= "FROM" & vbCrLf & vbTab

        'Special Logic for numbering records in sort order when jumping to specific record
        If m_eSelectType = clsSelectInfo.enumSelectType.GOTO_PAGE Then
            m_strSQL &= "(" & vbCrLf & "SELECT" & vbCrLf & _
                vbTab & "[" & m_objTable.DatabaseName & "].[" & _
                clsDBConstants.Fields.cID & "]," & vbCrLf & vbTab & _
                "ROW_NUMBER() OVER (" & m_strOrderBy & ") AS RowNum" & vbCrLf
            m_strSQL &= "FROM" & vbCrLf & vbTab
        End If

        For Each strTable As String In m_colTableDeclarations
            m_strSQL &= strTable & vbCrLf
        Next

        If m_colFilters IsNot Nothing AndAlso m_colFilters.Count > 0 Then
            m_strSQL &= "WHERE" & vbCrLf
            Dim strWhere As String = ""
            For Each strFilter As String In m_colFilters
                AppendToString(strWhere, vbTab & strFilter, vbCrLf & "AND" & vbCrLf)
            Next
            m_strSQL &= strWhere & vbCrLf
        End If

        If Not m_eSelectType = clsSelectInfo.enumSelectType.GOTO_PAGE Then
            m_strSQL &= m_strOrderBy
        Else
            m_strSQL &= ") #TempTable" & vbCrLf

            For intLoop As Integer = 1 To m_colTableDeclarations.Count
                Dim strTable As String = CType(m_colTableDeclarations(intLoop), String)

                If intLoop = 1 Then
                    m_strSQL &= "INNER JOIN " & strTable & " ON #TempTable.[" & _
                        clsDBConstants.Fields.cID & "] = " & strTable & ".[" & _
                        clsDBConstants.Fields.cID & "]" & vbCrLf
                Else
                    m_strSQL &= strTable & vbCrLf
                End If
            Next

            m_strSQL &= "WHERE #TempTable.RowNum > " & m_intRecordNumber & _
                " AND #TempTable.RowNum <= " & (m_intRecordNumber + _
                m_intNumRecords)
        End If
    End Sub
#End Region

#Region " Table Alias Functionality "

    ''' <summary>
    ''' Will make sure that the table passed in has an alias if necessary,
    ''' it also adds the table declaration to the table declarations collection
    ''' </summary>
    Private Function CreateTableAlias(ByVal strRootRef As String,
                                      ByVal strNewFieldRef As String, ByVal strTable As String,
                                      ByVal objFieldLink As clsFieldLink, ByVal colTableDeclarations As Collection,
                                      Optional ByVal blnLinkedTable As Boolean = False,
                                      Optional ByVal blnForceAddToCollection As Boolean = False) As String
        Dim strAlias As String
        Dim strRootAlias As String

        strAlias = CType(m_colTableAliases(strNewFieldRef), String)
        strRootAlias = CType(m_colTableAliases(strRootRef), String)

        If strAlias Is Nothing Then
            strAlias = strTable

            If m_colTablesUsages(strTable) Is Nothing Then
                m_colTablesUsages.Add(strTable, strTable)
            Else
                strAlias = Nothing
                Dim intInc As Integer = 1
                Dim strTempAlias As String

                While strAlias Is Nothing
                    strTempAlias = strTable & intInc

                    If m_colTablesUsages(strTempAlias) Is Nothing Then
                        m_colTablesUsages.Add(strTempAlias, strTable)
                        strAlias = strTempAlias
                    End If

                    intInc += 1
                End While
            End If

            m_colTableAliases.Add(strNewFieldRef, strAlias)

            If Not blnLinkedTable Then
                blnForceAddToCollection = True
            End If
        End If

        If blnForceAddToCollection Then
            If objFieldLink Is Nothing Then
                colTableDeclarations.Add("[" & strTable & "]")
            Else
                Dim strSQL As String = "LEFT JOIN [" & strTable & "] "
                If Not strAlias = strTable Then
                    strSQL &= "AS [" & strAlias & "] "
                End If

                strSQL &= "ON [" & strRootAlias & "]" & _
                    ".[" & objFieldLink.ForeignKeyField.DatabaseName & "] = " & _
                    "[" & strAlias & "]" & _
                    ".[" & clsDBConstants.Fields.cID & "]"

                colTableDeclarations.Add(strSQL)
            End If
        End If

        Return strAlias
    End Function

#End Region

#Region " Field Selection Functionality "

    '[Naing] Who the fuck writes code like this!!!
    ''' <summary>
    ''' Goes through fields for the table and adds the selectable ones to the fields collection
    ''' </summary>
    Private Sub InitializeFields()

        'the selection types below allow for fields to be selected
        'PAGE = 1 --select fields        
        'EXPORT = 3 --select fields        
        'SELECT_FIELDS = 6 --select fields
        'GOTO_PAGE = 7 --select fields
        If m_eSelectType = enumSelectType.SELECT_FIELDS OrElse
            m_eSelectType = enumSelectType.EXPORT OrElse
            m_eSelectType = enumSelectType.PAGE OrElse
            m_eSelectType = enumSelectType.GOTO_PAGE Then

            'Get ListColumn fields for the table or type if its type dependent
            Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = m_objTable.GetListColumns(m_intTypeID)

            For Each objField As clsField In m_objTable.Fields.Values

                Dim blnAddField As Boolean = False

                'Always include these fields in the select
                If objField.IsIdentityField OrElse
                    objField.DatabaseName = clsDBConstants.Fields.cEXTERNALID OrElse
                    objField.DatabaseName = clsDBConstants.Fields.cSECURITYID Then

                    blnAddField = True

                End If

                'Always include if its export and the field type is Binary then do not include
                If m_eSelectType = enumSelectType.EXPORT AndAlso
                    (m_blnExcludeBinary AndAlso objField.IsBinaryType) Then
                    Continue For
                End If

                'Always include Suffix field of EDOC table  if Show Image In List is true and select type is page or go to page
                If (m_eSelectType = enumSelectType.PAGE OrElse
                    m_eSelectType = enumSelectType.GOTO_PAGE) AndAlso
                m_objTable.DatabaseName = clsDBConstants.Tables.cEDOC AndAlso
                m_objDB.SysInfo.K1Configuration.ShowImageColInList AndAlso
                objField.DatabaseName = clsDBConstants.Fields.EDOC.cSUFFIX Then

                    blnAddField = True

                End If

                If m_colOnlyFields.Count > 0 Then 'Check to see if client specified s set of fields for select

                    Dim hasAccess = (m_objDB.Profile.HasAccess(objField.SecurityID) AndAlso
                                     m_objDB.Profile.LinkFields(objField.KeyID) IsNot Nothing) OrElse
                                 m_blnIgnoreFieldSecurity

                    'Not included yet? We must add this field if current user 
                    'has access and it was specified in the set for selection

                    '2017-09-11 -- Peter Melisi -- Bug fix for #1700003358
                    If Not blnAddField AndAlso hasAccess AndAlso
                        (m_colOnlyFields(objField.DatabaseName) IsNot Nothing OrElse
                         (m_blnUseFKeyExternalIDs AndAlso
                          objField.IsForeignKey AndAlso
                          m_colSorts.Contains(objField.DatabaseName))) Then

                        blnAddField = True

                    End If

                    'If Not blnAddField AndAlso
                    '    m_objDB.Profile.HasAccess(objField.SecurityID) AndAlso
                    '    m_objDB.Profile.SecurityGroup.LinkFields(objField.KeyID) IsNot Nothing AndAlso
                    '    (m_colOnlyFields(objField.DatabaseName) IsNot Nothing OrElse
                    '     (m_blnUseFKeyExternalIDs AndAlso
                    '      objField.IsForeignKey AndAlso
                    '      m_colSorts(objField.DatabaseName) IsNot Nothing)) Then

                    '    blnAddField = True

                    'End If

                Else 'Client did not specify any fields for select. I.E. "select *"

                    Dim hasAccess = (m_objDB.Profile.HasAccess(objField.SecurityID) AndAlso
                                     m_objDB.Profile.LinkFields(objField.KeyID) IsNot Nothing) OrElse
                                 m_blnIgnoreFieldSecurity

                    '2017-09-11 -- Peter Melisi -- Bug fix for #1700003358
                    If Not blnAddField AndAlso
                            objField.IsVisible AndAlso hasAccess AndAlso
                            Not objField.IsInvalidListType AndAlso
                            ((Not m_eSelectType = enumSelectType.PAGE AndAlso
                              Not m_eSelectType = enumSelectType.GOTO_PAGE) OrElse
                          colListColumns(objField.TableID & "_" & objField.ID) IsNot Nothing OrElse
                          (objField.IsForeignKey AndAlso m_colSorts.Contains(objField.DatabaseName))) Then

                        blnAddField = True

                    End If

                    'If Not blnAddField AndAlso
                    '       objField.IsVisible AndAlso
                    '       m_objDB.Profile.HasAccess(objField.SecurityID) AndAlso
                    '       m_objDB.Profile.SecurityGroup.LinkFields(objField.KeyID) IsNot Nothing AndAlso
                    '       Not objField.IsInvalidListType AndAlso
                    '       ((Not m_eSelectType = enumSelectType.PAGE AndAlso
                    '         Not m_eSelectType = enumSelectType.GOTO_PAGE) OrElse
                    '     colListColumns(objField.TableID & "_" & objField.ID) IsNot Nothing OrElse
                    '     (objField.IsForeignKey AndAlso m_colSorts(objField.DatabaseName) IsNot Nothing)) Then

                    '    blnAddField = True

                    'End If

                End If

                If blnAddField Then
                    If objField.IsForeignKey AndAlso m_blnUseFKeyExternalIDs Then
                        Dim strAlias As String = CreateTableAlias(m_objTable.DatabaseName,
                                                                  m_objTable.DatabaseName & "." & objField.DatabaseName,
                                                                  objField.FieldLink.IdentityTable.DatabaseName,
                                                                  objField.FieldLink,
                                                                  m_colTableDeclarations)

                        Dim strField = String.Format("[{0}].[{1}] AS [{2}]",
                                                     strAlias,
                                                     clsDBConstants.Fields.cEXTERNALID,
                                                     objField.DatabaseName)
                        m_colFields.Add(strField)

                    Else
                        If m_eSelectType = enumSelectType.PAGE AndAlso
                            (objField.DataType = SqlDbType.NText OrElse
                             objField.DataType = SqlDbType.Text) Then

                            m_colFields.Add("CONVERT(NVARCHAR(100),[" & m_objTable.DatabaseName & "].[" & objField.DatabaseName & "]) AS " & objField.DatabaseName)

                        Else

                            '[Naing] 11/03/2013 This was added for Computed Fields. Initial attempt to only work with Task (ToDoList) only!
                            'These fields do not map to a physical column but to a scalar function.
                            If objField.IsComputedField Then
                                If (String.IsNullOrEmpty(objField.ComputeFunction)) Then
                                    'Always opt for convention over configuration, therefore this is the default convention for computed field scalar functions.
                                    objField.ComputeFunction = String.Format("[dbo].[{0}_ComputedColumn]([{0}].[{1}]) as [{2}] ",
                                                                             objField.Table.DatabaseName,
                                                                             clsDBConstants.Fields.cID,
                                                                             objField.CaptionText)

                                End If
                                m_colFields.Add(objField.ComputeFunction)
                            End If
                            m_colFields.Add("[" & m_objTable.DatabaseName & "].[" & objField.DatabaseName & "]")

                        End If
                    End If
                End If
            Next

        Else

            'COUNT = 0
            'COUNT_TO_REC = 2
            'SELECTION = 4 -- IDs Only
            'FIND = 5 -- IDs Only
            If m_colSorts IsNot Nothing Then
                For Each objSort As clsSort In m_colSorts

                    If objSort.Field.IsForeignKey Then

                        CreateTableAlias(m_objTable.DatabaseName,
                                         m_objTable.DatabaseName & "." & objSort.Field.DatabaseName,
                                         objSort.Field.FieldLink.IdentityTable.DatabaseName,
                                         objSort.Field.FieldLink,
                                         m_colTableDeclarations)
                    End If

                Next
            End If

        End If

    End Sub

#End Region

#Region " Add Record Security "

    ''' <summary>
    ''' Add the record security check to the SQL
    ''' </summary>
    Private Sub AddSecurityFilter()
        If m_blnIgnoreDataSecurity Then
            Return
        End If

        Dim strFilter As String = "("
        Dim objProfile As clsUserProfile = m_objDB.Profile

        m_intTotalIDs += objProfile.LinkSecurities.Values.Count

        Dim strIDs As String = CreateIDStringFromCollection( _
            objProfile.LinkSecurities.Values)

        If strIDs IsNot Nothing AndAlso strIDs.Length > 0 Then
            strFilter &= "[" & m_objTable.DatabaseName & "].[" & _
                clsDBConstants.Fields.cSECURITYID & "] IN (" & strIDs & ") OR "
        End If
        strFilter &= "[" & m_objTable.DatabaseName & "].[" & _
            clsDBConstants.Fields.cSECURITYID & "] IS NULL)"

        m_colFilters.Add(strFilter)
    End Sub
#End Region

#Region " Parse Main Filter "

    ''' <summary>
    ''' Will build a SQL Filter based on the Search Filter provided (if provided)
    ''' </summary>
    Private Sub ParseSearchFilter()
        If m_objSearchFilter IsNot Nothing Then
            Dim strFilter As String = ""

            RecurseSearchFilter(m_objSearchFilter.Group, strFilter, "", True)

            m_colFilters.Add(strFilter)

        End If
    End Sub

    ''' <summary>
    ''' Recurses through the search filter's search groups and search elements
    ''' </summary>
    Private Sub RecurseSearchFilter(ByVal objSearchObj As clsSearchObjBase,
                                    ByRef strFilter As String,
                                    ByVal strTabs As String,
                                    ByVal blnFirstOperator As Boolean)

        If objSearchObj.OperatorType = clsSearchFilter.enumOperatorType.NOT OrElse
            Not blnFirstOperator Then
            strFilter &= strTabs & clsSearchFilter.GetSQLOperator(objSearchObj.OperatorType) & vbCrLf
        End If

        If TypeOf objSearchObj Is clsSearchGroup Then

            Dim objSearchGroup As clsSearchGroup = CType(objSearchObj, clsSearchGroup)

            If strTabs.Length = 0 Then
                strFilter &= "(" & vbCrLf
            Else
                strFilter &= strTabs & vbTab & "(" & vbCrLf
            End If

            strTabs = strTabs & vbTab

            blnFirstOperator = True
            For Each objSubSearchObj As clsSearchObjBase In objSearchGroup.SearchObjs
                RecurseSearchFilter(objSubSearchObj, strFilter, strTabs, blnFirstOperator)
                blnFirstOperator = False
            Next

            strFilter &= strTabs & ")" & vbCrLf
        Else
            HandleSearchElement(CType(objSearchObj, clsSearchElement), strFilter, strTabs & vbTab)
            strFilter &= vbCrLf
        End If

    End Sub

    ''' <summary>
    ''' Functionality to handle a search element
    ''' </summary>
    Private Sub HandleSearchElement(ByVal objSearchElement As clsSearchElement,
                                    ByRef strFilter As String,
                                    ByVal strTabs As String)

        Dim arrRefs As String() = Split(objSearchElement.FieldRef, ".")
        Dim objTable As clsTable
        Dim objField As clsField = Nothing
        Dim strRef As String
        Dim strRootRef As String
        Dim strValue As String
        Dim strAlias As String
        Dim blnForeignKeyNotInTable As Boolean = False
        Dim colTableDeclarations As Collection = m_colTableDeclarations
        Dim strLinkTableWhereClause As String = ""
        Dim intLinkTableCount As Integer = 0
        Dim strMoreTabs As String = ""

        objTable = m_objDB.SysInfo.Tables(arrRefs(0))
        strAlias = objTable.DatabaseName
        strRef = objTable.DatabaseName
        strRootRef = strRef

        If Not objTable.DatabaseName = m_objTable.DatabaseName Then
            Throw New Exception("Invalid Filter")
        End If

        For intLoop As Integer = 1 To arrRefs.Length - 1
            strValue = arrRefs(intLoop)

            If strValue(0) = "*"c Then 'This is a linked table (where fkey is in child or link table)

                If strValue.Length <= 1 OrElse intLoop = arrRefs.Length - 1 Then
                    Throw New Exception("Invalid Filter")
                End If

                strValue = strValue.Substring(1, strValue.Length - 1)

                objTable = m_objDB.SysInfo.Tables(strValue)

                strRef &= ".*" & objTable.DatabaseName

                blnForeignKeyNotInTable = True

                If intLinkTableCount > 0 Then
                    strMoreTabs &= vbTab
                End If

                intLinkTableCount += 1
            Else
                objField = m_objDB.SysInfo.Fields(objTable.ID & "_" & strValue)

                If objField Is Nothing Then
                    Throw New Exception("Field does not exist: " & strValue)
                End If

                strRef &= "." & objField.DatabaseName

                If objField.IsForeignKey AndAlso Not intLoop = arrRefs.Length - 1 Then
                    Dim strTable As String

                    If blnForeignKeyNotInTable Then
                        strTable = objField.FieldLink.ForeignKeyTable.DatabaseName
                    Else
                        strTable = objField.FieldLink.IdentityTable.DatabaseName
                        objTable = m_objDB.SysInfo.Tables(strTable)
                    End If

                    If blnForeignKeyNotInTable Then
                        strAlias = CreateTableAlias(strRootRef, strRef, strTable, _
                            objField.FieldLink, colTableDeclarations, True)

                        If intLinkTableCount > 1 Then
                            For Each strTableDec As String In colTableDeclarations
                                strFilter &= strTabs & strMoreTabs & strTableDec & vbCrLf
                            Next
                            strFilter &= strLinkTableWhereClause
                        End If

                        colTableDeclarations = New Collection

                        strFilter &= strTabs & strMoreTabs & "EXISTS (" & vbCrLf & strTabs & strMoreTabs & _
                            "SELECT [" & strAlias & "].[" & clsDBConstants.Fields.cID & "] " & _
                            "FROM [" & strTable & "] AS [" & strAlias & "]" & vbCrLf
                        strLinkTableWhereClause = strTabs & strMoreTabs & "WHERE [" & strAlias & "].[" & objField.FieldLink.ForeignKeyField.DatabaseName & "] = " & "[" & CType(m_colTableAliases(strRootRef), String) & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & strTabs & strMoreTabs & "AND " & vbCrLf
                    Else
                        strAlias = CreateTableAlias(strRootRef, strRef, strTable, _
                            objField.FieldLink, colTableDeclarations, False, (intLinkTableCount > 0))
                    End If

                    strRootRef = strRef
                End If

                blnForeignKeyNotInTable = False
            End If
        Next

        If objField Is Nothing OrElse
            Not objField.Table.DatabaseName = objTable.DatabaseName Then
            Throw New Exception("Invalid Filter")
        End If

        If intLinkTableCount >= 1 Then
            strFilter = colTableDeclarations.Cast(Of String)().Aggregate(strFilter, Function(current, strTableDec) current & (vbCrLf & strTabs & strMoreTabs & strTableDec))
            strFilter &= strLinkTableWhereClause
        End If

        strFilter &= strTabs & strMoreTabs & GetFilter(objSearchElement,
                                                       objField,
                                                       "[" & strAlias & "].[" & objField.DatabaseName & "]",
                                                       False)

        If intLinkTableCount >= 1 Then
            strFilter &= vbCrLf
            For intLoop As Integer = intLinkTableCount - 1 To 0 Step -1
                'vbTab
                strFilter &= strTabs & "".PadRight(intLoop, Chr(9)) & ")" & vbCrLf
            Next
        End If

    End Sub

#End Region

#Region " Paging Filter Functionality "

    ''' <summary>
    ''' If this is being used for paging, this will build the SQL depending on the start record
    ''' If this is a selection type, you can also specify an end record
    ''' </summary>
    Private Sub HandlePagingRecords()
        If m_eSelectType = clsSelectInfo.enumSelectType.PAGE OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.SELECTION OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.COUNT_TO_REC Then

            If m_colRecord IsNot Nothing Then
                GetPagingFilter(m_colRecord, m_blnForward)
            End If

            If m_eSelectType = clsSelectInfo.enumSelectType.SELECTION AndAlso _
            m_colRecordEnd IsNot Nothing Then
                GetPagingFilter(m_colRecordEnd, Not m_blnForward)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Builds SQL filters based on direction and record
    ''' </summary>
    Private Sub GetPagingFilter(ByVal objRecord As clsMaskFieldDictionary, _
    ByVal blnForward As Boolean)
        Dim colSearchElements As New List(Of clsSearchObjBase)

        Dim strFilter As String = "(" & vbCrLf

        For intLoop As Integer = 0 To m_colSorts.Count - 1
            Dim objSort As clsSort = m_colSorts(intLoop)
            Dim objField As clsField = objSort.Field

            Dim strFieldRef As String = objField.Table.DatabaseName
            Dim blnIsForeignKey As Boolean = objField.IsForeignKey
            If blnIsForeignKey Then
                strFieldRef = objField.Table.DatabaseName & "." & objField.DatabaseName
                objField = m_objDB.SysInfo.Fields(objField.FieldLink.IdentityTable.ID & _
                    "_" & clsDBConstants.Fields.cEXTERNALID)
            End If

            Dim eCompareType As clsSearchFilter.enumComparisonType
            If (blnForward AndAlso objSort.Ascending) OrElse _
            (Not blnForward AndAlso Not objSort.Ascending) Then
                If intLoop = m_colSorts.Count - 1 AndAlso m_blnInclRefRecord Then
                    eCompareType = clsSearchFilter.enumComparisonType.GREATER_THAN_EQUAL
                Else
                    eCompareType = clsSearchFilter.enumComparisonType.GREATER_THAN
                End If
            Else
                If intLoop = m_colSorts.Count - 1 AndAlso m_blnInclRefRecord Then
                    eCompareType = clsSearchFilter.enumComparisonType.LESS_THAN_EQUAL
                Else
                    eCompareType = clsSearchFilter.enumComparisonType.LESS_THAN
                End If
            End If

            Dim objValue As Object
            If blnIsForeignKey Then
                objValue = objRecord(objSort.Field.DatabaseName).Value1.Display
            Else
                ' Ara Melkonian - 2100003649
                ' Adding the inverse relationship
                If objField.FormatType = clsDBConstants.enumFormatType.FileSize Then
                    objValue = objRecord(objField.DatabaseName).Value2.Value
                Else
                    objValue = objRecord(objField.DatabaseName).Value1.Value
                End If
            End If

            Dim objSE As New clsSearchElement(clsSearchFilter.enumOperatorType.AND, _
                strFieldRef, eCompareType, objValue, objField)

            colSearchElements.Add(objSE)

            Dim strTest As String = ""
            strTest &= Nothing

            Dim strPageFilter As String = ""
            For intIndex As Integer = 0 To colSearchElements.Count - 1
                Dim objSearchElement As clsSearchElement = _
                    CType(colSearchElements(intIndex), clsSearchElement)

                Dim strAlias As String = CType(m_colTableAliases(objSearchElement.FieldRef), String)
                Dim strField As String = "[" & strAlias & "].[" & objSearchElement.Field.DatabaseName & "]"
                Dim strTempFilter As String = GetFilter(objSearchElement, objSearchElement.Field, strField, True)

                If strTempFilter IsNot Nothing Then
                    If objSearchElement.CompareType = clsSearchFilter.enumComparisonType.LESS_THAN Then
                        strTempFilter = "(" & strTempFilter & " OR " & strField & " IS NULL)"
                    End If

                    strPageFilter &= strTempFilter

                    If Not intIndex = colSearchElements.Count - 1 Then
                        strPageFilter &= vbCrLf & vbTab & vbTab & _
                            clsSearchFilter.GetSQLOperator(objSearchElement.OperatorType) & _
                            vbCrLf & vbTab & vbTab
                    End If
                Else
                    strPageFilter = Nothing
                    Exit For
                End If
            Next

            objSE.CompareType = clsSearchFilter.enumComparisonType.EQUAL

            If Not String.IsNullOrEmpty(strPageFilter) Then
                strFilter &= vbTab & vbTab & "(" & strPageFilter & ")"

                If Not intLoop = m_colSorts.Count - 1 Then
                    strFilter &= vbCrLf & vbTab & "OR" & vbCrLf
                End If
            End If
        Next
        strFilter &= vbCrLf & vbTab & ")"

        m_colFilters.Add(strFilter)
    End Sub

#End Region

#Region " Get Order By Clause "

    ''' <summary>
    ''' constructs a SQL Order By statement based off the sort objects
    ''' </summary>
    Private Sub GetOrderBy()
        m_strOrderBy = ""

        If m_colSorts IsNot Nothing AndAlso _
        m_colSorts.Count > 0 AndAlso _
        (m_eSelectType = clsSelectInfo.enumSelectType.EXPORT OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.PAGE OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.FIND OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.SELECT_FIELDS OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.GOTO_PAGE OrElse _
        m_eSelectType = clsSelectInfo.enumSelectType.SELECTION) Then
            m_strOrderBy = "ORDER BY" & vbCrLf & vbTab

            Dim strOrderBy As String = ""
            For intLoop As Integer = 0 To m_colSorts.Count - 1
                Dim objSort As clsSort = m_colSorts(intLoop)
                Dim objField As clsField = objSort.Field

                Dim strFieldRef As String = m_objTable.DatabaseName
                Dim strField As String = objField.DatabaseName

                If objField.IsForeignKey AndAlso m_blnUseFKeyExternalIDs Then
                    strFieldRef &= "." & objField.DatabaseName
                    strField = clsDBConstants.Fields.cEXTERNALID
                End If

                Dim strAlias As String = CType(m_colTableAliases(strFieldRef), String)
                '-- skip if we don't have access to it
                If strAlias Is Nothing Then
                    Continue For
                End If

                'For Fields with date time formats
                If objField.IsDateType AndAlso objField.DateType = clsDBConstants.enumDateTypes.TIME_ONLY Then
                    Dim strAdd As String = ""
                    If (m_blnForward AndAlso Not objSort.Ascending) OrElse _
                    (Not m_blnForward AndAlso objSort.Ascending) Then
                        strAdd = " DESC"
                    End If
                    AppendToCommaString(strOrderBy, DatePartSQL(enumDatePartType.Hour, "[" & strAlias & "].[" & strField & "]") & strAdd)
                    AppendToCommaString(strOrderBy, DatePartSQL(enumDatePartType.Minute, "[" & strAlias & "].[" & strField & "]") & strAdd)
                    AppendToCommaString(strOrderBy, DatePartSQL(enumDatePartType.Second, "[" & strAlias & "].[" & strField & "]") & strAdd)
                Else
                    '[Naing] 12/03/2013 Computed fields don't require table alias since they are not physical column on a table
                    'Caption text is used here so that sorting is supported on computed fields.
                    If (objField.IsComputedField) Then
                        AppendToCommaString(strOrderBy, String.Format("[{0}]", objField.CaptionText))
                    Else
                        AppendToCommaString(strOrderBy, String.Format("[{0}].[{1}]", strAlias, strField))
                    End If
                    If (m_blnForward AndAlso Not objSort.Ascending) OrElse
                        (Not m_blnForward AndAlso objSort.Ascending) Then
                        strOrderBy &= " DESC"
                    End If
                End If
            Next

            m_strOrderBy &= strOrderBy
        End If
    End Sub
#End Region

#Region " Filter Construction Methods "

#Region " AppendToFilter "

    ''' <summary>
    ''' Appends or creates a new search filter for the select statement.
    ''' </summary>
    Private Sub AppendToFilter(ByVal objSE As clsSearchElement)
        Dim colSOs As New List(Of clsSearchObjBase)

        If m_objSearchFilter Is Nothing Then
            colSOs.Add(objSE)
        Else
            colSOs.Add(m_objSearchFilter.Group)
            objSE.OperatorType = clsSearchFilter.enumOperatorType.AND
            colSOs.Add(objSE)
        End If

        m_objSearchFilter = New clsSearchFilter(m_objDB, _
            New clsSearchGroup(clsSearchFilter.enumOperatorType.NONE, colSOs), _
            m_objTable.DatabaseName)
    End Sub

#End Region

#Region " GetFilter "

    ''' <summary>
    ''' Returns a filter based off the data in the search element
    ''' </summary>
    Public Function GetFilter(ByVal objSearchElement As clsSearchElement,
                              ByVal objField As clsField,
                              ByVal strField As String,
                              ByVal blnPagingFilter As Boolean) As String

        Dim strFilter As String = ""
        Dim strOp As String = ""
        Dim strSubOp As String = ""
        Dim blnRangeEnd As Boolean = False

        Select Case objSearchElement.CompareType
            '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
            Case clsSearchFilter.enumComparisonType.EQUAL, clsSearchFilter.enumComparisonType.EXACT
                If objSearchElement.Value Is Nothing Then
                    Return strField & " IS NULL"
                End If

                strOp = " = "
            Case clsSearchFilter.enumComparisonType.GREATER_THAN
                If objSearchElement.Value Is Nothing Then
                    Return strField & " IS NOT NULL"
                End If

                strOp = " > "
                strSubOp = " > "
            Case clsSearchFilter.enumComparisonType.GREATER_THAN_EQUAL
                If objSearchElement.Value Is Nothing Then
                    Return Nothing
                End If

                strOp = " >= "
                strSubOp = " > "
            Case clsSearchFilter.enumComparisonType.LESS_THAN
                If objSearchElement.Value Is Nothing Then
                    Return Nothing
                End If

                blnRangeEnd = True
                strOp = " < "
                strSubOp = " < "
            Case clsSearchFilter.enumComparisonType.LESS_THAN_EQUAL

                If objSearchElement.Value Is Nothing Then
                    Return strField & " IS NULL"
                End If

                blnRangeEnd = True
                strOp = " <= "
                strSubOp = " < "

            Case clsSearchFilter.enumComparisonType.IN

                If objSearchElement.Value Is Nothing OrElse
                    (TypeOf objSearchElement.Value Is Hashtable AndAlso
                     CType(objSearchElement.Value, Hashtable).Count = 0) Then

                    Return strField & " IS NULL"

                End If

        End Select

        Select Case objField.DataType

            Case SqlDbType.NVarChar, SqlDbType.NChar, SqlDbType.NText, SqlDbType.Char, SqlDbType.Text, SqlDbType.VarChar

                Dim strValue As String = CStr(objSearchElement.Value)

                Dim blnUnicode As Boolean = (objField.DataType = SqlDbType.NVarChar OrElse _
                    objField.DataType = SqlDbType.NChar OrElse objField.DataType = SqlDbType.NText)

                Select Case objSearchElement.CompareType

                    '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
                    Case clsSearchFilter.enumComparisonType.EXACT
                        If blnPagingFilter = False Then
                            strField = "ISNULL(" & strField & ", '')"
                        End If

                        strFilter = "RTRIM(" & strField & ") = '" & strValue.TrimEnd & "'"

                    Case clsSearchFilter.enumComparisonType.EQUAL
                        If blnPagingFilter = False Then
                            strField = "ISNULL(" & strField & ", '')"
                        End If

                        strFilter = "RTRIM(" & strField & ") LIKE " & EscapeWildcardChars(strValue.TrimEnd, blnUnicode, blnPagingFilter)

                    Case clsSearchFilter.enumComparisonType.IN
                        If blnPagingFilter = False Then
                            strField = "ISNULL(" & strField & ", '')"
                        End If

                        strFilter = strField & " IN ('" & clsDB.SQLString(strValue) & "')"

                    Case Else
                        strValue = RangeReplace(strValue, blnRangeEnd, blnUnicode, blnPagingFilter)
                        If Not m_blnUseParameters Then
                            strValue = "'" & strValue & "'"
                        End If
                        strFilter = strField & strOp & strValue

                End Select

            Case SqlDbType.DateTime, SqlDbType.SmallDateTime

                Dim dtDate As DateTime = CDate(objSearchElement.Value)

                '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
                If ((objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EXACT) Or (objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EQUAL)) Then
                    Select Case objField.DateType
                        Case clsDBConstants.enumDateTypes.DATE_AND_TIME, clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME, _
                            clsDBConstants.enumDateTypes.DATE_ONLY, clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY 'localtimedate
                            strFilter = "(" & DatePartSQL(enumDatePartType.Year, strField) & " = " & CreateParameter(dtDate.Year.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Month, strField) & " = " & CreateParameter(dtDate.Month.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Day, strField) & " = " & CreateParameter(dtDate.Day.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Hour, strField) & " = " & CreateParameter(dtDate.Hour.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Minute, strField) & " = " & CreateParameter(dtDate.Minute.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Second, strField) & " = " & CreateParameter(dtDate.Second.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Millesecond, strField) & " = " & CreateParameter(dtDate.Millisecond.ToString) & ")"
                            'Case clsDBConstants.enumDateTypes.DATE_ONLY, clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY 'localtimedate
                            '    strFilter = "(" & DatePartSQL(enumDatePartType.Year, strField) & " = " & CreateParameter(dtDate.Year.ToString) & " AND " & _
                            '        "" & DatePartSQL(enumDatePartType.Month, strField) & " = " & CreateParameter(dtDate.Month.ToString) & " AND " & _
                            '        "" & DatePartSQL(enumDatePartType.Day, strField) & " = " & CreateParameter(dtDate.Day.ToString) & " )"
                        Case clsDBConstants.enumDateTypes.TIME_ONLY
                            strFilter = "(" & DatePartSQL(enumDatePartType.Hour, strField) & " = " & CreateParameter(dtDate.Hour.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Minute, strField) & " = " & CreateParameter(dtDate.Minute.ToString) & " AND " & _
                                "" & DatePartSQL(enumDatePartType.Second, strField) & " = " & CreateParameter(dtDate.Second.ToString) & ")"
                    End Select
                Else
                    If String.IsNullOrEmpty(strOp) Then
                        strOp = " = "
                    End If

                    Select Case objField.DateType
                        Case clsDBConstants.enumDateTypes.DATE_AND_TIME, clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME, _
                        clsDBConstants.enumDateTypes.DATE_ONLY, clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY
                            Dim strValue As String = "'" & dtDate.ToString("MMMM dd, yyyy HH:mm:ss.fff") & "'"
                            If m_blnUseParameters Then
                                strValue = CreateParameter(dtDate.ToString("MMMM dd, yyyy HH:mm:ss.fff"))
                            End If
                            strFilter = strField & strOp & "CONVERT(DATETIME, " & strValue & ")"
                            'Case clsDBConstants.enumDateTypes.DATE_ONLY, clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY 'localtimedate
                            '    Dim strValue As String = "'" & dtDate.ToString("MMMM dd, yyyy") & "'"
                            '    If m_blnUseParameters Then
                            '        strValue = CreateParameter(dtDate.ToString("MMMM dd, yyyy"))
                            '    End If
                            '    strFilter = strField & strOP & "CONVERT(DATETIME, " & strValue & ")"
                        Case clsDBConstants.enumDateTypes.TIME_ONLY
                            strFilter = "(" & DatePartSQL(enumDatePartType.Hour, strField) & _
                                strSubOp & CreateParameter(dtDate.Hour.ToString) & _
                                " OR " & vbCrLf & "(" & DatePartSQL(enumDatePartType.Hour, strField) & _
                                " = " & CreateParameter(dtDate.Hour.ToString) & _
                                " AND " & DatePartSQL(enumDatePartType.Minute, strField) & _
                                strSubOp & CreateParameter(dtDate.Minute.ToString) & ")" & _
                                " OR " & vbCrLf & "(" & DatePartSQL(enumDatePartType.Hour, strField) & _
                                " = " & CreateParameter(dtDate.Hour.ToString) & _
                                " AND " & DatePartSQL(enumDatePartType.Minute, strField) & _
                                " = " & CreateParameter(dtDate.Minute.ToString) & _
                                " AND " & DatePartSQL(enumDatePartType.Second, strField) & _
                                strOp & CreateParameter(dtDate.Second.ToString) & "))"
                    End Select
                End If

            Case SqlDbType.Bit

                Dim intVal As Integer = Math.Abs(CInt(objSearchElement.Value))

                If blnPagingFilter = False Then
                    If intVal = 1 Then
                        strField = "ISNULL(" & strField & ", 0)"
                    Else
                        strField = "ISNULL(" & strField & ", 1)"
                    End If
                End If

                If String.IsNullOrEmpty(strOp) Then
                    strOp = " = "
                End If

                strFilter = strField & strOp & CreateParameter(intVal.ToString())

            Case SqlDbType.Image

                Dim strValue As String = CStr(objSearchElement.Value).Trim

                strValue = strValue.Replace(""""c, """""").Replace("'"c, "''")

                If strValue.Last = "*"c Then
                    strValue = """" & strValue & """"
                Else
                    strValue = "FORMSOF(thesaurus,""" & strValue & """)"
                End If

                If Not m_blnUseParameters Then
                    strValue = "'" & strValue & "'"
                End If

                strFilter = "CONTAINS((" & strField & "), " & CreateParameter(strValue) & ")"

            Case SqlDbType.BigInt, SqlDbType.Int, SqlDbType.SmallInt, SqlDbType.TinyInt

                If objSearchElement.CompareType = clsSearchFilter.enumComparisonType.IN Then
                    If blnPagingFilter = False Then
                        strField = "ISNULL(" & strField & ", 0)"
                    End If

                    If TypeOf objSearchElement.Value Is Hashtable Then
                        Dim colIDs As Hashtable = CType(objSearchElement.Value, Hashtable)

                        If m_intTotalIDs + colIDs.Values.Count > clsDBConstants.cMAX_IDS_BEFORE_TEMPTABLE Then
                            strFilter = strField & " IN (SELECT [ID] FROM " & _
                                CreateTempTable(ConvertIDsToDataTable(colIDs)) & ")"
                        Else
                            strFilter = strField & " IN (" & CreateIDStringFromCollection(colIDs.Values) & ")"
                            m_intTotalIDs += colIDs.Values.Count
                        End If
                    Else
                        strFilter = strField & " IN (" & DirectCast(objSearchElement.Value, String) & ")"
                    End If
                    '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
                ElseIf ((objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EXACT) Or (objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EQUAL)) Then
                    If blnPagingFilter = False Then
                        If CInt(objSearchElement.Value) = 0 Then
                            strField = "ISNULL(" & strField & ", 1)"
                        Else
                            strField = "ISNULL(" & strField & ", 0)"
                        End If
                    End If

                    strFilter = strField & strOp & CreateParameter(CStr(objSearchElement.Value), objField.DataType)
                Else
                    strFilter = strField & strOp & CreateParameter(CStr(objSearchElement.Value), objField.DataType)
                End If

            Case Else
                '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
                If blnPagingFilter = False AndAlso
                    ((objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EXACT) Or (objSearchElement.CompareType = clsSearchFilter.enumComparisonType.EQUAL)) AndAlso
                    (objField.DataType = SqlDbType.Decimal OrElse
                     objField.DataType = SqlDbType.Float OrElse
                     objField.DataType = SqlDbType.Money OrElse
                     objField.DataType = SqlDbType.Real OrElse
                     objField.DataType = SqlDbType.SmallMoney) Then

                    If CInt(objSearchElement.Value) = 0 Then
                        strField = "ISNULL(" & strField & ", 1)"
                    Else
                        strField = "ISNULL(" & strField & ", 0)"
                    End If
                End If

                strFilter = strField & strOp & CreateParameter(CStr(objSearchElement.Value), objField.DataType)

        End Select

        Return strFilter

    End Function

    Private Function CreateParameter(ByVal strValue As String) As String
        Dim strName As String = strValue

        If m_blnUseParameters Then
            If m_colParams Is Nothing Then
                m_colParams = New clsDBParameterDictionary
            End If

            strName = "@Var" & m_colParams.Count + 1
            m_colParams.Add(New clsDBParameter(strName, strValue))
        End If

        Return strName
    End Function

    Private Function CreateParameter(ByVal strValue As String, ByVal sqlDbType As SqlDbType) As String
        Dim strName As String = strValue

        If m_blnUseParameters Then
            If m_colParams Is Nothing Then
                m_colParams = New clsDBParameterDictionary
            End If

            strName = "@Var" & m_colParams.Count + 1
            m_colParams.Add(New clsDBParameter(strName, strValue, ParameterDirection.Input, sqlDbType))
        End If

        Return strName
    End Function
#End Region

#Region " CreateTempTable "

    ''' <summary>
    ''' Adds a temp table to the temp table collection
    ''' </summary>
    ''' <remarks>
    ''' It is necessary to use temp tables when we come accross the scenario where we select
    ''' an ID that is contained IN a large set of values.  SQL Server currently has a limit on
    ''' the number of constants\variables that can be used in a SQL Statement, and so the temp
    ''' tables are a way to circumvent this limitation
    ''' </remarks>
    Private Function CreateTempTable(ByVal objDT As DataTable) As String
        Dim strTempTable As String = "#tmpTable" & m_colTempTables.Count + 1

        m_colTempTables.Add(strTempTable, objDT)

        Return strTempTable
    End Function
#End Region

#Region " Common "

    ''' <summary>
    ''' Returns the SQL DatePart function
    ''' </summary>
    Public Shared Function DatePartSQL(ByVal eDatePartType As enumDatePartType, _
    ByVal strField As String) As String
        Dim strDatePart As String

        Select Case eDatePartType
            Case enumDatePartType.Year
                strDatePart = "yyyy"
            Case enumDatePartType.Month
                strDatePart = "mm"
            Case enumDatePartType.Day
                strDatePart = "dd"
            Case enumDatePartType.Hour
                strDatePart = "hh"
            Case enumDatePartType.Minute
                strDatePart = "n"
            Case enumDatePartType.Second
                strDatePart = "s"
            Case enumDatePartType.Millesecond
                strDatePart = "ms"
            Case Else
                strDatePart = ""
        End Select

        Return "DATEPART(" & strDatePart & ", " & strField & ")"
    End Function

    ''' <summary>
    ''' Converts a text value to its necessary range equivalent
    ''' </summary>
    ''' <param name="strText"></param>
    ''' <param name="blnHighRange"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RangeReplace(ByVal strText As String, ByVal blnHighRange As Boolean, _
    ByVal blnUnicodeField As Boolean, ByVal blnPagingFilter As Boolean) As String
        Dim strNewString As String = strText

        If Not blnPagingFilter Then
            Dim intIndex As Integer = strText.IndexOf("*")

            If intIndex >= 0 Then
                'strip off the *
                strNewString = strText.Substring(0, intIndex)
            End If

            'If blnUnicodeField Then
            '    strNewString = "N'" & strNewString
            'Else
            '    strNewString = "'" & strNewString
            'End If

            If blnHighRange AndAlso intIndex >= 0 Then
                strNewString &= "zzzzzzzzzzzzzzzzzzz"
            End If
        End If
        'strNewString &= "'"

        Return CreateParameter(strNewString)
    End Function

    ''' <summary>
    ''' When using the SQL LIKE statement certain characters need to be prefixed with an escape character
    ''' because they mean something specific to the LIKE function
    ''' </summary>
    Private Function EscapeWildcardChars(ByVal strValue As String, _
    ByVal blnUnicodeField As Boolean, ByVal blnPagingFilter As Boolean) As String
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

            If Not m_blnIgnoreWildCards Then
                If strValue.IndexOf("*"c) < 0 Then
                    strValue = "%" & strValue & "%"
                End If
            End If


            strValue = strValue.Replace("*", "%")
        End If
        Dim strTempValue As String = strValue

        strValue = CreateParameter(strValue)

        If Not Me.UseParameters Then
            '-- need to wrap it in qoutation marks
            strValue = clsDB.SQLString(strValue)
            strValue = String.Format("'{0}'", strValue)
        End If

        If strTempValue.IndexOf(arrEscapeChar(intLoop)) >= 0 Then
            strValue &= " ESCAPE '" & arrEscapeChar(intLoop) & "'"
        End If

        Return strValue
    End Function
#End Region

#End Region

#Region " Process SQL and Return DataTable "

    ''' <summary>
    ''' Gets a datatable from the SQL Statement constructed
    ''' </summary>
    ''' <remarks>
    ''' If the SQL Statement contains temporary tables, this procedure first creates the temp
    ''' tables and populates them prior to running the main SQL Select query
    ''' </remarks>
    Public Sub ProcessSQL()
        Dim blnUseTransaction As Boolean = False

        Try
            If m_colTempTables.Count > 0 Then

                If Not m_objDB.HasTransaction Then
                    blnUseTransaction = True

                    m_objDB.BeginTransaction()
                End If

                For Each strTempTable As String In m_colTempTables.Keys
                    m_objDB.ExecuteSQL("CREATE TABLE [" & strTempTable & "] ([ID] INT)")

                    m_objDB.BulkInsert(CType(m_colTempTables(strTempTable), DataTable), strTempTable)
                Next
            End If

            m_objDT = m_objDB.GetDataTableBySQL(m_strSQL, m_colParams, blnIsDropDownRequest)

            If blnUseTransaction Then
                m_objDB.EndTransaction(True)
            End If
        Catch ex As Exception
            If blnUseTransaction Then
                m_objDB.EndTransaction(False)
            End If

            Throw
        End Try

        If Not m_blnForward Then
            ReverseDataTable(m_objDT)
        End If
    End Sub
#End Region

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
                m_objTable = Nothing
                m_objSearchFilter = Nothing
                m_colSorts = Nothing
                m_colRecord = Nothing
                m_colRecordEnd = Nothing
                m_objDT = Nothing

                If m_colFields IsNot Nothing Then
                    m_colFields.Clear()
                    m_colFields = Nothing
                End If

                If m_colFilters IsNot Nothing Then
                    m_colFilters.Clear()
                    m_colFilters = Nothing
                End If

                If m_colTableDeclarations IsNot Nothing Then
                    m_colTableDeclarations.Clear()
                    m_colTableDeclarations = Nothing
                End If

                If m_colTablesUsages IsNot Nothing Then
                    m_colTablesUsages.Clear()
                    m_colTablesUsages = Nothing
                End If

                If m_colTableAliases IsNot Nothing Then
                    m_colTableAliases.Clear()
                    m_colTableAliases = Nothing
                End If

                If m_colTempTables IsNot Nothing Then
                    m_colTempTables.Clear()
                    m_colTempTables = Nothing
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

End Class
