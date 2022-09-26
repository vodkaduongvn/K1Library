Namespace Database

    Public Class K1DataService

        Enum PagedSelectDirection

            Forward

            BackWard

        End Enum

        Private m_objDb As clsDB

        Sub New(objDb As clsDB)

            m_objDb = objDb

        End Sub

#Region "Count"

        Function [Count](table As clsTable, filter As clsSearchFilter, Optional ignoreSecurity As Boolean = False) As Integer

            Dim objSelectInfo = New clsSelectInfo(table, filter, ignoreSecurity)
            PrintSql(objSelectInfo.SQL)
            Dim resultSet = objSelectInfo.DataTable
            Return CInt(resultSet.Rows(0)("Count"))

        End Function

        Function CountUpToRecord(table As clsTable,
                                 sortBy As clsSortCollection,
                                 filter As clsSearchFilter,
                                 recordId As Integer,
                                 Optional countDirection As PagedSelectDirection = PagedSelectDirection.Forward,
                                 Optional ignoreSecurity As Boolean = False) As Integer

            Dim forward = countDirection = PagedSelectDirection.Forward

            Dim startRecord = clsMaskField.CreateMaskCollection(table, clsTableMask.enumMaskType.VIEW, recordId, CInt(False))

            Return CInt(SelectInternal(table,
                                       sortBy,
                                       filter,
                                       clsDBConstants.cintNULL,
                                       forward,
                                       startRecord,
                                       False,
                                       clsSelectInfo.enumSelectType.COUNT_TO_REC,
                                       clsDBConstants.cintNULL,
                                       ignoreSecurity, Nothing).Rows(0)("Count"))

        End Function

#End Region

#Region "Select to get IDs only"

        ''' <summary>
        ''' Returns IDs of Records between two Records
        ''' </summary>
        ''' <param name="table"></param>
        ''' <param name="filter"></param>
        ''' <param name="sortBy"></param>
        ''' <param name="pageStartRecordId"></param>
        ''' <param name="pageEndRecordId"></param>
        ''' <param name="ignoreSecurity"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function SelectBetween(table As clsTable,
                               filter As clsSearchFilter,
                               sortBy As clsSortCollection,
                               pageStartRecordId As Integer,
                               pageEndRecordId As Integer,
                               Optional ignoreSecurity As Boolean = False) As IEnumerable(Of Integer)

            If (pageStartRecordId = pageEndRecordId) Then
                Throw New ArgumentException(String.Format("{0} can not be the same as {1} of page for a select query.", "pageStartRecordId",
                                                          "pageEndRecordId"))
            End If

            Dim record1 = clsMaskField.CreateMaskCollection(table, clsTableMask.enumMaskType.VIEW, pageStartRecordId, CInt(False))
            Dim record2 = clsMaskField.CreateMaskCollection(table, clsTableMask.enumMaskType.VIEW, pageEndRecordId, CInt(False))

            Dim objSelectInfo = New clsSelectInfo(table, sortBy, filter, ignoreSecurity, record1, record2)

            Return (From objDataRow As DataRow In objSelectInfo.DataTable.Rows.OfType(Of DataRow)()
                Select objDataRow(clsDBConstants.Fields.cID)).Cast(Of Integer)()

        End Function

        ''' <summary>
        ''' Returns IDs of records that matches the criteria in search filter and search element combined
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function SelectFind(ByVal objTable As clsTable,
                            ByVal colSorts As clsSortCollection,
                            ByVal objSearchFilter As clsSearchFilter,
                            ByVal objSearchElement As clsSearchElement) As IEnumerable(Of Integer)

            Dim selectInfo = New clsSelectInfo(objTable, colSorts, objSearchFilter, objSearchElement)
            Dim resultSet As DataTable = selectInfo.DataTable

            Return (From row In resultSet.Rows.OfType(Of DataRow)()
                    Select row(clsDBConstants.Fields.cID)).Cast(Of Integer)()

        End Function

#End Region

        Function SelectAllFieldsByRowNumIndex(table As clsTable,
                                              filter As clsSearchFilter,
                                              pageSize As Integer,
                                              pageStartRowIndex As Integer,
                                              Optional sortBy As clsSortCollection = Nothing,
                                              Optional pageSelectDirection As PagedSelectDirection = PagedSelectDirection.Forward,
                                              Optional typeId As Nullable(Of Integer) = Nothing,
                                              Optional ignoreSecurity As Boolean = False) As DataTable

            If (sortBy Is Nothing) Then

                sortBy = New clsSortCollection()
                sortBy.AddRange(clsSort.GetDefaultSortCollection(table))

            End If

            Dim mtypeId = clsDBConstants.cintNULL
            If (typeId.HasValue) Then
                mtypeId = typeId.Value
            End If

            Dim objSelectInfo As clsSelectInfo = New clsSelectInfo(table, sortBy,
                                                                   filter, pageSize,
                                                                   pageSelectDirection = PagedSelectDirection.Forward, pageStartRowIndex,
                                                                   mtypeId, ignoreSecurity)

            PrintSql(objSelectInfo.SQL)

            Return objSelectInfo.DataTable

        End Function

        Function SelectAllFieldsByPageIndex(table As clsTable,
                                            filter As clsSearchFilter,
                                            pageSize As Integer,
                                            pageIndex As Integer,
                                            Optional sortBy As clsSortCollection = Nothing,
                                            Optional typeId As Nullable(Of Integer) = Nothing,
                                            Optional ignoreSecurity As Boolean = False) As DataTable

            Return SelectAllFieldsByRowNumIndex(table, filter, pageSize,
                                         (pageSize * (pageIndex - 1)), sortBy,
                                         PagedSelectDirection.Forward, typeId, False)


        End Function

        Function SelectSpecificFieldsByRecordId(table As clsTable,
                                                filter As clsSearchFilter,
                                                pageSize As Integer,
                                                pageStartRecordId As Integer,
                                                Optional sortBy As clsSortCollection = Nothing,
                                                Optional fieldsToSelect() As String = Nothing,
                                                Optional includeRelatedData As Boolean = False,
                                                Optional pageSelectDirection As PagedSelectDirection = PagedSelectDirection.Forward,
                                                Optional typeId As Nullable(Of Integer) = Nothing,
                                                Optional ignoreSecurity As Boolean = False) As DataTable

            Dim forward = pageSelectDirection = PagedSelectDirection.Forward

            If (sortBy Is Nothing) Then
                sortBy = clsSort.GetDefaultSortCollection(table)
            End If

            'fetches the data from 
            Dim startRecord = clsMaskField.CreateMaskCollection(table, clsTableMask.enumMaskType.VIEW, pageStartRecordId, CInt(False))

            Return SelectInternal(table,
                                  sortBy,
                                  filter,
                                  pageSize,
                                  forward,
                                  startRecord,
                                  includeRelatedData,
                                  clsSelectInfo.enumSelectType.PAGE,
                                  typeId.Value,
                                  ignoreSecurity,
                                  fieldsToSelect)

        End Function

        Function SelectSpecificFieldsByRowNumIndex(table As clsTable,
                                                   filter As clsSearchFilter,
                                                   pageSize As Integer,
                                                   pageStartRowIndex As Integer,
                                                   fieldsToSelect() As String,
                                                   Optional sortBy As clsSortCollection = Nothing,
                                                   Optional includeRelatedData As Boolean = False,
                                                   Optional pageSelectDirection As PagedSelectDirection = PagedSelectDirection.Forward,
                                                   Optional typeId As Nullable(Of Integer) = Nothing,
                                                   Optional ignoreSecurity As Boolean = False) As DataTable

            If (sortBy Is Nothing) Then
                sortBy = New clsSortCollection()
                sortBy.AddRange(clsSort.GetDefaultSortCollection(table))
            End If

            Dim mtypeId = clsDBConstants.cintNULL
            If (typeId.HasValue) Then
                mtypeId = typeId.Value
            End If

            Dim objSelectInfo As clsSelectInfo = New clsSelectInfo(table, sortBy,
                                                                   filter, pageSize,
                                                                   pageSelectDirection = PagedSelectDirection.Forward, pageStartRowIndex,
                                                                   mtypeId, ignoreSecurity)

            'objSelectInfo.SelectFields = fieldsToSelect

            PrintSql(objSelectInfo.SQL)

            Return objSelectInfo.DataTable

        End Function

        Function SelectSpecificFieldsByPageIndex(table As clsTable,
                                            filter As clsSearchFilter,
                                            pageSize As Integer,
                                            pageIndex As Integer,
                                            fieldsToSelect() As String,
                                            Optional sortBy As clsSortCollection = Nothing,
                                            Optional includeRelatedData As Boolean = False,
                                            Optional pageSelectDirection As PagedSelectDirection = PagedSelectDirection.Forward,
                                            Optional typeId As Nullable(Of Integer) = Nothing,
                                            Optional ignoreSecurity As Boolean = False) As DataTable

            Return SelectSpecificFieldsByRowNumIndex(table,
                                              filter,
                                              pageSize,
                                              (pageSize * (pageIndex - 1)),
                                              fieldsToSelect,
                                              sortBy,
                                              includeRelatedData,
                                              pageSelectDirection,
                                              typeId,
                                              ignoreSecurity)

        End Function

#Region "Internal helper methods"

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="table"></param>
        ''' <param name="sortBy"></param>
        ''' <param name="filter"></param>
        ''' <param name="pageSize"></param>
        ''' <param name="forward"></param>
        ''' <param name="startRecord"></param>
        ''' <param name="includeReferencedRecord"></param>
        ''' <param name="selectType">'selectType can be enumSelectType.PAGE or enumSelectType.COUNT_TO_REC</param>
        ''' <param name="typeID"></param>
        ''' <param name="ignoreSecurity"></param>
        ''' <param name="selectFields"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function SelectInternal(ByVal table As clsTable,
                                        ByVal sortBy As clsSortCollection,
                                        ByVal filter As clsSearchFilter,
                                        ByVal pageSize As Integer,
                                        ByVal forward As Boolean,
                                        ByVal startRecord As clsMaskFieldDictionary,
                                        ByVal includeReferencedRecord As Boolean,
                                        ByVal selectType As clsSelectInfo.enumSelectType,
                                        ByVal typeID As Integer,
                                        ByVal ignoreSecurity As Boolean,
                                        Optional ByVal selectFields() As String = Nothing) As DataTable

            Dim selectInfo = New clsSelectInfo(table,
                                               sortBy,
                                               filter,
                                               pageSize,
                                               forward,
                                               startRecord,
                                               includeReferencedRecord,
                                               selectType,
                                               typeID,
                                               ignoreSecurity,
                                               selectFields)

            PrintSql(selectInfo.SQL)

            Return selectInfo.DataTable

        End Function

        Private Sub PrintSql(ByVal sql As String)
            Debug.WriteLine(sql)
        End Sub

#End Region

    End Class

End Namespace