Public Class clsFullTextInfo

#Region " Members "

    Private m_objDB As clsDB_Direct
    Private m_strTable As String

    Private m_blnBackgroundUpdateOn As Boolean
    Private m_intPendingChanges As Integer
    Private m_ePopulateStatus As enumPopulateStatus
    Private m_strPopulateStatus As String
    Private m_blnHasActiveIndex As Boolean
    Private m_intCatalogId As Integer
    Private m_blnChangeTrackingOn As Boolean
    Private m_intDocsProcessed As Integer
    Private m_intFailCount As Integer
    Private m_intItemCount As Integer
    Private m_dtPopulateStart As DateTime
    Private m_dtPopulateEnd As DateTime

#End Region

#Region " Enumerators "

    Public Enum enumPopulateStatus As Integer
        Ilde = 0
        FullPopulation
        IncrementalPopulation
        PropagationOfTrackedChanges
        BackgroundUpdateIndex
        ThrottledOrPaused
    End Enum

#End Region

#Region " Constructors "

    ''' <summary>
    ''' Constructor for an instance of FullText Information Object
    ''' </summary>
    ''' <param name="objDB">Database that the FullText index is in.</param>
    ''' <param name="strTableName">Table or View name that the FullText index belongs to.</param>
    Public Sub New(ByVal objDB As clsDB_Direct, ByVal strTableName As String)
        m_objDB = objDB
        m_strTable = strTableName
        Refresh()
    End Sub

#End Region

#Region " Properties "

    ''' <summary>
    ''' The table has full-text background update index (autochange tracking) enabled.
    ''' </summary>
    ''' <remarks>
    ''' When
    ''' BackgroundUpdateOn and ChangeTrackingOn are True = Automatic
    ''' BackgroundUpdateOn is False and ChangeTrackingOn is True = Manual
    ''' ChangeTrackingOn is False = Does not track changes
    ''' </remarks>
    Public ReadOnly Property BackgroundUpdateOn() As Boolean
        Get
            Return m_blnBackgroundUpdateOn
        End Get
    End Property

    ''' <summary>
    ''' Table has an active full-text index.
    ''' </summary>
    Public ReadOnly Property HasActiveIndex() As Boolean
        Get
            Return m_blnHasActiveIndex
        End Get
    End Property

    ''' <summary>
    ''' Table has full-text change-tracking enabled.
    ''' </summary>
    ''' <remarks>
    ''' If this is false most properties will not be populated.
    ''' </remarks>
    Public ReadOnly Property ChangeTrackingOn() As Boolean
        Get
            Return m_blnChangeTrackingOn
        End Get
    End Property

    ''' <summary>
    ''' Number of rows processed since the start of full-text indexing.
    ''' </summary>
    Public ReadOnly Property DocsProcessed() As Integer
        Get
            Return m_intDocsProcessed
        End Get
    End Property

    ''' <summary>
    ''' ID of the full-text catalog in which the full-text index data for the table resides.
    ''' </summary>
    Public ReadOnly Property CatalogId() As Integer
        Get
            Return m_intCatalogId
        End Get
    End Property

    ''' <summary>
    ''' Number of rows that are waiting to be indexed.
    ''' </summary>
    Public ReadOnly Property PendingChanges() As Integer
        Get
            Return m_intPendingChanges
        End Get
    End Property

    ''' <summary>
    ''' The number of rows that full-text search did not index. 
    ''' </summary>
    Public ReadOnly Property FailCount() As Integer
        Get
            Return m_intFailCount
        End Get
    End Property

    ''' <summary>
    ''' Number of rows that were full-text indexed successfully.
    ''' </summary>
    Public ReadOnly Property IndexCount() As Integer
        Get
            Return m_intItemCount
        End Get
    End Property

    ''' <summary>
    ''' The status of the index population for the table.
    ''' </summary>
    Public ReadOnly Property PopulateStatus() As enumPopulateStatus
        Get
            Return m_ePopulateStatus
        End Get
    End Property

    ''' <summary>
    ''' The status of the index population for the table in a human readable version.
    ''' </summary>
    Public ReadOnly Property PopulateStatusText() As String
        Get
            Return m_strPopulateStatus
        End Get
    End Property

    ''' <summary>
    ''' The date and time the last population was started.
    ''' </summary>
    Public ReadOnly Property PopulateStart() As DateTime
        Get
            Return m_dtPopulateStart
        End Get
    End Property

    ''' <summary>
    ''' The date and time the last population was completed.
    ''' </summary>
    Public ReadOnly Property PopulateEnd() As DateTime
        Get
            Return m_dtPopulateEnd
        End Get
    End Property

#End Region

#Region " Methods "

#Region " Refresh "

    ''' <summary>
    ''' Refreshes the information for the Full-Text index.
    ''' </summary>
    Public Sub Refresh()
        Dim strSQL As String = "SELECT OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFullTextBackgroundUpdateIndexOn') as BackgroundUpdateIndexOn," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextPendingChanges') as PendingChanges," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextPopulateStatus') as PopulateStatus," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableHasActiveFulltextIndex') as HasActiveFulltextIndex," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextCatalogId') as CatalogId," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFullTextChangeTrackingOn') as ChangeTrackingOn," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextDocsProcessed') as DocsProcessed," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextFailCount') as FailCount," & vbCrLf & _
                                    "OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextItemCount') as ItemCount," & vbCrLf & _
                                    "crawl_start_date as PopulationStartDate," & vbCrLf & _
                                    "crawl_end_date as PopulationEndDate" & vbCrLf & _
                               "FROM sys.fulltext_indexes WHERE" & vbCrLf & _
                                    "fulltext_catalog_id = OBJECTPROPERTYEX(OBJECT_ID(N'{0}'), N'TableFulltextCatalogId')" & vbCrLf & _
                                    "AND object_id = OBJECT_ID(N'{0}')"

        Dim objDT As DataTable = m_objDB.GetDataTableBySQL(String.Format(strSQL, m_strTable))

        If objDT IsNot Nothing And objDT.Rows.Count > 0 Then
            m_blnBackgroundUpdateOn = CBool(clsDB.DataRowValue(objDT.Rows(0), "BackgroundUpdateIndexOn", False))
            m_intPendingChanges = CInt(clsDB.DataRowValue(objDT.Rows(0), "PendingChanges", clsDBConstants.cintNULL))
            m_ePopulateStatus = CType(clsDB.DataRowValue(objDT.Rows(0), "PopulateStatus", False), enumPopulateStatus)
            m_strPopulateStatus = PopulateStatusToString(m_ePopulateStatus)
            m_blnHasActiveIndex = CBool(clsDB.DataRowValue(objDT.Rows(0), "HasActiveFulltextIndex", False))
            m_intCatalogId = CInt(clsDB.DataRowValue(objDT.Rows(0), "CatalogId", 0))
            m_blnChangeTrackingOn = CBool(clsDB.DataRowValue(objDT.Rows(0), "ChangeTrackingOn", False))
            m_intDocsProcessed = CInt(clsDB.DataRowValue(objDT.Rows(0), "DocsProcessed", clsDBConstants.cintNULL))
            m_intFailCount = CInt(clsDB.DataRowValue(objDT.Rows(0), "FailCount", clsDBConstants.cintNULL))
            m_intItemCount = CInt(clsDB.DataRowValue(objDT.Rows(0), "ItemCount", clsDBConstants.cintNULL))
            m_dtPopulateStart = CDate(clsDB.DataRowValue(objDT.Rows(0), "PopulationStartDate", DateTime.MinValue))
            m_dtPopulateEnd = CDate(clsDB.DataRowValue(objDT.Rows(0), "PopulationEndDate", m_dtPopulateStart))
        Else
            m_blnBackgroundUpdateOn = False
            m_intPendingChanges = clsDBConstants.cintNULL
            m_ePopulateStatus = clsFullTextInfo.enumPopulateStatus.Ilde
            m_strPopulateStatus = PopulateStatusToString(m_ePopulateStatus)
            m_blnHasActiveIndex = False
            m_intCatalogId = 0
            m_blnChangeTrackingOn = False
            m_intDocsProcessed = clsDBConstants.cintNULL
            m_intFailCount = clsDBConstants.cintNULL
            m_intItemCount = clsDBConstants.cintNULL
            m_dtPopulateStart = DateTime.MinValue
            m_dtPopulateEnd = m_dtPopulateStart
        End If
    End Sub

#End Region

#Region " Populate Status To String "

    ''' <summary>
    ''' Converts the Populate Status value into a human readable string.
    ''' </summary>
    ''' <param name="eValue">Populate Status in the database.</param>
    ''' <returns>Human readable string of the populateStatus value.</returns>
    ''' <remarks></remarks>
    Private Function PopulateStatusToString(ByVal eValue As enumPopulateStatus) As String
        Select Case eValue
            Case enumPopulateStatus.Ilde
                Return "Idle."

            Case enumPopulateStatus.FullPopulation
                Return "Full population is in progress."

            Case enumPopulateStatus.IncrementalPopulation
                Return "Incremental population is in progress."

            Case enumPopulateStatus.PropagationOfTrackedChanges
                Return "Propagation of tracked changes is in progress."

            Case enumPopulateStatus.BackgroundUpdateIndex
                Return "Background update index is in progress, such as autochange tracking."

            Case enumPopulateStatus.ThrottledOrPaused
                Return "Full-text indexing is throttled or paused."

            Case Else
                Return "Status is unknown."
        End Select
    End Function

#End Region

#End Region

End Class
