#Region " File Information "

'======================================================================
' This is a class which holds commonly used K1 database objects in memory
'======================================================================

#Region " Revision History "

'======================================================================
' Name      Date        Description
'----------------------------------------------------------------------
' KSD       16/01/2007  Implemented.
'======================================================================

#End Region

#End Region

Public Class clsSysInfo
    Implements IDisposable

#Region " Members "

    Private m_objDB As clsDB
    Private m_colTables As FrameworkCollections.K1DualKeyDictionary(Of clsTable, Integer)
    Private m_colFields As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
    Private m_colFieldLinks As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
    Private m_objK1Configuration As clsK1Configuration
    Private m_colMethods As FrameworkCollections.K1DualKeyDictionary(Of clsMethod, Integer)
    Private m_colAppMethods As FrameworkCollections.K1DualKeyDictionary(Of clsApplicationMethod, Integer)
    Private m_colDRMMethods As FrameworkCollections.K1Dictionary(Of clsDRMMethod)
    Private m_colDRMFunctions As FrameworkCollections.K1Dictionary(Of clsDRMFunction)



    Private m_colLanguagesLoaded As FrameworkCollections.K1Dictionary(Of String)
    Private m_colCaptionStrings As FrameworkCollections.K1Dictionary(Of String)
    Private m_colMethodStrings As FrameworkCollections.K1Dictionary(Of String)
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
    Private m_colSecGroups As FrameworkCollections.K1Dictionary(Of clsSecurityGroup)
    Private m_colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn)
    Private m_colFilters As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
    Private m_colNoiseWords As FrameworkCollections.K1Collection(Of String)
    Private m_objK1Groups As clsK1Groups
    Private m_objTimeDifference As Double = 0
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB)
        m_objDB = objDB

        m_objDB.LastRefresh = m_objDB.GetCurrentTime()

        m_colTables = clsTable.GetList(m_objDB)
        m_colFields = clsField.GetList2(m_objDB)
        m_colFieldLinks = clsFieldLink.GetList2(m_objDB)
        m_colMethods = clsMethod.GetList(m_objDB)
        m_objK1Configuration = clsK1Configuration.GetDefault(m_objDB)
        m_colSecGroups = clsSecurityGroup.GetList(m_objDB)
        m_colListColumns = clsListColumn.GetList(m_objDB)

        LoadFilters()

        If m_colLanguagesLoaded IsNot Nothing Then
            m_colLanguagesLoaded.Clear()
            m_colLanguagesLoaded = Nothing
        End If

        If m_colCaptionStrings IsNot Nothing Then
            m_colCaptionStrings.Clear()
            m_colCaptionStrings = Nothing
        End If

        If m_colMethodStrings IsNot Nothing Then
            m_colMethodStrings.Clear()
            m_colMethodStrings = Nothing
        End If

        LoadStrings(m_objK1Configuration.DefaultProfile.LanguageID)

        If m_objK1Configuration.DbVersion >= 11 Then
            m_colAppMethods = clsApplicationMethod.GetList(m_objDB, m_objK1Configuration)

        End If

        'New DRMFunctions and methods introduced in 11.10
        If m_objK1Configuration.DbVersion >= 11.1 Then
            m_colDRMFunctions = clsDRMFunction.GetList(m_objDB)
            m_colDRMMethods = clsDRMMethod.GetList(m_objDB)
        End If

        For Each objField As clsField In m_colFields.Values
            Dim objTable As clsTable = m_colTables(objField.TableID)

            objTable.Fields.Add(objField.KeyID, objField)
        Next

        Dim colTFs As FrameworkCollections.K1Dictionary(Of clsTypeField)
        Dim colTFLs As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)

        colTFs = clsTypeField.GetList(m_objDB)
        colTFLs = clsTypeFieldLink.GetList(m_objDB)

        For Each objTF As clsTypeField In colTFs.Values
            Dim objField As clsField = m_colFields(objTF.FieldID)

            objField.TypeFieldInfos.Add(CStr(objTF.AppliesToTypeID), objTF)
        Next

        For Each objTFL As clsTypeFieldLink In colTFLs.Values
            Dim objFL As clsFieldLink = m_colFieldLinks(objTFL.FieldLinkID)

            objFL.TypeFieldLinkInfos.Add(CStr(objTFL.AppliesToTypeID), objTFL)
        Next

        AssignRelatedLinks()
        AssignListColumns()

        'Assign Icons to tables (saves processing on load)
        Dim colIcons As FrameworkCollections.K1Dictionary(Of clsIcon)
        colIcons = clsIcon.GetList(m_objDB)

        For Each objTable As clsTable In m_colTables.Values
            If Not objTable.IconID = clsDBConstants.cintNULL Then
                objTable.Icon = colIcons(CStr(objTable.IconID))
            End If
        Next

        '2016-10-06 -- Peter & James -- Bug fix for #1600003221
        If objDB.Profile IsNot Nothing Then
            objDB.Profile.ResetProfile()
        End If

        m_objTimeDifference = m_objDB.GetServerTimeDifference()
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' All the standard applicaiton methods indexed by UIID (clsApplicationMethod.enumAppType as string) or ID (integer)
    ''' </summary>
    Public ReadOnly Property ApplicationMethods() As FrameworkCollections.K1DualKeyDictionary(Of clsApplicationMethod, Integer)
        Get
            Return m_colAppMethods
        End Get
    End Property
    ''' <summary>
    ''' All the DRM methods indexed by ID
    ''' </summary>
    Public ReadOnly Property DRMMethods() As FrameworkCollections.K1Dictionary(Of clsDRMMethod)
        Get
            Return m_colDRMMethods
        End Get
    End Property
    ''' <summary>
    ''' All the DRM Functions indexed by ID
    ''' </summary>
    Public ReadOnly Property DRMFunctions() As FrameworkCollections.K1Dictionary(Of clsDRMFunction)
        Get
            Return m_colDRMFunctions
        End Get
    End Property

    ''' <summary>
    ''' All the K1 Tables indexed by DatabaseName and Table ID (Those tables which are in the "Table" table)
    ''' </summary>
    Public ReadOnly Property Tables() As FrameworkCollections.K1DualKeyDictionary(Of clsTable, Integer)
        Get
            Return m_colTables
        End Get
    End Property

    ''' <summary>
    ''' All the K1 Fields indexed by [Table ID]_[Field DatabaseName] and Field ID
    ''' </summary>
    Public ReadOnly Property Fields() As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
        Get
            Return m_colFields
        End Get
    End Property

    ''' <summary>
    ''' All the K1 Field Links indexed by FieldLink ID
    ''' </summary>
    Public ReadOnly Property FieldLinks() As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
        Get
            Return m_colFieldLinks
        End Get
    End Property

    ''' <summary>
    ''' The K1 Configuration Object
    ''' </summary>
    Public ReadOnly Property K1Configuration() As clsK1Configuration
        Get
            Return m_objK1Configuration
        End Get
    End Property

    ''' <summary>
    ''' All the K1 methods indexed by UIID (Use clsMethod.enumMethods for dictionary key)
    ''' </summary>
    Public ReadOnly Property Methods() As FrameworkCollections.K1DualKeyDictionary(Of clsMethod, Integer)
        Get
            Return m_colMethods
        End Get
    End Property

    ''' <summary>
    ''' All the Security Groups indexed by ID
    ''' </summary>
    Public ReadOnly Property SecurityGroups() As FrameworkCollections.K1Dictionary(Of clsSecurityGroup)
        Get
            Return m_colSecGroups
        End Get
    End Property

    Public ReadOnly Property ListColumns() As FrameworkCollections.K1Dictionary(Of clsListColumn)
        Get
            Return m_colListColumns
        End Get
    End Property

    Public ReadOnly Property Filters() As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
        Get
            Return m_colFilters
        End Get
    End Property

    Public ReadOnly Property NoiseWords() As FrameworkCollections.K1Collection(Of String)
        Get
            If m_colNoiseWords Is Nothing Then
                LoadNoiseWords()
            End If

            Return m_colNoiseWords
        End Get
    End Property

    Public ReadOnly Property K1Groups() As clsK1Groups
        Get
            If m_objK1Groups Is Nothing Then
                m_objK1Groups = New clsK1Groups(m_objDB)
            End If
            Return m_objK1Groups
        End Get
    End Property

    '2017-08-24 -- Peter Melisi -- Changes for Timezones for User Profiles
    Public Property TimeDifference() As Double
        Get
            Return m_objTimeDifference
        End Get
        Set(value As Double)
            m_objTimeDifference = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Memory Object Creation "

    ''' <summary>
    ''' Updates the field link collections when creating a new field link
    ''' </summary>
    Private Sub AssignFieldLink(ByVal objFKeyTable As clsTable, ByVal objFieldLink As clsFieldLink, _
    Optional ByVal objReloadTable As clsTable = Nothing)
        If objFKeyTable IsNot Nothing Then
            UpdateSystemFieldLinks(objFieldLink, objFKeyTable)
        End If

        If objReloadTable Is Nothing Then
            objFKeyTable.FieldLinks(CStr(objFieldLink.ForeignKeyFieldID)) = objFieldLink
        End If
    End Sub

    ''' <summary>
    ''' Updates the field link collections when creating a new field link
    ''' </summary>
    Private Sub UpdateSystemFieldLinks(ByVal objFieldLink As clsFieldLink, ByVal objFKeyTable As clsTable)
        Dim objIDField As clsField = m_colFields(objFieldLink.IdentityFieldID)
        Dim objIDTable As clsTable = m_colTables(objIDField.TableID)

        If objIDTable IsNot Nothing Then
            If objFKeyTable.IsLinkTable Then
                If objIDTable.ManyToManyLinks(objFieldLink.KeyID) Is Nothing Then
                    objIDTable.ManyToManyLinks.Add(objFieldLink.KeyID, objFieldLink)
                Else
                    objIDTable.ManyToManyLinks(objFieldLink.KeyID) = objFieldLink
                End If
            Else
                If objIDTable.OneToManyLinks(objFieldLink.KeyID) Is Nothing Then
                    objIDTable.OneToManyLinks.Add(objFieldLink.KeyID, objFieldLink)
                Else
                    objIDTable.OneToManyLinks(objFieldLink.KeyID) = objFieldLink
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Updates the field collections when creating a new field
    ''' </summary>
    Private Sub UpdateSystemFields(ByVal objField As clsField)
        m_colFields(objField.KeyName) = objField
    End Sub

    Public Sub LoadStrings(ByVal intLanguageID As Integer)

        If m_colLanguagesLoaded Is Nothing Then
            m_colLanguagesLoaded = New FrameworkCollections.K1Dictionary(Of String)
        End If

        If m_colCaptionStrings Is Nothing Then
            m_colCaptionStrings = New FrameworkCollections.K1Dictionary(Of String)
        End If

        If m_colMethodStrings Is Nothing Then
            m_colMethodStrings = New FrameworkCollections.K1Dictionary(Of String)
        End If

        If m_colLanguagesLoaded(CStr(intLanguageID)) Is Nothing Then
            Dim strSQL As String = "SELECT" & vbCrLf & _
                "[" & clsDBConstants.Tables.cCAPTION & "].[" & clsDBConstants.Fields.cID & "]," & vbCrLf & _
                "CONVERT(NVARCHAR(4000), [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cSTRING & "]) AS [String]" & vbCrLf & _
                "FROM [" & clsDBConstants.Tables.cCAPTION & "]" & vbCrLf & _
                "INNER JOIN [" & clsDBConstants.Tables.cSTRING & "] ON [" & clsDBConstants.Tables.cCAPTION & "].[" & clsDBConstants.Fields.Caption.cSTRINGID & "] = [" & clsDBConstants.Tables.cSTRING & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & _
                "INNER JOIN [" & clsDBConstants.Tables.cLANGUAGESTRING & "] ON [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = [" & clsDBConstants.Tables.cSTRING & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & _
                "WHERE [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cLANGUAGEID & "] = " & intLanguageID

            Dim objDT As DataTable = m_objDB.GetDataTableBySQL(strSQL)
            For Each objRow As DataRow In objDT.Rows
                Dim strKey As String = intLanguageID & "_" & CStr(objRow("ID"))
                If m_colCaptionStrings(strKey) Is Nothing Then
                    m_colCaptionStrings.Add(strKey, CStr(clsDB_Direct.NullValue(objRow("String"), "")))
                End If
            Next

            strSQL = "SELECT" & vbCrLf & _
                "[" & clsDBConstants.Tables.cMETHOD & "].[" & clsDBConstants.Fields.cID & "]," & vbCrLf & _
                "CONVERT(NVARCHAR(4000), [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cSTRING & "]) AS [String]" & vbCrLf & _
                "FROM [" & clsDBConstants.Tables.cMETHOD & "]" & vbCrLf & _
                "INNER JOIN [" & clsDBConstants.Tables.cBUTTON & "] ON [" & clsDBConstants.Tables.cMETHOD & "].[" & clsDBConstants.Fields.Method.cBUTTONID & "] = [" & clsDBConstants.Tables.cBUTTON & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & _
                "INNER JOIN [" & clsDBConstants.Tables.cSTRING & "] ON [" & clsDBConstants.Tables.cBUTTON & "].[" & clsDBConstants.Fields.Button.cSTRINGID & "] = [" & clsDBConstants.Tables.cSTRING & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & _
                "INNER JOIN [" & clsDBConstants.Tables.cLANGUAGESTRING & "] ON [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = [" & clsDBConstants.Tables.cSTRING & "].[" & clsDBConstants.Fields.cID & "]" & vbCrLf & _
                "WHERE [" & clsDBConstants.Tables.cLANGUAGESTRING & "].[" & clsDBConstants.Fields.LanguageString.cLANGUAGEID & "] = " & intLanguageID

            objDT = m_objDB.GetDataTableBySQL(strSQL)
            For Each objRow As DataRow In objDT.Rows
                Dim strKey As String = intLanguageID & "_" & CStr(objRow("ID"))
                If m_colMethodStrings(strKey) Is Nothing Then
                    m_colMethodStrings.Add(strKey, CStr(clsDB_Direct.NullValue(objRow("String"), "")))
                End If
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            m_colLanguagesLoaded.Add(CStr(intLanguageID), CStr(intLanguageID))
        End If
    End Sub

    Public Function GetCaptionString(ByVal intCaptionID As Integer) As String
        If intCaptionID = clsDBConstants.cintNULL Then
            Return Nothing
        End If

        If m_colLanguagesLoaded(CStr(m_objDB.Profile.LanguageID)) Is Nothing Then
            LoadStrings(m_objDB.Profile.LanguageID)
        End If

        If m_colLanguagesLoaded(CStr(m_objDB.Profile.DefaultLanguageID)) Is Nothing Then
            LoadStrings(m_objDB.Profile.DefaultLanguageID)
        End If

        Dim strText As String
        If m_colCaptionStrings.ContainsKey(m_objDB.Profile.LanguageID & "_" & intCaptionID) Then
            strText = m_colCaptionStrings(m_objDB.Profile.LanguageID & "_" & intCaptionID)
            Return strText
        End If

        If m_colCaptionStrings.ContainsKey(m_objDB.Profile.DefaultLanguageID & "_" & intCaptionID) Then
            strText = m_colCaptionStrings(m_objDB.Profile.DefaultLanguageID & "_" & intCaptionID)
            Return strText
        End If

        Dim objCaption As clsCaption = clsCaption.GetItem(intCaptionID, m_objDB)
        If objCaption Is Nothing Then
            strText = Nothing
        Else
            strText = objCaption.GetString(m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)
        End If

        m_colCaptionStrings.Add(m_objDB.Profile.LanguageID & "_" & intCaptionID, strText)

        Return strText
    End Function

    Public Function GetMethodString(ByVal intMethodID As Integer) As String
        If m_colLanguagesLoaded(CStr(m_objDB.Profile.LanguageID)) Is Nothing Then
            LoadStrings(m_objDB.Profile.LanguageID)
        End If

        If m_colLanguagesLoaded(CStr(m_objDB.Profile.DefaultLanguageID)) Is Nothing Then
            LoadStrings(m_objDB.Profile.DefaultLanguageID)
        End If

        Dim strText As String = m_colMethodStrings(m_objDB.Profile.LanguageID & "_" & intMethodID)
        If strText Is Nothing Then
            strText = m_colMethodStrings(m_objDB.Profile.DefaultLanguageID & "_" & intMethodID)
        End If

        Return strText
    End Function

    Public Sub ConvertField_DateTimeToLocal(ByVal objMask As clsMaskBase, Optional ByVal blnToServer As Boolean = False)
        Try
            If objMask.GetType Is GetType(clsMaskFieldLink) OrElse TimeDifference = 0 Then
                Return
            End If

            Dim objMaskField As clsMaskField = CType(objMask, clsMaskField)

            If (objMaskField.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME OrElse _
            objMaskField.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY OrElse _
            objMaskField.Field.DateType = clsDBConstants.enumDateTypes.TIME_ONLY) Then 'localdatetime
                Return
            End If

            If objMaskField.Value1 IsNot Nothing AndAlso _
            objMaskField.Value1.Value IsNot Nothing AndAlso _
            objMaskField.Value1.Value.GetType Is GetType(Date) Then
                objMaskField.Value1.Value = ToLocalTime(CDate(objMaskField.Value1.Value), blnToServer) 'CType(objMaskField.Value1.Value, Date).AddHours(dblHours) 
            End If

            If objMaskField.Value2 IsNot Nothing AndAlso _
            objMaskField.Value2.Value IsNot Nothing AndAlso _
            objMaskField.Value2.Value.GetType Is GetType(Date) Then
                objMaskField.Value2.Value = ToLocalTime(CDate(objMaskField.Value2.Value), blnToServer) 'CType(objMaskField.Value2.Value, Date).AddHours(dblHours) 
            End If

            Return
        Catch ex As Exception
            'log error?
            Return
        End Try
    End Sub

    Public Function ToLocalTime(ByVal objDatetime As DateTime, Optional ByVal blnToServerTime As Boolean = False) As DateTime 'localdatetime
        Try
            If TimeDifference = 0 OrElse objDatetime = Nothing Then
                Return objDatetime
            End If

            Dim dblHours As Double = Math.Abs(TimeDifference)

            If (TimeDifference > 0 AndAlso Not blnToServerTime) OrElse _
            (TimeDifference < 0 AndAlso blnToServerTime) Then
                dblHours = -dblHours
            End If

            '2018-04-03 -- Peter Melisi -- Bug fix for Incorrect Date Timezone Bug

            '2017-08-24 -- Peter Melisi -- Changes for Timezones for User Profiles
            'If (objDatetime.TimeOfDay.TotalSeconds = 0) Then
            'objDatetime = m_objDB.GetCurrentTime.AddHours(dblHours).Date
            'Return objDatetime
            'Else
            Return objDatetime.AddHours(dblHours)
            'End If
        Catch ex As Exception
            'log error?
            Return Nothing
        End Try
    End Function

    Public Function GetCurrentTimeByType(ByRef objTableMask As clsTableMask, ByRef strFieldName As String) As DateTime 'localdatetime
        Try

            If (objTableMask Is Nothing) Then
                Return m_objDB.GetCurrentTime()
            End If

            Dim objField As clsMaskField = objTableMask.MaskFieldCollection(strFieldName)
            If objField IsNot Nothing AndAlso
                (objField.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME OrElse
                 objField.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY OrElse
                 objField.Field.DateType = clsDBConstants.enumDateTypes.TIME_ONLY) Then
                Return ToLocalTime(m_objDB.GetCurrentTime) 'localdatetime
            Else
                Return m_objDB.GetCurrentTime()
            End If
        Catch ex As Exception
            'log error?
            Return m_objDB.GetCurrentTime()
        End Try
    End Function

    Private Sub AssignListColumns(Optional ByVal objReloadTable As clsTable = Nothing)
        For Each objListCol As clsListColumn In m_colListColumns.Values
            Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn)
            Dim objField As clsField = m_colFields(objListCol.FieldID)

            If objReloadTable Is Nothing OrElse objReloadTable.ID = objField.TableID Then
                Dim objTable As clsTable = m_colTables(objField.TableID)

                If objListCol.AppliesToTypeID = clsDBConstants.cintNULL Then
                    colListColumns = objTable.ListColumns
                Else
                    colListColumns = objTable.TypeListColumns(CStr(objListCol.AppliesToTypeID))

                    If colListColumns Is Nothing Then
                        colListColumns = New FrameworkCollections.K1Dictionary(Of clsListColumn)
                        objTable.TypeListColumns.Add(CStr(objListCol.AppliesToTypeID), colListColumns)
                    End If
                End If

                colListColumns.Add(objField.TableID & "_" & objListCol.FieldID, objListCol)
            End If
        Next
    End Sub

    Private Sub AssignRelatedLinks(Optional ByVal objReloadTable As clsTable = Nothing)
        For Each objFieldLink As clsFieldLink In m_colFieldLinks.Values
            Dim objFKeyField As clsField = m_colFields(objFieldLink.ForeignKeyFieldID)
            Dim objFKeyTable As clsTable = m_colTables(objFKeyField.TableID)

            If objReloadTable Is Nothing OrElse objReloadTable.ID = objFieldLink.IdentityTable.ID Then
                AssignFieldLink(objFKeyTable, objFieldLink, objReloadTable)
            End If
        Next
    End Sub

    Private Sub LoadFilters()
        Dim colBaseFilters As List(Of clsFilter)

        colBaseFilters = clsFilter.GetList(m_objDB)

        If m_colFilters IsNot Nothing Then
            m_colFilters.Dispose()
            m_colFilters = Nothing
        End If

        m_colFilters = New FrameworkCollections.K1Dictionary(Of List(Of clsFilter))

        Dim colFilters As List(Of clsFilter)

        For Each objFilter As clsFilter In colBaseFilters
            Dim strKey As String = objFilter.KeyName

            colFilters = m_colFilters(strKey)
            If colFilters Is Nothing Then
                colFilters = New List(Of clsFilter)
                m_colFilters.Add(strKey, colFilters)
            End If

            colFilters.Add(objFilter)

            'Parent Fields lists
            If objFilter.FilterType = clsDBConstants.enumFilterTypes.PARENT_FIELD Then
                Dim objParentField As clsField = m_colFields(objFilter.FilterValueFieldID)
                Dim intTypeID As Integer = objFilter.AppliesToTypeID

                Dim colParentFilters As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
                If objParentField.ParentOfFilters Is Nothing Then
                    colParentFilters = New FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
                    objParentField.ParentOfFilters = colParentFilters
                Else
                    colParentFilters = objParentField.ParentOfFilters
                End If

                colFilters = colParentFilters(CStr(intTypeID))
                If colFilters Is Nothing Then
                    colFilters = New List(Of clsFilter)
                    colParentFilters.Add(CStr(intTypeID), colFilters)
                End If

                colFilters.Add(objFilter)
            End If
        Next
    End Sub

    Private Sub LoadNoiseWords()
        Dim strSQL As String = "SELECT * FROM NoiseWords"
        Dim dtWords As DataTable = m_objDB.GetDataTableBySQL(strSQL)
        Dim colNoiseWords As New FrameworkCollections.K1Collection(Of String)

        If dtWords IsNot Nothing AndAlso dtWords.Rows.Count > 0 Then
            For Each objRow As DataRow In dtWords.Rows
                Dim strWord As String = CStr(objRow(0))
                If Not colNoiseWords.Contains(CStr(objRow(0))) Then
                    colNoiseWords.Add(strWord.ToLower)
                End If
            Next
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If dtWords IsNot Nothing Then
            dtWords.Dispose()
            dtWords = Nothing
        End If

        m_colNoiseWords = colNoiseWords
    End Sub
#End Region

#Region " DRM Updates "

#Region " Inserts, Updates "

    ''' <summary>
    ''' Updates the field collections when creating a new field
    ''' </summary>
    Public Sub DRMInsertUpdateField(ByVal objTable As clsTable, ByVal objField As clsField)
        objTable.Fields(objField.KeyID) = objField

        UpdateSystemFields(objField)
    End Sub

    ''' <summary>
    ''' Updates the field link collections when creating a new field link
    ''' </summary>
    Public Sub DRMInsertUpdateFieldLink(ByVal objTable As clsTable, ByVal objFieldLink As clsFieldLink)
        m_colFieldLinks(CStr(objFieldLink.ForeignKeyFieldID)) = objFieldLink
        AssignFieldLink(objTable, objFieldLink)
    End Sub

    ''' <summary>
    ''' Updates the table collections when creating a new table
    ''' </summary>
    Public Sub DRMInsertUpdateTable(ByVal objTable As clsTable)
        Dim blnNewTable As Boolean = False

        If m_colTables(objTable.ID) Is Nothing Then
            blnNewTable = True
        End If

        m_colTables(objTable.DatabaseName) = objTable

        If blnNewTable Then
            For Each objField As clsField In objTable.Fields.Values
                UpdateSystemFields(objField)

                If Not objField.FieldLink Is Nothing Then
                    UpdateSystemFieldLinks(objField.FieldLink, objTable)
                End If
            Next
        Else
            AssignRelatedLinks(objTable)
        End If

        If Not blnNewTable Then
            AssignListColumns(objTable)
        End If
    End Sub

    Public Sub DRMUpdateK1Config()
        m_objK1Configuration = clsK1Configuration.GetDefault(m_objDB)
    End Sub

    Public Sub DRMInsertUpdateListColumn(ByVal objTable As clsTable, ByVal objListColumn As clsListColumn)
        If m_colListColumns(CStr(objListColumn.ID)) Is Nothing Then
            m_colListColumns.Add(CStr(objListColumn.ID), objListColumn)
        Else
            m_colListColumns(CStr(objListColumn.ID)) = objListColumn
        End If

        Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = Nothing

        If Not objListColumn.AppliesToTypeID = clsDBConstants.cintNULL Then
            colListColumns = objListColumn.Field.Table.TypeListColumns(CStr(objListColumn.AppliesToTypeID))
            If colListColumns Is Nothing Then
                colListColumns = New FrameworkCollections.K1Dictionary(Of clsListColumn)
                objListColumn.Field.Table.TypeListColumns.Add(CStr(objListColumn.AppliesToTypeID), colListColumns)
            End If
        End If

        If colListColumns Is Nothing Then
            colListColumns = objTable.ListColumns
        End If

        If colListColumns(objListColumn.Field.TableID & "_" & objListColumn.FieldID) Is Nothing Then
            colListColumns.Add(objListColumn.Field.TableID & "_" & objListColumn.FieldID, objListColumn)
        Else
            colListColumns(objListColumn.Field.TableID & "_" & objListColumn.FieldID) = objListColumn
        End If
    End Sub

    Public Sub DRMInsertUpdateFilter(ByVal objFilter As clsFilter)
        Dim colFilters As List(Of clsFilter)

        colFilters = clsFilter.GetList(objFilter, m_objDB)

        Dim strKey As String = objFilter.KeyName

        If m_objDB.SysInfo.Filters(strKey) Is Nothing Then
            m_objDB.SysInfo.Filters.Add(strKey, colFilters)
        Else
            m_objDB.SysInfo.Filters(strKey) = colFilters
        End If
    End Sub

    Public Sub DRMInsertUpdateCaption(ByVal objCaption As clsCaption)
        Dim strText As String = objCaption.StringObj.GetLanguageString(m_objDB.Profile.LanguageID, _
                m_objDB.Profile.DefaultLanguageID)

        If m_colCaptionStrings(m_objDB.Profile.LanguageID & "_" & objCaption.ID) Is Nothing Then
            m_colCaptionStrings.Add(m_objDB.Profile.LanguageID & "_" & objCaption.ID, strText)
        Else
            m_colCaptionStrings(m_objDB.Profile.LanguageID & "_" & objCaption.ID) = strText
        End If
    End Sub

    Public Sub DRMRefreshK1Groups()
        If m_objK1Groups IsNot Nothing Then
            m_objK1Groups.Dispose()
            m_objK1Groups = Nothing
        End If

        m_objK1Groups = New clsK1Groups(m_objDB)
    End Sub
#End Region

#Region " Deletes "

    ''' <summary>
    ''' Updates the field collections when deleting a field
    ''' </summary>
    Public Sub DRMDeleteField(ByVal objTable As clsTable, ByVal objField As clsField)
        objTable.Fields.Remove(objField.KeyID)
        m_colFields.Remove(objField.KeyName)

        Dim strFilterKey1 As String

        strFilterKey1 = "F_" & objField.ID

        Dim colKeys As New List(Of String)
        For Each strKey As String In m_colFilters.Keys
            If strKey.StartsWith(strFilterKey1) Then
                colKeys.Add(strKey)
            End If
        Next

        For Each strKey As String In colKeys
            m_colFilters.Remove(strKey)
        Next

        Dim strLCKey As String = objTable.ID & "_" & objField.ID
        Dim objLC As clsListColumn

        Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = objTable.ListColumns
        objLC = colListColumns(strLCKey)
        If objLC IsNot Nothing Then m_colListColumns.Remove(CStr(objLC.ID))
        colListColumns.Remove(strLCKey)

        For Each strTypeKey As String In objTable.TypeListColumns.Keys
            colListColumns = objTable.TypeListColumns(strTypeKey)
            objLC = colListColumns(strLCKey)
            If objLC IsNot Nothing Then m_colListColumns.Remove(CStr(objLC.ID))
            colListColumns.Remove(strLCKey)
        Next
    End Sub

    ''' <summary>
    ''' Updates the field link collections when deleting a field link
    ''' </summary>
    Public Sub DRMDeleteFieldLink(ByVal objFieldLink As clsFieldLink)
        m_colFieldLinks.Remove(CStr(objFieldLink.ForeignKeyFieldID))

        If objFieldLink.ForeignKeyTable.IsLinkTable Then
            objFieldLink.IdentityTable.ManyToManyLinks.Remove(objFieldLink.KeyID)
        Else
            objFieldLink.IdentityTable.OneToManyLinks.Remove(objFieldLink.KeyID)
        End If

        objFieldLink.ForeignKeyTable.FieldLinks.Remove(CStr(objFieldLink.ForeignKeyFieldID))

        Dim strFilterKey1 As String

        strFilterKey1 = "L_" & objFieldLink.ID

        Dim colKeys As New List(Of String)
        For Each strKey As String In m_colFilters.Keys
            If strKey.StartsWith(strFilterKey1) Then
                colKeys.Add(strKey)
            End If
        Next

        For Each strKey As String In colKeys
            m_colFilters.Remove(strKey)
        Next
    End Sub

    ''' <summary>
    ''' Updates the table collections when deleting a table
    ''' </summary>
    Public Sub DRMDeleteTable(ByVal objTable As clsTable)
        m_colTables.Remove(objTable.DatabaseName)

        For Each objField As clsField In objTable.Fields.Values
            If objField.IsForeignKey Then
                If objTable.IsLinkTable Then
                    objField.FieldLink.IdentityTable.ManyToManyLinks.Remove(objField.FieldLink.KeyID)
                Else
                    objField.FieldLink.IdentityTable.OneToManyLinks.Remove(objField.FieldLink.KeyID)
                End If

                m_colFieldLinks.Remove(objField.KeyID)
            End If

            m_colFields.Remove(objField.KeyName)
        Next
    End Sub

    Public Sub DRMDeleteListColumn(ByVal objListColumn As clsListColumn)
        Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = Nothing

        If Not objListColumn.AppliesToTypeID = clsDBConstants.cintNULL Then
            colListColumns = objListColumn.Field.Table.TypeListColumns(CStr(objListColumn.AppliesToTypeID))
        End If

        If colListColumns Is Nothing Then
            colListColumns = objListColumn.Field.Table.ListColumns
        End If

        colListColumns.Remove(objListColumn.Field.TableID & "_" & objListColumn.FieldID)
        m_colListColumns.Remove(CStr(objListColumn.ID))
    End Sub

    Public Sub DRMDeleteFilter(ByVal objFilter As clsFilter)
        Dim colFilters As List(Of clsFilter)

        colFilters = clsFilter.GetList(objFilter, m_objDB)

        Dim strKey As String = objFilter.KeyName

        If colFilters Is Nothing OrElse colFilters.Count = 0 Then
            m_objDB.SysInfo.Filters.Remove(strKey)
        Else
            m_objDB.SysInfo.Filters(strKey) = colFilters
        End If
    End Sub

    Public Sub DRMDeleteCaption(ByVal objCaption As clsCaption)
        m_colCaptionStrings.Remove(m_objDB.Profile.LanguageID & "_" & objCaption.ID)
    End Sub
#End Region

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                If m_colTables IsNot Nothing Then
                    m_colTables.Dispose()
                    m_colTables = Nothing
                End If

                If m_colFields IsNot Nothing Then
                    m_colFields.Dispose()
                    m_colFields = Nothing
                End If

                If m_colFieldLinks IsNot Nothing Then
                    m_colFieldLinks.Dispose()
                    m_colFieldLinks = Nothing
                End If

                If m_objK1Configuration IsNot Nothing Then
                    m_objK1Configuration.Dispose()
                    m_objK1Configuration = Nothing
                End If

                If m_colMethods IsNot Nothing Then
                    m_colMethods.Dispose()
                    m_colMethods = Nothing
                End If

                If m_colLanguagesLoaded IsNot Nothing Then
                    m_colLanguagesLoaded.Dispose()
                    m_colLanguagesLoaded = Nothing
                End If

                If m_colCaptionStrings IsNot Nothing Then
                    m_colCaptionStrings.Dispose()
                    m_colCaptionStrings = Nothing
                End If

                If m_colSecGroups IsNot Nothing Then
                    m_colSecGroups.Dispose()
                    m_colSecGroups = Nothing
                End If

                If m_objK1Groups IsNot Nothing Then
                    m_objK1Groups.Dispose()
                    m_objK1Groups = Nothing
                End If

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

End Class
