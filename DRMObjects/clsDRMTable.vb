Public Class clsDRMTable
    Inherits clsDRMBase

#Region " Members "

    Private m_strDatabaseName As String
    Private m_intCaptionID As Integer = clsDBConstants.cintNULL
    Private m_strCaption As String
    Private m_eClassType As clsDBConstants.enumTableClass
    Private m_blnTypeDependent As Boolean
    Private m_strIconPath As String
    Private m_blnShowIcon As Boolean
    Private m_objDRMFieldType As clsDRMField
    Private m_colNewTableMethods As New FrameworkCollections.K1Collection(Of clsMethod.enumMethods)
    Private m_colInsertTableMethods As New FrameworkCollections.K1Dictionary(Of clsDRMTableMethod)
    Private m_colDeletedTableMethods As New FrameworkCollections.K1Collection(Of clsMethod.enumMethods)
    Private m_colFields As New FrameworkCollections.K1Collection(Of clsDRMField)
    Private m_colNewIndexes As New FrameworkCollections.K1Collection(Of clsTableIndex)
    Private m_colSecGroupIDs As FrameworkCollections.K1Collection(Of Integer)
#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new table
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal strTable As String, _
    ByVal intSecurityID As Integer, _
    ByVal strCaption As String, _
    ByVal eClassType As clsDBConstants.enumTableClass, _
    ByVal blnTypeDependent As Boolean, _
    ByVal strIconPath As String, _
    ByVal blnShowIcon As Boolean)
        MyBase.New(objDB, strTable, intSecurityID, clsDBConstants.cintNULL)
        m_strDatabaseName = strTable
        m_strCaption = strCaption
        m_eClassType = eClassType
        m_blnTypeDependent = blnTypeDependent
        m_strIconPath = strIconPath
        m_blnShowIcon = blnShowIcon
    End Sub

#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRM Field from an existing field database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objTable As clsTable)
        MyBase.New(objDB, objTable)
        m_strDatabaseName = objTable.DatabaseName
        m_intCaptionID = objTable.Caption.ID
        m_strCaption = objTable.Caption.GetString( _
            objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
        m_eClassType = objTable.ClassType
        m_blnTypeDependent = objTable.TypeDependent
        m_blnShowIcon = objTable.ShowIcon
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property DatabaseName() As String
        Get
            Return m_strDatabaseName
        End Get
        Set(ByVal value As String)
            m_strDatabaseName = value
            m_strExternalID = value
        End Set
    End Property

    Public Property Caption() As String
        Get
            Return m_strCaption
        End Get
        Set(ByVal value As String)
            m_strCaption = value
        End Set
    End Property

    Public Property ClassType() As clsDBConstants.enumTableClass
        Get
            Return m_eClassType
        End Get
        Set(ByVal value As clsDBConstants.enumTableClass)
            m_eClassType = value
        End Set
    End Property

    Public Property TypeDependent() As Boolean
        Get
            Return m_blnTypeDependent
        End Get
        Set(ByVal value As Boolean)
            m_blnTypeDependent = value
        End Set
    End Property

    Public Property IconPath() As String
        Get
            Return m_strIconPath
        End Get
        Set(ByVal value As String)
            m_strIconPath = value
        End Set
    End Property

    Public Property ShowIcon() As Boolean
        Get
            Return m_blnShowIcon
        End Get
        Set(ByVal value As Boolean)
            m_blnShowIcon = value
        End Set
    End Property

    Public ReadOnly Property Table() As clsTable
        Get
            Return CType(m_objDBObj, clsTable)
        End Get
    End Property

    Public ReadOnly Property NewTableMethods() As FrameworkCollections.K1Collection(Of clsMethod.enumMethods)
        Get
            Return m_colNewTableMethods
        End Get
    End Property

    Public ReadOnly Property DeletedTableMethods() As FrameworkCollections.K1Collection(Of clsMethod.enumMethods)
        Get
            Return m_colDeletedTableMethods
        End Get
    End Property

    Public Property SecurityGroupIDs() As FrameworkCollections.K1Collection(Of Integer)
        Get
            Return m_colSecGroupIDs
        End Get
        Set(ByVal value As FrameworkCollections.K1Collection(Of Integer))
            m_colSecGroupIDs = value
        End Set
    End Property
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False
        Dim objTable As clsTable
        Dim objEDOC As clsEDOC
        Dim objIcon As clsIcon
        Dim intCaptionID As Integer

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            '================================================================================
            'Create the table caption
            '================================================================================
            If m_intCaptionID = clsDBConstants.cintNULL OrElse m_intCaptionID = clsDBConstants.cintNULL Then
                If Not String.IsNullOrEmpty(m_strCaption) Then
                    intCaptionID = CreateCaption("Table - " & m_strDatabaseName, m_strCaption)
                End If
            Else
                intCaptionID = m_intCaptionID
                Dim strCaption As String = Table.Caption.GetString( _
                    m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)

                If Not String.IsNullOrEmpty(m_strCaption) AndAlso Not m_strCaption = strCaption Then
                    intCaptionID = CreateCaption(Table.Caption, m_strCaption)
                End If
            End If

            '================================================================================
            'Create the icon and icon edoc
            '================================================================================
            Dim intIconID As Integer = clsDBConstants.cintNULL
            If Not (Table Is Nothing OrElse Table.Icon Is Nothing) Then
                intIconID = Table.Icon.ID
            End If

            If Not (m_strIconPath Is Nothing OrElse m_strIconPath.Trim.Length = 0) Then
                Dim intEDOCID As Integer = clsDBConstants.cintNULL
                Dim intDownEDOCID As Integer = clsDBConstants.cintNULL
                Dim intOverEDOCID As Integer = clsDBConstants.cintNULL

                If Not (Table Is Nothing OrElse Table.Icon Is Nothing) Then
                    If Not Table.Icon.UpEDOC Is Nothing Then
                        intEDOCID = Table.Icon.UpEDOC.ID
                    End If
                    If Not Table.Icon.DownEDOC Is Nothing Then
                        intDownEDOCID = Table.Icon.DownEDOC.ID
                    End If
                    If Not Table.Icon.HoverEDOC Is Nothing Then
                        intOverEDOCID = Table.Icon.HoverEDOC.ID
                    End If
                End If

                objEDOC = New clsEDOC(m_objDB, intEDOCID, "Icon - " & m_strDatabaseName, _
                    m_intSecurityID, m_strIconPath)
                objEDOC.InsertUpdate()

                m_objDB.WriteBLOB(clsDBConstants.Tables.cEDOC, clsDBConstants.Fields.EDOC.cIMAGE, _
                    SqlDbType.Image, objEDOC.Size, objEDOC.ID, m_strIconPath)

                objIcon = New clsIcon(m_objDB, intIconID, m_strDatabaseName, m_intSecurityID, _
                    objEDOC.ID, intOverEDOCID, intDownEDOCID)
                objIcon.InsertUpdate()
                intIconID = objIcon.ID
            End If

            '================================================================================
            'Create the table
            '================================================================================
            objTable = New clsTable(m_objDB, m_intID, m_strExternalID, m_intSecurityID, _
                m_strDatabaseName, m_eClassType, intIconID, intCaptionID, _
                m_blnTypeDependent, m_blnShowIcon)
            objTable.InsertUpdate()

            If m_intID = clsDBConstants.cintNULL Then
                Me.SystemDB.CreateTable(m_strDatabaseName)
                InsertStandardFields(objTable)                

                '================================================================================
                'Add the table to the user's security group
                '================================================================================
                If Not m_colSecGroupIDs Is Nothing AndAlso m_colSecGroupIDs.Count >= 1 Then
                    For Each intSecID As Integer In m_colSecGroupIDs
                        m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLE, _
                            clsDBConstants.Fields.LinkSecurityGroupTable.cTABLEID, objTable.ID, _
                            clsDBConstants.Fields.LinkSecurityGroupTable.cSECURITYGROUPID, intSecID)
                    Next
                Else
                    If m_objDB.Profile.SecurityGroups Is Nothing AndAlso m_objDB.Profile.SecurityGroup IsNot Nothing Then
                        m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLE, _
                            clsDBConstants.Fields.LinkSecurityGroupTable.cTABLEID, objTable.ID, _
                            clsDBConstants.Fields.LinkSecurityGroupTable.cSECURITYGROUPID, _
                            m_objDB.Profile.SecurityGroup.ID)
                    Else
                        For Each intSecurityGroupID As Integer In m_objDB.Profile.LinkSecurityGroups.Values
                            m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLE, _
                                clsDBConstants.Fields.LinkSecurityGroupTable.cTABLEID, objTable.ID, _
                                clsDBConstants.Fields.LinkSecurityGroupTable.cSECURITYGROUPID, _
                                intSecurityGroupID)
                        Next
                    End If
                End If
            Else
                '================================================================================
                'Copy Fields over to new object
                '================================================================================
                For Each objKeyPair As Generic.KeyValuePair(Of String, clsField) In Me.Table.Fields
                    objTable.Fields(objKeyPair.Key) = objKeyPair.Value
                Next

                '================================================================================
                'Copy Field Links over to new object
                '================================================================================
                For Each objKeyPair As Generic.KeyValuePair(Of String, clsFieldLink) In Me.Table.FieldLinks
                    objTable.FieldLinks(objKeyPair.Key) = objKeyPair.Value
                Next

                '================================================================================
                'Delete Table Methods
                '================================================================================
                For Each eMethod As clsMethod.enumMethods In Me.m_colDeletedTableMethods
                    Dim objMethod As clsMethod = SystemDB.SysInfo.Methods(CStr(eMethod))
                    If Me.Table.TableMethods.ContainsKey(objMethod.ID.ToString()) Then
                        '-- Table has method so delete it
                        DBDeleteTableMethod(m_objDB, Me.Table.TableMethods(objMethod.ID.ToString()).ID)
                    End If
                Next
            End If

            '================================================================================
            '-- Insert/Update Fields
            '================================================================================
            For Each objDRMField As clsDRMField In Me.m_colFields
                objDRMField.Table = objTable
                objDRMField.InsertUpdate()
            Next

            If m_intID = clsDBConstants.cintNULL AndAlso _
            (Me.ClassType = clsDBConstants.enumTableClass.LINK_TABLE _
            OrElse Me.ClassType = clsDBConstants.enumTableClass.LINK_TABLE_ESSENTIAL) Then
                '================================================================================
                '-- Create SP for link tables
                '================================================================================
                SystemDB.CreateStandardSPs(objTable)
            End If

            '================================================================================
            '-- Insert Table Indexes (covers more then one field)
            '================================================================================
            For Each objIndex As clsTableIndex In Me.m_colNewIndexes
                objIndex.Create(Me.SystemDB)
            Next

            '================================================================================
            'Create Table Methods
            '================================================================================
            For Each eMethod As clsMethod.enumMethods In Me.m_colNewTableMethods
                Dim objMethod As clsMethod = m_objDB.SysInfo.Methods(CStr(eMethod))
                If Me.Table Is Nothing OrElse Not Me.Table.TableMethods.ContainsKey(CStr(objMethod.ID)) Then
                    '-- Table doesn't have the method so add it
                    DBInsertTableMethod(m_objDB, objTable, eMethod)
                End If
            Next

            For Each objDRMTableMethod In Me.m_colInsertTableMethods.Values
                objDRMTableMethod.InsertUpdate(m_objDB)
            Next

            m_objDB.SysInfo.DRMInsertUpdateTable(objTable)

            If objTable.DatabaseName = clsDBConstants.Tables.cEDOC Or objTable.DatabaseName = clsDBConstants.Tables.cMETADATAPROFILE Then
                UpdateRetentionTriggers(objTable.DatabaseName)
            End If

            m_objDBObj = objTable
            m_intID = objTable.ID
            m_intCaptionID = intCaptionID

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

    Private Sub InsertStandardFields(ByVal objTable As clsTable)
        Dim objDRMField As clsDRMField

        objDRMField = New clsDRMField(m_objDB, objTable, clsDRMField.enumStandardField.ID, m_intSecurityID)
        objDRMField.InsertUpdate()

        If Not (Me.ClassType = clsDBConstants.enumTableClass.LINK_TABLE _
        OrElse Me.ClassType = clsDBConstants.enumTableClass.LINK_TABLE_ESSENTIAL) Then
            '-- Create Extra fields for normal tables
            objDRMField = New clsDRMField(m_objDB, objTable, clsDRMField.enumStandardField.EXTERNALID, m_intSecurityID)
            objDRMField.SecurityGroupIDs = m_colSecGroupIDs
            objDRMField.InsertUpdate()
            Dim objDRMListCol As New clsDRMListColumn(m_objDB, objDRMField.Field.ExternalID, m_intSecurityID, _
                clsDBConstants.cintNULL, objDRMField.Field.ID, clsDBConstants.cintNULL, 100, "", 1)
            objDRMListCol.InsertUpdate(objTable)
            m_objDRMFieldType = New clsDRMField(m_objDB, objTable, clsDRMField.enumStandardField.TYPEID, m_intSecurityID)
            m_objDRMFieldType.SecurityGroupIDs = m_colSecGroupIDs
            m_objDRMFieldType.InsertUpdate()
            objDRMField = New clsDRMField(m_objDB, objTable, clsDRMField.enumStandardField.SECURITYID, m_intSecurityID)
            objDRMField.SecurityGroupIDs = m_colSecGroupIDs
            objDRMField.InsertUpdate()

            '================================================================================
            '-- Create SP for normal tables
            '================================================================================
            SystemDB.CreateStandardSPs(objTable)

            '================================================================================
            '-- Create a filter on the type id (only show types for the table)
            '================================================================================
            CreateTypeFilter(objTable)
        End If
    End Sub

    Private Sub CreateTypeFilter(ByVal objTable As clsTable)
        Dim objTypeTableIDField As clsField = m_objDB.SysInfo.Fields( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTYPE).ID & "_" & _
            clsDBConstants.Fields.Type.cTABLEID)

        Dim objFilter As New clsFilter(m_objDB, clsDBConstants.cintNULL, _
            "Type.TableID - " & objTable.DatabaseName & "[TABLE]", m_intSecurityID, _
            m_objDRMFieldType.ID, clsDBConstants.cintNULL, objTypeTableIDField.ID, _
            clsDBConstants.enumFilterTypes.TABLE, clsDBConstants.cstrNULL, _
            clsDBConstants.cintNULL, clsDBConstants.cintNULL)
        objFilter.InsertUpdate()
    End Sub

#End Region

#Region " Delete "

    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False
        Try
            Dim strSQL As String
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            'SS 28/11/11 (Bug#1100002149) - Delete the toolbar record as well. 
            'Begin
            strSQL = "DELETE FROM [" & clsDBConstants.Tables.cLINKSECURITYGROUPAPPMETHOD & "] " & _
                "WHERE [" & clsDBConstants.Fields.LinkSecurityGroupAppMethod.cAPPMETHODID & "] IN " & _
                "(SELECT " & clsDBConstants.Fields.cID & " FROM [" & clsDBConstants.Tables.cAPPLICATIONMETHOD & "] " & _
                "WHERE [" & clsDBConstants.Fields.ApplicationMethod.cTABLEID & "] = " & Table.ID.ToString & ")"
            m_objDB.ExecuteSQL(strSQL)

            strSQL = "DELETE FROM [" & clsDBConstants.Tables.cAPPLICATIONMETHOD & "] " & _
            "WHERE [" & clsDBConstants.Fields.ApplicationMethod.cTABLEID & "] = " & Table.ID.ToString
            m_objDB.ExecuteSQL(strSQL)
            'End
            RemoveStandardSPs()

            strSQL = String.Format("SELECT [{0}].[{2}] FROM [{0}] " & vbCrLf & _
                "INNER JOIN [{1}] ON [{1}].[{2}] = [{0}].[{3}]" & vbCrLf & _
                "INNER JOIN [{4}] ON [{4}].[{2}] = [{1}].[{5}]" & vbCrLf & _
                "WHERE [{4}].[{2}] = {6}", _
                clsDBConstants.Tables.cFILTER, clsDBConstants.Tables.cFIELD, _
                clsDBConstants.Fields.cID, clsDBConstants.Fields.Filter.cINITFILTERFIELDID, _
                clsDBConstants.Tables.cTABLE, clsDBConstants.Fields.Field.cTABLEID, Table.ID)

            Dim objDTRecords As DataTable = m_objDB.GetDataTableBySQL(strSQL)

            '-- Delete Filters prior to setting related records to NULL
            If objDTRecords.Rows.Count > 0 Then

                Dim arrIDs = (From r In objDTRecords.Rows.OfType(Of DataRow)()
                              Select CStr(r(clsDBConstants.Fields.cID))).ToArray

                m_objDB.DeleteRecordRange(clsDBConstants.Tables.cFILTER,
                                          clsDBConstants.Fields.cID,
                                          String.Join(", ", arrIDs))
            End If

            strSQL = String.Format("SELECT [{0}].[{2}] FROM [{0}] " & vbCrLf & _
                "INNER JOIN [{1}] ON [{1}].[{2}] = [{0}].[{3}]" & vbCrLf & _
                "INNER JOIN [{4}] ON [{4}].[{2}] = [{1}].[{5}]" & vbCrLf & _
                "INNER JOIN [{7}] ON [{7}].[{2}] = [{4}].[{8}]" & vbCrLf & _
                "WHERE [{7}].[{2}] = {6}", _
                clsDBConstants.Tables.cFILTER, clsDBConstants.Tables.cFIELDLINK, _
                clsDBConstants.Fields.cID, clsDBConstants.Fields.Filter.cINITFILTERFIELDLINKID, _
                clsDBConstants.Tables.cFIELD, clsDBConstants.Fields.FieldLink.cFOREIGNFIELDID, _
                Table.ID, clsDBConstants.Tables.cTABLE, clsDBConstants.Fields.Field.cTABLEID)

            objDTRecords = m_objDB.GetDataTableBySQL(strSQL)

            'Delete Filters prior to setting related records to NULL
            If objDTRecords.Rows.Count > 0 Then
                Dim arrIDs = (From r In objDTRecords.Rows.OfType(Of DataRow)()
                              Select CStr(r(clsDBConstants.Fields.cID))).ToArray
                m_objDB.DeleteRecordRange(clsDBConstants.Tables.cFILTER,
                                          clsDBConstants.Fields.cID,
                                          String.Join(", ", arrIDs))
            End If

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDTRecords IsNot Nothing Then
                objDTRecords.Dispose()
                objDTRecords = Nothing
            End If

            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cTABLE, Table.ID)

            'remove fields from tables which link to this table
            Dim colFieldLinks As New FrameworkCollections.K1Collection(Of clsFieldLink)

            For Each objFieldLink As clsFieldLink In Table.OneToManyLinks.Values
                colFieldLinks.Add(objFieldLink)
            Next
            For Each objFieldLink As clsFieldLink In colFieldLinks
                Dim objDRMField As New clsDRMField(m_objDB, objFieldLink.ForeignKeyField)
                objDRMField.Delete()
            Next
            colFieldLinks.Clear()

            'remove all link tables connected to this table
            Dim arrLinkTables As New Generic.List(Of String)
            For Each objFieldLink As clsFieldLink In Table.ManyToManyLinks.Values
                arrLinkTables.Add(objFieldLink.ForeignKeyTable.DatabaseName)
            Next

            For Each strTable As String In arrLinkTables
                Dim objTable As clsTable = m_objDB.SysInfo.Tables(strTable)
                If objTable IsNot Nothing Then
                    Dim objDRMTable As New clsDRMTable(m_objDB, objTable)
                    objDRMTable.Delete()
                End If
            Next

            For Each objField As clsField In Table.Fields.Values
                If objField.IsForeignKey Then
                    DeleteCaption(objField.FieldLink.Caption)
                End If
                DeleteCaption(objField.Caption)
            Next

            DeleteCaption(Table.Caption)

            SystemDB.DeleteTable(Table.DatabaseName)

            m_objDB.SysInfo.DRMDeleteTable(Table)

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)
            m_objDB.RefreshSysInfo()

            Throw
        End Try
    End Sub

    Private Sub RemoveStandardSPs()
        SystemDB.DeleteStoredProcedure(Table.DatabaseName & clsDBConstants.StoredProcedures.cINSERT)
        SystemDB.DeleteStoredProcedure(Table.DatabaseName & clsDBConstants.StoredProcedures.cUPDATE)
        SystemDB.DeleteStoredProcedure(Table.DatabaseName & clsDBConstants.StoredProcedures.cDELETE)
        SystemDB.DeleteStoredProcedure(Table.DatabaseName & clsDBConstants.StoredProcedures.cGETITEM)
        SystemDB.DeleteStoredProcedure(Table.DatabaseName & clsDBConstants.StoredProcedures.cGETLIST)
    End Sub

#End Region

#Region " Make Trigger "

    ''' <summary>
    ''' Create a trigger for the table in K1
    ''' </summary>
    ''' <param name="strTriggerName">Database name of the trigger</param>
    ''' <param name="strTriggerBody">SQL body of the trigger</param>
    ''' <param name="strTriggerAction">Trigger execution action</param>
    ''' <param name="blnOnInsert">Execute trigger on Insert?</param>
    ''' <param name="blnOnUpdate">Execute trigger on Update?</param>
    ''' <param name="blnOnDelete">Execute trigger on Delete?</param>
    ''' <param name="strExternalID">ExternalID for the Trigger in "Trigger" table</param>
    ''' <remarks></remarks>
    Public Function MakeTrigger(ByVal strTriggerName As String, ByVal strTriggerBody As String, _
      ByVal strTriggerAction As String, ByVal blnOnInsert As Boolean, ByVal blnOnUpdate As Boolean, _
      ByVal blnOnDelete As Boolean, ByVal strExternalID As String) As clsDRMTrigger
        Try
            Return MakeTrigger(strTriggerName, strTriggerBody, strTriggerAction, blnOnInsert, blnOnUpdate, _
                blnOnDelete, strExternalID, m_objDB.Profile.SecurityID)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Create a trigger for the table in K1
    ''' </summary>
    ''' <param name="strTriggerName">Database name of the trigger</param>
    ''' <param name="strTriggerBody">SQL body of the trigger</param>
    ''' <param name="strTriggerAction">Trigger execution action</param>
    ''' <param name="blnOnInsert">Execute trigger on Insert?</param>
    ''' <param name="blnOnUpdate">Execute trigger on Update?</param>
    ''' <param name="blnOnDelete">Execute trigger on Delete?</param>
    ''' <param name="strExternalID">ExternalID for the Trigger in "Trigger" table</param>
    ''' <param name="intSecurityID">SecurityID for the Trigger in "Trigger" table</param>
    ''' <remarks></remarks>
    Public Function MakeTrigger(ByVal strTriggerName As String, ByVal strTriggerBody As String, _
    ByVal strTriggerAction As String, ByVal blnOnInsert As Boolean, ByVal blnOnUpdate As Boolean, _
    ByVal blnOnDelete As Boolean, ByVal strExternalID As String, ByVal intSecurityID As Integer) As clsDRMTrigger
        Try
            Dim objDRMTrigger As New clsDRMTrigger(m_objDB, m_strDatabaseName, strTriggerName, strTriggerBody, _
                strTriggerAction, blnOnInsert, blnOnUpdate, blnOnDelete, strExternalID, intSecurityID)

            objDRMTrigger.InsertUpdate()

            Return objDRMTrigger
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Delete Trigger "

    ''' <summary>
    ''' Deletes a trigger from K1 Trigger table and the database
    ''' </summary>
    ''' <param name="strTriggerName">Name of the trigger</param>
    ''' <remarks></remarks>
    Public Sub DeleteTrigger(ByVal strTriggerName As String)
        Try
            Dim objDRMTrigger As New clsDRMTrigger(m_objDB, strTriggerName)

            objDRMTrigger.Delete()
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
        If Not m_objDRMFieldType Is Nothing Then
            m_objDRMFieldType.Dispose()
            m_objDRMFieldType = Nothing
        End If
    End Sub
#End Region

    Public Sub ClearTableMethod()
        Me.m_colNewTableMethods.Clear()
        Me.m_colDeletedTableMethods.Clear()
    End Sub

    Public Function AddUpdateTableMethod(ByVal eMethod As clsMethod.enumMethods) As clsDRMTableMethod
        Dim objDRMTableMethod As clsDRMTableMethod = Nothing

        If Not Me.m_colNewTableMethods.Contains(eMethod) Then
            Dim objMethod As clsMethod = m_objDB.SysInfo.Methods(CStr(eMethod))
            Dim objTableMethod As clsTableMethod = Me.Table.TableMethods(objMethod.ID.ToString)
            If objTableMethod Is Nothing Then
                objDRMTableMethod = New clsDRMTableMethod(m_objDB, Me.Table, objMethod, False, False, Nothing)
            Else
                objDRMTableMethod = New clsDRMTableMethod(m_objDB, objTableMethod, Me.Table.ID)
                objDRMTableMethod.Audit = objTableMethod.Audit
                objDRMTableMethod.AuditData = objTableMethod.AuditData
            End If

            Me.m_colInsertTableMethods.Add(eMethod.ToString(), objDRMTableMethod)
        End If

        If Me.m_colDeletedTableMethods.Contains(eMethod) Then
            '-- remove method from the delete list if we are adding it
            Me.m_colDeletedTableMethods.Remove(eMethod)
        End If

        Return objDRMTableMethod
    End Function

    Public Sub AddTableMethod(ByVal eMethod As clsMethod.enumMethods)
        If Not Me.m_colNewTableMethods.Contains(eMethod) Then
            Me.m_colNewTableMethods.Add(eMethod)
        End If

        If Me.m_colDeletedTableMethods.Contains(eMethod) Then
            '-- remove method from the delete list if we are adding it
            Me.m_colDeletedTableMethods.Remove(eMethod)
        End If
    End Sub

    Public Sub DeleteTableMethod(ByVal eMethod As clsMethod.enumMethods)
        If Not Me.m_colDeletedTableMethods.Contains(eMethod) Then
            Me.m_colDeletedTableMethods.Add(eMethod)
        End If

        If Me.m_colNewTableMethods.Contains(eMethod) Then
            '-- remove method from the delete list if we are adding it
            Me.m_colNewTableMethods.Remove(eMethod)
        End If
    End Sub

    Private Sub DBInsertTableMethod(ByVal objK1DB As clsDB_System, ByVal objTable As clsTable, _
    ByVal eMethod As clsMethod.enumMethods)
        Try
            Dim colParameters As New clsDBParameterDictionary
            Dim eDirection As Data.ParameterDirection = ParameterDirection.Input
            Dim objMethod As clsMethod = objK1DB.SysInfo.Methods(CStr(eMethod))
            Dim strName As String = objTable.DatabaseName & "." & objMethod.ExternalID
            Dim strMethodID As Integer = objMethod.ID

            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cEXTERNALID, strName, eDirection, SqlDbType.NVarChar))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cSECURITYID, objTable.SecurityID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cTABLEID, objTable.ID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cMETHODID, strMethodID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cAUDIT, False, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cID, clsDBConstants.cintNULL, ParameterDirection.Output, SqlDbType.Int))

            objK1DB.ExecuteProcedure(clsDBConstants.Tables.cTABLEMETHOD & _
                clsDBConstants.StoredProcedures.cINSERT, colParameters)

            Dim intID As Integer = CInt(colParameters(clsDBConstants.Fields.cID).Value)

            If Not m_colSecGroupIDs Is Nothing AndAlso m_colSecGroupIDs.Count >= 1 Then
                For Each intSecID As Integer In m_colSecGroupIDs
                    m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                        clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID, intID, _
                        clsDBConstants.Fields.LinkSecurityGroupTableMethod.cSECURITYGROUPID, intSecID)
                Next
            Else
                For Each intSecurityGroupID As Integer In objK1DB.Profile.LinkSecurityGroups.Values
                    m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                        clsDBConstants.Fields.LinkSecurityGroupTableMethod.cSECURITYGROUPID, _
                        intSecurityGroupID, _
                        clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID, intID)
                Next
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub DBDeleteTableMethod(ByVal objK1DB As clsDB_System, ByVal intID As Integer)
        Try
            objK1DB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cACCESSRIGHTMETHOD & "] " & _
                "WHERE [" & clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID & "] = " & intID)

            Dim colParameters As New clsDBParameterDictionary

            '-- Delete dependent record first
            objK1DB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD & "] " & _
                "WHERE [" & clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID & "] = " & intID)

            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cID, intID, ParameterDirection.Input, SqlDbType.Int))

            objK1DB.ExecuteProcedure(clsDBConstants.Tables.cTABLEMETHOD & _
                clsDBConstants.StoredProcedures.cDELETE, colParameters)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub AddField(ByVal objField As clsField, ByVal LinkVisible As Boolean, ByVal strCaption As String, _
    ByVal intSortOrder As Integer)
        Try
            Dim objDRMField As K1Library.clsDRMField = New clsDRMField(SystemDB, objField)
            objDRMField.FieldLinkVisible = LinkVisible
            objDRMField.FieldLinkCaption = strCaption
            objDRMField.FieldLinkSortOrder = intSortOrder

            m_colFields.Add(objDRMField)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub AddNewField(ByVal strFieldName As String, ByVal strIdentityTable As String, _
    ByVal LinkVisible As Boolean, ByVal strCaption As String, ByVal intSortOrder As Integer, ByVal blnIsNullable As Boolean)
        Try
            Dim objDRMField As K1Library.clsDRMField = New clsDRMField(SystemDB, Me.Table, strFieldName, _
                strIdentityTable, m_intSecurityID, LinkVisible, strCaption, intSortOrder, blnIsNullable)
            objDRMField.SortOrder = m_colFields.Count + 2
            m_colFields.Add(objDRMField)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub AddTableIndex(ByVal colFields As FrameworkCollections.K1Collection(Of String), _
    ByVal blnClustered As Boolean, ByVal blnUnique As Boolean)
        m_colNewIndexes.Add(New clsTableIndex(m_strDatabaseName, colFields, blnClustered, blnUnique, False, False))
    End Sub

    Public Sub AddTableIndex(ByVal strField As String, _
    ByVal blnClustered As Boolean, ByVal blnUnique As Boolean)
        Dim colFields As New FrameworkCollections.K1Collection(Of String)
        colFields.Add(strField)

        AddTableIndex(colFields, blnClustered, blnUnique)
    End Sub

    Public Sub AddLinkTableIndex(ByVal strField1 As String, ByVal strField2 As String)
        Dim colFields As New FrameworkCollections.K1Collection(Of String)
        colFields.Add(strField1)
        colFields.Add(strField2)

        m_colNewIndexes.Add(New clsTableIndex(m_strDatabaseName, colFields, False, True, False, False))
    End Sub

    Private Sub UpdateRetentionTriggers(ByVal strTableName As String)
        Dim strTriggerName As String = "tr" & strTableName & "_Retention"
        If m_objDB.TriggerExists(strTriggerName) Then
            Dim objTrigger As New clsDRMTrigger(m_objDB, strTriggerName)
            Dim strOriginalSQL As String = objTrigger.Body

            ''2017-03-16 -- James Dodd -- Fix for Bug #1700003274 -- Changes decimal to 0 instead of '' for comparison
            Dim dtColumns As DataTable = SystemDB.GetDataTableBySQL("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE (TABLE_NAME = '" & strTableName & "') AND ((DATA_TYPE <> 'image') AND (DATA_TYPE <> 'ntext') AND (DATA_TYPE <> 'text'))")

            ' Ara Melkonian - 2100003727
            ' Added AllocatedSpaceID to excluded columns
            Dim ExcludedColumns() As String = {"BoxID", "ParentMetadataProfileID", "CurrentLocation", "LastMovedDate", "LastModifiedDate",
                "LastAttachDate", "LastRequestedDate", "RetentionCodeID", "RetentionCode2ID", "LifeCycle1DueDate", "LifeCycle2DueDate",
                "VitalRecordID", "VitalRecordLastReviewDate", "VitalRecordNextReviewDate", "ReplyByDate", "ReplySentDate",
                "MetadataProfileID", "CheckedOut", "LastCheckOutPersonID", "LastCheckOutTime", "CheckedInPersonID", "isLatestVersion", "LifeCycleDueDate",
                "ThumbnailFailCount", "Thumbnail", "HoverThumbnail", "HoverFailCount", "AllocatedSpaceID"}

            Dim strNewSQL As String = Nothing
            For Each dtRow As DataRow In dtColumns.Rows
                Dim Result As String = Array.Find(ExcludedColumns, Function(s) (s = (CStr(dtRow.Item(0)))))

                If Result Is Nothing Then
                    If strNewSQL IsNot Nothing Then
                        strNewSQL &= " AND "
                    End If
                    Dim strInsertChar As String = "''"
                    If CStr(dtRow.Item(1)) = "decimal" Then
                        strInsertChar = "0"
                    End If
                    strNewSQL &= " ISNULL(Inserted.[" & CStr(dtRow.Item(0)) & "]," & strInsertChar & ")=ISNULL(Deleted.[" & CStr(dtRow.Item(0)) & "]," & strInsertChar & ")"
                End If
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If dtColumns IsNot Nothing Then
                dtColumns.Dispose()
                dtColumns = Nothing
            End If

            strNewSQL = "SELECT @COUNTER = COUNT(*) from Inserted Join Deleted on Inserted.[ID]=Deleted.[ID] WHERE" & strNewSQL & vbCrLf

            Dim pos As Integer = strOriginalSQL.ToUpper.IndexOf("SELECT @COUNTER = COUNT(*) FROM INSERTED")

            '2019-06-28 -- Emmanuel - 1900003524 Kludgy fix to an awful bit of design.   The above does not get found if there is a CR in the middle.
            '                                    Also updated so that its case insensitive.   Maybe Regex will help.   Maybe kill it with fire. 
            If pos < 0 Then
                pos = strOriginalSQL.ToUpper.IndexOf($"SELECT @COUNTER = COUNT(*){vbCrLf}FROM INSERTED", 0)
            End If
            Dim strSQL As String = ""

            If pos > 0 Then
                strSQL = strOriginalSQL.Substring(0, pos)
            End If

            If strSQL.LastIndexOf("E") = strSQL.Length - 1 Then
                strSQL &= vbCrLf
            End If

            strSQL &= strNewSQL
            strSQL &= strOriginalSQL.Substring(strOriginalSQL.IndexOf("--DYNAMIC SQL SCRIPT END - DO NOT CHANGE"))

            objTrigger.Body = strSQL
            objTrigger.InsertUpdate()
        End If
    End Sub
End Class
