#Region " File Information "

'==============================================================================
' This class contains all the mask collections for a particular table
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       09/03/2007  Implemented.
'==============================================================================

#End Region

#End Region

Public Class clsTableMask
    Implements IDisposable

#Region " Members "

    Private m_objTable As clsTable
    Private m_colMaskField As clsMaskFieldDictionary
    Private m_colMaskOneToMany As clsMaskFieldLinkDictionary
    Private m_colMaskManyToMany As clsMaskFieldLinkDictionary
    Private m_colMaskBase As FrameworkCollections.K1Collection(Of clsMaskBase) 'all mask objects in order
    Private m_eMaskType As enumMaskType
    Private m_objParent As clsDBObject
    Private m_objParentMask As clsMaskBase
    Private m_blnDisposedValue As Boolean = False
    Private m_intTypeID As Integer
    Private m_colNewFreeTextIDs As Hashtable
    Private m_colDeleteFreeTextIDs As Hashtable
#End Region

#Region " Constants "

    Public Class Columns
        Public Const cCOL_SORTORDER As String = "SortOrder"
        Public Const cCOL_OBJTYPE As String = "ObjType"
        Public Const cCOL_KEY As String = "Key"
        Public Const cCOL_ID As String = "ID"
    End Class
#End Region

#Region " Enumerations "

    Public Enum enumMaskType
        VIEW = 0
        ADD = 1
        MODIFY = 2
        SEARCH = 3
        MULTI_MODIFY = 4
        MULTI_ADD = 5
    End Enum
#End Region

#Region " Constructors "

    ''' <summary>
    ''' Creates a new table mask object
    ''' </summary>
    ''' <param name="objTable">The table we are creating mask objects for</param>
    ''' <param name="eMaskType">What type of mask screen is being used (ADD, MODIFY, etc.)</param>
    ''' <param name="intID">The ID of the record we are modifying, viewing, etc.</param>
    ''' <param name="intTypeID">Needed if the table is type dependent</param>
    ''' <param name="objParent">This is necessary for some types of autofills</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objTable As clsTable, ByVal eMaskType As enumMaskType, _
    Optional ByVal intID As Integer = clsDBConstants.cintNULL, _
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL, _
    Optional ByVal objParent As clsDBObject = Nothing, _
    Optional ByVal objParentMask As clsMaskBase = Nothing)
        m_objTable = objTable
        m_eMaskType = eMaskType
        m_objParent = objParent
        m_objParentMask = objParentMask

        'If objTable.TypeDependent AndAlso _
        'intTypeID = clsDBConstants.cintNULL AndAlso _
        'Not intID = clsDBConstants.cintNULL Then
        '    Dim objDT As DataTable = objTable.Database.GetItem(objTable.DatabaseName, intID)

        '    If objDT.Rows.Count = 0 Then
        '        objDT.Dispose()
        '        objDT = Nothing
        '        intID = clsDBConstants.cintNULL
        '    Else
        '        intTypeID = CInt(clsDB.NullValue(objDT.Rows(0)(clsDBConstants.Fields.cTYPEID), clsDBConstants.cintNULL))
        '    End If
        'End If

        m_intTypeID = intTypeID

        m_colMaskField = clsMaskField.CreateMaskCollection( _
            objTable, eMaskType, intID, intTypeID, objParent, objParentMask)

        m_intTypeID = CInt(m_colMaskField.GetMaskValue(clsDBConstants.Fields.cTYPEID, m_intTypeID))
        m_colMaskField.TableMask = Me

        m_colMaskOneToMany = clsMaskFieldLink.CreateMaskCollection( _
            objTable, clsMaskBase.enumMaskObjectType.ONETOMANY, eMaskType, m_intTypeID)
        m_colMaskManyToMany = clsMaskFieldLink.CreateMaskCollection( _
            objTable, clsMaskBase.enumMaskObjectType.MANYTOMANY, eMaskType, m_intTypeID)
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' The table this pertains to
    ''' </summary>
    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objTable
        End Get
    End Property

    ''' <summary>
    ''' The collection of all mask objects in sort order
    ''' </summary>
    Public ReadOnly Property MaskCollection() As FrameworkCollections.K1Collection(Of clsMaskBase)
        Get
            If m_colMaskBase Is Nothing Then
                CreateMaskCollection()
            End If
            Return m_colMaskBase
        End Get
    End Property

    ''' <summary>
    ''' Key is of type String and its value is the DatabaseName property of the clsMaskField.
    ''' </summary>
    Public ReadOnly Property MaskFieldCollection() As clsMaskFieldDictionary
        Get
            Return m_colMaskField
        End Get
    End Property

    ''' <summary>
    ''' The mask field link dictionary pertaining to one-to-many field links
    ''' </summary>
    Public ReadOnly Property MaskOneToManyCollection() As clsMaskFieldLinkDictionary
        Get
            Return m_colMaskOneToMany
        End Get
    End Property

    ''' <summary>
    ''' The mask field link dictionary pertaining to many-to-many field links
    ''' </summary>
    Public ReadOnly Property MaskManyToManyCollection() As clsMaskFieldLinkDictionary
        Get
            Return m_colMaskManyToMany
        End Get
    End Property

    ''' <summary>
    ''' Returns the ID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property ID() As Integer
        Get
            If m_colMaskField(clsDBConstants.Fields.cID) Is Nothing OrElse _
            m_colMaskField(clsDBConstants.Fields.cID).Value1.Value Is Nothing Then
                Return clsDBConstants.cintNULL
            Else
                Return CInt(m_colMaskField(clsDBConstants.Fields.cID).Value1.Value)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the ExternalID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property ExternalID() As String
        Get
            If m_colMaskField(clsDBConstants.Fields.cEXTERNALID) Is Nothing OrElse _
            m_colMaskField(clsDBConstants.Fields.cEXTERNALID).Value1.Value Is Nothing Then
                Return clsDBConstants.cstrNULL
            Else
                Return CStr(m_colMaskField(clsDBConstants.Fields.cEXTERNALID).Value1.Value)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the SecurityID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property SecurityID() As Integer
        Get
            If m_colMaskField(clsDBConstants.Fields.cSECURITYID) Is Nothing OrElse _
            m_colMaskField(clsDBConstants.Fields.cSECURITYID).Value1.Value Is Nothing Then
                Return clsDBConstants.cintNULL
            Else
                Return CInt(m_colMaskField(clsDBConstants.Fields.cSECURITYID).Value1.Value)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the SecurityID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property TypeID() As Integer
        Get
            If m_colMaskField(clsDBConstants.Fields.cTYPEID) Is Nothing OrElse _
            m_colMaskField(clsDBConstants.Fields.cTYPEID).Value1.Value Is Nothing Then
                Return clsDBConstants.cintNULL
            Else
                Return CInt(m_colMaskField(clsDBConstants.Fields.cTYPEID).Value1.Value)
            End If
        End Get
    End Property

    ''' <summary>
    ''' The parent table mask object of this object (if coming from a prior mask screen)
    ''' </summary>
    Public ReadOnly Property Parent() As clsDBObject
        Get
            Return m_objParent
        End Get
    End Property

    Public ReadOnly Property ParentMask() As clsMaskBase
        Get
            Return m_objParentMask
        End Get
    End Property

    Public Property MaskType() As enumMaskType
        Get
            Return m_eMaskType
        End Get
        Set(ByVal value As enumMaskType)
            m_eMaskType = value
            UpdateMaskCollection(m_colMaskField, value)
            UpdateMaskCollection(m_colMaskOneToMany, value)
            UpdateMaskCollection(m_colMaskManyToMany, value)
        End Set
    End Property
#End Region

#Region " Methods "

    Private Sub UpdateMaskCollection(ByVal colMasks As IDictionary, ByVal eMaskType As enumMaskType)
        If colMasks IsNot Nothing Then
            For Each objMaskBase As clsMaskBase In colMasks.Values
                objMaskBase.MaskType = eMaskType
            Next
        End If
    End Sub

    Public Function IsDirty() As Boolean
        Dim blnDirty As Boolean

        If (m_eMaskType = clsTableMask.enumMaskType.ADD OrElse _
        m_eMaskType = clsTableMask.enumMaskType.MODIFY OrElse _
        m_eMaskType = clsTableMask.enumMaskType.MULTI_ADD OrElse _
        m_eMaskType = clsTableMask.enumMaskType.MULTI_MODIFY) Then
            For Each objMaskField As clsMaskField In m_colMaskField.Values
                If (Not m_eMaskType = enumMaskType.MULTI_MODIFY OrElse _
                (m_eMaskType = clsTableMask.enumMaskType.MULTI_MODIFY AndAlso _
                objMaskField.CheckState = Windows.Forms.CheckState.Checked)) AndAlso _
                objMaskField.Value1.IsDirty Then
                    blnDirty = True
                    Exit For
                End If
            Next

            If Not blnDirty Then
                For Each objMaskFL As clsMaskFieldLink In m_colMaskManyToMany.Values
                    If (Not m_eMaskType = enumMaskType.MULTI_MODIFY OrElse _
                    (m_eMaskType = clsTableMask.enumMaskType.MULTI_MODIFY AndAlso _
                    objMaskFL.CheckState = Windows.Forms.CheckState.Checked)) AndAlso _
                    objMaskFL.IsDirty Then
                        blnDirty = True
                        Exit For
                    End If
                Next
            End If
        End If

        Return blnDirty
    End Function

#Region " Mask Sorting "

    ''' <summary>
    ''' Creates the Complete mask object collection in sort order
    ''' </summary>
    Private Sub CreateMaskCollection()
        m_colMaskBase = New FrameworkCollections.K1Collection(Of clsMaskBase)

        Dim objDV As DataView = GetMaskOrderDataView(m_objTable, True, False, m_intTypeID)

        For intLoop As Integer = 0 To objDV.Count - 1
            Dim objRow As DataRow = objDV(intLoop).Row
            Dim eObjType As clsMaskBase.enumMaskObjectType = _
                CType(objRow(clsTableMask.Columns.cCOL_OBJTYPE),  _
                clsMaskBase.enumMaskObjectType)
            Dim strKey As String = CType(objRow(clsTableMask.Columns.cCOL_KEY), String)

            Select Case eObjType
                Case clsMaskBase.enumMaskObjectType.FIELD
                    Dim objField As clsField = m_objTable.Fields(strKey)
                    m_colMaskBase.Add(m_colMaskField(objField.DatabaseName))

                Case clsMaskBase.enumMaskObjectType.ONETOMANY
                    Dim objFieldLink As clsFieldLink = m_objTable.OneToManyLinks(strKey)
                    m_colMaskBase.Add(m_colMaskOneToMany(objFieldLink.KeyID))

                Case clsMaskBase.enumMaskObjectType.MANYTOMANY
                    Dim objFieldLink As clsFieldLink = m_objTable.ManyToManyLinks(strKey)
                    m_colMaskBase.Add(m_colMaskManyToMany(objFieldLink.KeyID))

            End Select
        Next
    End Sub

    ''' <summary>
    ''' Creates a DataView of the mask objects
    ''' </summary>
    Public Shared Function GetMaskOrderDataView(ByVal objTable As clsTable, _
    ByVal blnIncludeFieldLinks As Boolean, ByVal blnOnlyListColumns As Boolean, _
    ByVal intTypeID As Integer, Optional ByVal blnShowAll As Boolean = False) As DataView
        Dim objProfile As clsUserProfile = objTable.Database.Profile
        Dim dtMaskOrder As New DataTable

        dtMaskOrder.Columns.Add(New DataColumn(Columns.cCOL_SORTORDER, GetType(System.Int32)))
        dtMaskOrder.Columns.Add(New DataColumn(Columns.cCOL_OBJTYPE, GetType(System.Int16)))
        dtMaskOrder.Columns.Add(New DataColumn(Columns.cCOL_KEY, GetType(System.String)))
        dtMaskOrder.Columns.Add(New DataColumn(Columns.cCOL_ID, GetType(System.Int64)))

        Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = _
            objTable.GetListColumns(intTypeID)

        For Each objField As clsField In objTable.Fields.Values
            If (blnShowAll AndAlso Not objField.IsIdentityField) OrElse _
            ((Not blnOnlyListColumns OrElse colListColumns(objField.TableID & "_" & objField.ID) IsNot Nothing) AndAlso _
            objField.IsVisible AndAlso _
            objProfile.HasAccess(objField.SecurityID) AndAlso _
            objProfile.LinkFields(objField.KeyID) IsNot Nothing) Then
                dtMaskOrder.Rows.Add(CreateDataRow(dtMaskOrder, objField, intTypeID))
            End If
        Next

        If blnIncludeFieldLinks Then
            For Each objFieldLink As clsFieldLink In objTable.OneToManyLinks.Values
                If blnShowAll OrElse _
                (objFieldLink.IsVisible AndAlso _
                objProfile.HasAccess(objFieldLink.SecurityID)) Then
                    dtMaskOrder.Rows.Add(CreateDataRow(dtMaskOrder, _
                        objFieldLink, clsMaskBase.enumMaskObjectType.ONETOMANY, intTypeID))
                End If
            Next

            For Each objFieldLink As clsFieldLink In objTable.ManyToManyLinks.Values
                If blnShowAll OrElse _
                (objFieldLink.IsVisible AndAlso _
                objProfile.HasAccess(objFieldLink.SecurityID)) Then
                    dtMaskOrder.Rows.Add(CreateDataRow(dtMaskOrder, _
                        objFieldLink, clsMaskBase.enumMaskObjectType.MANYTOMANY, intTypeID))
                End If
            Next
        End If

        dtMaskOrder.DefaultView.Sort = Columns.cCOL_SORTORDER & ", " & Columns.cCOL_ID

        Return dtMaskOrder.DefaultView
    End Function

    Private Shared Function CreateDataRow(ByVal dtMaskOrder As DataTable, _
    ByVal objField As clsField, ByVal intTypeID As Integer) As DataRow
        Dim objDR As DataRow

        objDR = dtMaskOrder.NewRow()

        Dim objTF As clsTypeField = Nothing
        If Not intTypeID = clsDBConstants.cintNULL Then
            objTF = objField.TypeFieldInfos(CStr(intTypeID))
        End If
        If objTF Is Nothing Then
            objDR.Item(Columns.cCOL_SORTORDER) = objField.SortOrder
        Else
            objDR.Item(Columns.cCOL_SORTORDER) = objTF.SortOrder
        End If
        objDR.Item(Columns.cCOL_OBJTYPE) = clsMaskBase.enumMaskObjectType.FIELD
        objDR.Item(Columns.cCOL_KEY) = objField.KeyID
        objDR.Item(Columns.cCOL_ID) = objField.ID

        Return objDR
    End Function

    Private Shared Function CreateDataRow(ByVal dtMaskOrder As DataTable, _
    ByVal objFieldLink As clsFieldLink, ByVal eObjType As clsMaskBase.enumMaskObjectType, _
    ByVal intTypeID As Integer) As DataRow
        Dim objDR As DataRow

        objDR = dtMaskOrder.NewRow()

        Dim objTFL As clsTypeFieldLink = Nothing
        If Not intTypeID = clsDBConstants.cintNULL Then
            objTFL = objFieldLink.TypeFieldLinkInfos(CStr(intTypeID))
        End If
        If objTFL Is Nothing Then
            objDR.Item(Columns.cCOL_SORTORDER) = objFieldLink.SortOrder
        Else
            objDR.Item(Columns.cCOL_SORTORDER) = objTFL.SortOrder
        End If
        objDR.Item(Columns.cCOL_OBJTYPE) = eObjType
        objDR.Item(Columns.cCOL_KEY) = objFieldLink.KeyID
        objDR.Item(Columns.cCOL_ID) = objFieldLink.ForeignKeyFieldID

        Return objDR
    End Function
#End Region

#Region " Insert\Update "

    ''' <summary>
    ''' Either creates a new record (if mask type is add) or updates an existing record (modify)
    ''' </summary>
    Public Sub InsertUpdate(Optional ByVal blnUseAutoNumbering As Boolean = True)
        Dim blnCreatedTransaction As Boolean = False
        Dim objDB As clsDB = m_objTable.Database

        Try
            If Not objDB.HasTransaction Then
                objDB.BeginTransaction()
                blnCreatedTransaction = True
            End If

            HandleFreeTextFields()

            Try
                Select Case m_eMaskType
                    Case enumMaskType.ADD
                        m_colMaskField.Insert(objDB, blnUseAutoNumbering)
                        m_colMaskManyToMany.InsertUpdate(objDB, Me.ID)

                    Case enumMaskType.MODIFY
                        m_colMaskField.Update(objDB)
                        m_colMaskManyToMany.InsertUpdate(objDB, Me.ID)

                    Case Else
                        Throw New Exception("The Mask Type is not set correctly for an InsertUpdate operation")

                End Select
            Catch ex As Exception
                ResetFreeTextFields()

                Throw ex
            End Try

            CleanUpFreeTextFields()

            If blnCreatedTransaction Then
                objDB.EndTransaction(True)
            End If

            For Each objMask As clsMaskField In m_colMaskField.Values
                objMask.Value1.IsDirty = False
            Next

            For Each objMaskFL As clsMaskFieldLink In m_colMaskManyToMany.Values
                objMaskFL.IsDirty = False
            Next
        Catch ex As Exception
            If blnCreatedTransaction Then
                objDB.EndTransaction(False)
            End If

            Throw
        End Try
    End Sub

#Region " Free Text Fields "

    Private Sub HandleFreeTextFields()
        For Each objMask As clsMaskField In m_colMaskField.Values
            If objMask.Field.IsForeignKey AndAlso objMask.AllowFreeTextEntry Then
                If objMask.Value1.FreeText IsNot Nothing Then
                    InsertUpdateFreeTextRecord(objMask)
                Else
                    FlagFreeTextDeletion(objMask)
                End If
            End If
        Next
    End Sub

    Private Sub InsertUpdateFreeTextRecord(ByVal objMask As clsMaskField)
        Dim intID As Integer

        If objMask.Value1.Value Is Nothing Then
            intID = clsDBConstants.cintNULL
        Else
            intID = CInt(objMask.Value1.Value)
        End If

        If Not intID = clsDBConstants.cintNULL Then
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@ID", intID))
            colParams.Add(New clsDBParameter("@MDPID", Me.ID))

            'Check to see if any other records are referecing this free text record
            Dim intCount As Integer = m_objTable.Database.ExecuteScalar( _
                "SELECT COUNT([" & clsDBConstants.Fields.cID & "]) FROM [" & _
                m_objTable.DatabaseName & "] WHERE [" & objMask.Field.DatabaseName & "] = @ID " & _
                "AND NOT [" & clsDBConstants.Fields.cID & "] = @MDPID", _
                colParams)

            If intCount >= 1 Then
                intID = clsDBConstants.cintNULL 'force it to create a new free text title if multiple records use this ID
            End If
        End If

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            objMask.Field.FieldLink.IdentityTable, intID)

        Dim intSecurityID As Integer = CInt(m_colMaskField.GetMaskValue(clsDBConstants.Fields.cSECURITYID))

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, objMask.Value1.FreeText)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, intSecurityID)

        If intID = clsDBConstants.cintNULL Then
            intID = colMasks.Insert(m_objTable.Database)

            clsAuditTrail.CreateTableMethodRecord(m_objTable.Database, clsMethod.enumMethods.cADD, _
                objMask.Field.FieldLink.IdentityTable, intID, objMask.Value1.FreeText)

            If m_colNewFreeTextIDs Is Nothing Then
                m_colNewFreeTextIDs = New Hashtable
            End If

            m_colNewFreeTextIDs.Add(objMask.Field.DatabaseName, intID)
        Else
            colMasks.Update(m_objTable.Database)

            clsAuditTrail.CreateTableMethodRecord(m_objTable.Database, clsMethod.enumMethods.cMODIFY, _
                objMask.Field.FieldLink.IdentityTable, intID, objMask.Value1.FreeText)
        End If

        objMask.Value1.Value = intID
        objMask.Value1.Display = objMask.Value1.FreeText
        objMask.Value1.ObjectSecurityID = intSecurityID
    End Sub

    Private Sub ResetFreeTextFields()
        If m_colNewFreeTextIDs IsNot Nothing Then
            For Each strKey As String In m_colNewFreeTextIDs.Keys
                Dim objMask As clsMaskField = m_colMaskField(strKey)

                If objMask IsNot Nothing AndAlso objMask.Value1 IsNot Nothing Then
                    objMask.Value1.Value = Nothing
                    objMask.Value1.Display = ""
                    objMask.Value1.ObjectSecurityID = clsDBConstants.cintNULL
                End If
            Next

            m_colNewFreeTextIDs.Clear()
        End If

        If m_colDeleteFreeTextIDs IsNot Nothing Then
            For Each strKey As String In m_colDeleteFreeTextIDs.Keys
                Dim objMask As clsMaskField = m_colMaskField(strKey)

                If objMask IsNot Nothing AndAlso objMask.Value1 IsNot Nothing Then
                    objMask.Value1.Value = m_colDeleteFreeTextIDs(strKey)
                End If
            Next

            m_colDeleteFreeTextIDs.Clear()
        End If
    End Sub

    Private Sub CleanUpFreeTextFields()
        If m_colDeleteFreeTextIDs IsNot Nothing Then
            For Each strKey As String In m_colDeleteFreeTextIDs.Keys
                Dim objMask As clsMaskField = m_colMaskField(strKey)

                Try
                    m_objTable.Database.DeleteRecord(objMask.Field.FieldLink.IdentityTable.DatabaseName, CInt(m_colDeleteFreeTextIDs(strKey)))
                Catch ex As Exception
                End Try
            Next

            m_colDeleteFreeTextIDs.Clear()
        End If

        If m_colNewFreeTextIDs IsNot Nothing Then
            m_colNewFreeTextIDs.Clear()
        End If
    End Sub

    Private Sub FlagFreeTextDeletion(ByVal objMask As clsMaskField)
        If objMask.Value1.Value Is Nothing Then
            Return
        End If

        If m_colDeleteFreeTextIDs Is Nothing Then
            m_colDeleteFreeTextIDs = New Hashtable
        End If

        m_colDeleteFreeTextIDs.Add(objMask.Field.DatabaseName, CInt(objMask.Value1.Value))
        objMask.Value1.Value = Nothing
    End Sub
#End Region

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objTable = Nothing
                m_objParent = Nothing

                If Not m_colMaskField Is Nothing Then
                    m_colMaskField.Dispose()
                    m_colMaskField = Nothing
                End If

                If Not m_colMaskOneToMany Is Nothing Then
                    m_colMaskOneToMany.Dispose()
                    m_colMaskOneToMany = Nothing
                End If

                If Not m_colMaskManyToMany Is Nothing Then
                    m_colMaskManyToMany.Dispose()
                    m_colMaskManyToMany = Nothing
                End If

                If Not m_colMaskBase Is Nothing Then
                    m_colMaskBase.Clear()
                    m_colMaskBase = Nothing
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
