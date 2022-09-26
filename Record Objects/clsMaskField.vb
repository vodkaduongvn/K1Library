#Region " File Information "

'==============================================================================
' This class is a field/UI Control combination for use on the Mask UI Page
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       12/07/2004  Implemented.
'==============================================================================

#End Region

#End Region

Public Class clsMaskField
    Inherits clsMaskBase

#Region " Members "

    Private m_objField As clsField

    'Used In Add, Modify, and Searches
    Private m_objValue1 As clsMaskFieldValue

    'Used In Range Search
    Private m_objValue2 As clsMaskFieldValue

    'Used in Range Search for Foreign Key Fields (In Clause)
    Private m_objMaskFieldLink As clsMaskFieldLink
    Private m_blnHasAutoFill As Boolean = False
    Private m_colMasks As clsMaskFieldDictionary
    Private m_blnMandatory As Boolean
    Private m_blnReadOnly As Boolean
    Private m_blnAllowFreeTextEntry As Boolean
    Private m_blnIsMultipleSequenceField As Boolean
    Private m_blnDeterminesMultipleSequence As Boolean
    Private m_blnAutonumberUpdated As Boolean = False
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal colMasks As clsMaskFieldDictionary, _
    ByVal objField As clsField, _
    ByVal eMaskType As clsTableMask.enumMaskType)
        MyBase.New(objDB, enumMaskObjectType.FIELD, eMaskType)
        m_objField = objField
        m_colMasks = colMasks
    End Sub

    Protected Sub New(ByVal objDB As clsDB, _
    ByVal colMasks As clsMaskFieldDictionary, _
    ByVal objField As clsField, ByVal strCaption As String, _
    ByVal blnIsVisible As Boolean, _
    ByVal eMaskType As clsTableMask.enumMaskType, _
    ByVal blnIsMultipleSequenceField As Boolean, _
    ByVal blnDeterminesMultipleSequence As Boolean, _
    ByVal blnAllowFreeTextEntry As Boolean, _
    ByVal blnMandatory As Boolean, _
    ByVal blnReadOnly As Boolean)
        MyBase.New(objDB, enumMaskObjectType.FIELD, eMaskType)
        m_objField = objField
        m_colMasks = colMasks
        m_strCaption = strCaption
        m_blnIsVisible = blnIsVisible
        m_blnIsMultipleSequenceField = blnIsMultipleSequenceField
        m_blnDeterminesMultipleSequence = blnDeterminesMultipleSequence
        m_blnAllowFreeTextEntry = blnAllowFreeTextEntry
        m_blnMandatory = blnMandatory
        m_blnReadOnly = blnReadOnly
    End Sub
#End Region

#Region " Properties "

    Public Property Field() As clsField
        Get
            Return m_objField
        End Get
        Set(ByVal Value As clsField)
            m_objField = Value
        End Set
    End Property

    Public Property Value1() As clsMaskFieldValue
        Get
            If m_objValue1 Is Nothing Then
                m_objValue1 = New clsMaskFieldValue(Me)
            End If
            Return m_objValue1
        End Get
        Set(ByVal Value As clsMaskFieldValue)
            m_objValue1 = Value
        End Set
    End Property

    Public Property Value2() As clsMaskFieldValue
        Get
            If m_objValue2 Is Nothing Then
                m_objValue2 = New clsMaskFieldValue(Me)
            End If
            Return m_objValue2
        End Get
        Set(ByVal Value As clsMaskFieldValue)
            m_objValue2 = Value
        End Set
    End Property

    Public Property HasAutoFill() As Boolean
        Get
            Return m_blnHasAutoFill
        End Get
        Set(ByVal Value As Boolean)
            m_blnHasAutoFill = Value
        End Set
    End Property

    Public ReadOnly Property IsRangeControl() As Boolean
        Get
            Return m_objField.IsRangeField
        End Get
    End Property

    Public Overrides ReadOnly Property HasTableAccess() As Boolean
        Get
            If m_objField.IsForeignKey Then
                Return (m_objDB.Profile.HasAccess( _
                    m_objField.FieldLink.IdentityTable.SecurityID) AndAlso _
                    m_objDB.Profile.LinkTables( _
                    CType(m_objField.FieldLink.IdentityTable.ID, String)) IsNot Nothing)
            Else
                Return False
            End If
        End Get
    End Property

    Public Property MaskFieldLink() As clsMaskFieldLink
        Get
            Return m_objMaskFieldLink
        End Get
        Set(ByVal value As clsMaskFieldLink)
            m_objMaskFieldLink = value
        End Set
    End Property

    Public ReadOnly Property MaskFieldCollection() As clsMaskFieldDictionary
        Get
            Return m_colMasks
        End Get
    End Property

    Public ReadOnly Property IsMultipleSequenceField() As Boolean
        Get
            Return m_blnIsMultipleSequenceField
        End Get
    End Property

    Public ReadOnly Property DeterminesMultipleSequence() As Boolean
        Get
            Return m_blnDeterminesMultipleSequence
        End Get
    End Property

    Public ReadOnly Property AllowFreeTextEntry() As Boolean
        Get
            Return m_blnAllowFreeTextEntry
        End Get
    End Property

    Public ReadOnly Property IsMandatory() As Boolean
        Get
            Return m_blnMandatory
        End Get
    End Property

    Public Property IsReadOnly() As Boolean
        Get
            Return m_blnReadOnly
        End Get
        Set(ByVal value As Boolean)
            m_blnReadOnly = value
        End Set
    End Property

    Public Property AutonumberUpdated() As Boolean
        Get
            Return m_blnAutonumberUpdated
        End Get
        Set(ByVal value As Boolean)
            m_blnAutonumberUpdated = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " CreateMaskCollection "

#Region " Simple (No Linked Info, No Captions, and No Autofills) "

    ''' <summary>
    ''' Creates a mask field collection and populates the fields if the ID is not null
    ''' </summary>
    Public Shared Function CreateMaskCollection(ByVal objTable As clsTable,
    Optional ByVal intID As Integer = clsDBConstants.cintNULL,
    Optional ByVal blnLoadLinkData As Boolean = False) As clsMaskFieldDictionary
        Dim colMasks As New clsMaskFieldDictionary

        'Create our collection of maskobjects if new instance or we have found a record
        For Each objField As clsField In objTable.Fields.Values
            colMasks.Add(New clsMaskField(objTable.Database,
                colMasks, objField, clsTableMask.enumMaskType.VIEW))
        Next

        colMasks.Table = objTable

        If Not intID = clsDBConstants.cintNULL Then
            Dim objDT As DataTable = objTable.Database.GetItem(objTable.DatabaseName, intID)
            If objDT.Rows.Count = 0 Then
                Throw New ApplicationException("You no longer have access to the selected record.")
            End If
            colMasks = LoadRecord(objDT, colMasks, blnLoadLinkData, clsTableMask.enumMaskType.VIEW)

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If
        End If

        Return colMasks
    End Function
#End Region

#Region " Full (Linked Info, Captions, and AutoFills) "

    ''' <summary>
    ''' Creates a mask field collection and populates the fields if the ID is not null
    ''' If this is an ADD Mask Type, fields values are populated using autofills
    ''' </summary>
    Public Shared Function CreateMaskCollection(ByVal objTable As clsTable,
                                                ByVal eMaskType As clsTableMask.enumMaskType,
                                                Optional ByVal intID As Integer = clsDBConstants.cintNULL,
                                                Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL,
                                                Optional ByVal objParent As clsDBObject = Nothing,
                                                Optional ByVal objParentMask As clsMaskBase = Nothing,
                                                Optional ByVal objDB As clsDB = Nothing) As clsMaskFieldDictionary
        'Try
        Dim objProfile As clsUserProfile
        If objDB Is Nothing OrElse objDB.Profile Is Nothing Then
            objProfile = objTable.Database.Profile
        Else
            objProfile = objDB.Profile
        End If

        Dim colMasks As New clsMaskFieldDictionary
        colMasks.Table = objTable

        Dim objDT As DataTable = Nothing

        If Not intID = clsDBConstants.cintNULL Then
            objDT = objTable.Database.GetItem(objTable.DatabaseName, intID)

            If objDT.Rows.Count = 0 Then
                objDT.Dispose()
                objDT = Nothing
                intID = clsDBConstants.cintNULL
            ElseIf objTable.TypeDependent Then
                intTypeID = CInt(clsDB.NullValue(objDT.Rows(0)(clsDBConstants.Fields.cTYPEID), clsDBConstants.cintNULL))
            End If
        End If

        'Create our collection of maskobjects if new instance or we have found a record
        For Each objField As clsField In objTable.Fields.Values
            Dim objTypeFieldInfo As clsTypeField = Nothing
            Dim strCaption As String = Nothing
            Dim blnVisible As Boolean = objField.IsVisible
            Dim blnMandatory As Boolean = objField.IsMandatory
            Dim blnReadOnly As Boolean = objField.IsReadOnly
            Dim blnIsMultipleSequenceField As Boolean = objField.IsMultipleSequenceField
            Dim blnDeterminesMultipleSequence As Boolean = objField.DeterminesMultipleSequence
            Dim blnAllowFreeTextEntry As Boolean = objField.AllowFreeTextEntry
            '2016-09-22 -- Peter & James -- Bug fix for #1600003206
            Dim colAccessRightField As New List(Of clsAccessRightField)
            Dim intSecurityID As Integer = objField.SecurityID

            If objTable.TypeDependent AndAlso Not intTypeID = clsDBConstants.cintNULL Then
                objTypeFieldInfo = objField.TypeFieldInfos(CType(intTypeID, String))
                '2016-09-22 -- Peter & James -- Bug fix for #1600003206
                For Each intSecurityGroupID As Integer In objTable.Database.Profile.LinkSecurityGroups.Values
                    If objTable.Database.Profile.TypeFields(CStr(objField.ID & "_" & intTypeID & "_" & intSecurityGroupID)) IsNot Nothing Then
                        colAccessRightField.Add(objTable.Database.Profile.TypeFields(CStr(objField.ID & "_" & intTypeID & "_" & intSecurityGroupID)))
                    End If
                Next
            End If

            If objTable.TypeDependent AndAlso Not objTypeFieldInfo Is Nothing Then
                strCaption = objTypeFieldInfo.CaptionText
                blnVisible = objTypeFieldInfo.IsVisible
                blnMandatory = objTypeFieldInfo.IsMandatory
                blnIsMultipleSequenceField = objTypeFieldInfo.IsMultipleSequenceField
                blnDeterminesMultipleSequence = objTypeFieldInfo.DeterminesMultipleSequence
                blnAllowFreeTextEntry = objTypeFieldInfo.AllowFreeTextEntry
                blnReadOnly = objTypeFieldInfo.IsReadOnly
                intSecurityID = objTypeFieldInfo.SecurityID

                If blnAllowFreeTextEntry Then
                    blnIsMultipleSequenceField = False
                    blnDeterminesMultipleSequence = False
                End If
            End If

            If objTable.TypeDependent Then
                blnReadOnly = blnReadOnly OrElse objField.DatabaseName.ToUpper = clsDBConstants.Fields.cTYPEID.ToUpper
            End If

            Dim blnIsVisible As Boolean = ( _
                blnVisible AndAlso _
                objProfile.HasAccess(intSecurityID) AndAlso _
                objProfile.LinkFields(objField.KeyID) IsNot Nothing)

            'check if securitygroup linked to readonly field

            '#1300002522 a fields inherit readonly status will override the securitygroupreadonlyfield table
            If Not blnReadOnly Then
                blnReadOnly = objProfile.LinkReadOnlyFields(objField.KeyID) IsNot Nothing
            End If

            '2016-09-22 -- Peter & James -- Bug fix for #1600003206
            If colAccessRightField.Count > 0 Then
                Dim blnVisibleForType As Boolean = False
                Dim blnReadOnlyForType As Boolean = True

                For Each objAccessRightField As clsAccessRightField In colAccessRightField
                    If Not blnVisibleForType Then
                        blnVisibleForType = objAccessRightField.Visible
                    End If
                    If blnReadOnlyForType And objAccessRightField.Visible Then
                        blnReadOnlyForType = objAccessRightField.IsReadOnly
                    End If
                Next

                blnIsVisible = blnIsVisible AndAlso blnVisibleForType
                blnReadOnly = blnReadOnly OrElse blnReadOnlyForType
            End If

            If blnIsVisible Then
                If strCaption Is Nothing Then strCaption = objField.CaptionText
            End If

            colMasks.Add(New clsMaskField(objTable.Database, colMasks, objField, _
                strCaption, blnIsVisible, eMaskType, blnIsMultipleSequenceField, _
                blnDeterminesMultipleSequence, blnAllowFreeTextEntry, blnMandatory, blnReadOnly))
        Next

        If Not intID = clsDBConstants.cintNULL Then
            LoadRecord(objDT, colMasks, True, eMaskType)
        ElseIf objTable.TypeDependent AndAlso _
        Not intTypeID = clsDBConstants.cintNULL Then
            Dim objMask As clsMaskField = colMasks(clsDBConstants.Fields.cTYPEID)

            objMask.Value1.InitializeValue(intTypeID)
            LoadLinkedData(objMask, objTable)
        End If

        'check if any of the fields have an autofill
        If eMaskType = clsTableMask.enumMaskType.ADD OrElse _
        eMaskType = clsTableMask.enumMaskType.MULTI_ADD OrElse _
        eMaskType = clsTableMask.enumMaskType.MODIFY OrElse _
        eMaskType = clsTableMask.enumMaskType.MULTI_MODIFY Then
            For Each objMask As clsMaskField In colMasks.Values
                GetAutoFills(objMask, objTable, objProfile, intTypeID, objParent, objParentMask, intID)
            Next
        End If

        Return colMasks
        'Catch ex As Exception
        '    Throw
        'End Try
    End Function

#End Region

#Region "Type"
    Public Function isInType(typeId As Integer) As Boolean
        Return Field.isInType(typeId)
    End Function
#End Region

#End Region

#Region " AutoFill Functions "

    Private Shared Sub GetAutoFills(ByVal objMask As clsMaskField, _
    ByVal objTable As clsTable, ByVal objProfile As clsUserProfile, _
    ByVal intTypeID As Integer, ByVal objParent As clsDBObject, _
    ByVal objParentMask As clsMaskBase, ByVal intID As Integer)
        Dim objTypeFieldInfo As clsTypeField = Nothing

        If objTable.TypeDependent AndAlso Not intTypeID = clsDBConstants.cintNULL Then
            objTypeFieldInfo = objMask.Field.TypeFieldInfos(CType(intTypeID, String))
        End If

        Dim objAutoFill As clsAutoFillInfo
        If objTable.TypeDependent AndAlso Not objTypeFieldInfo Is Nothing Then
            objAutoFill = objTypeFieldInfo.AutoFillInfo
        Else
            objAutoFill = objMask.Field.AutoFillInfo
        End If

        Dim objValue As Object = Nothing
        Dim strDisplay As String = Nothing
        Dim intLinkedSecurityID As Integer
        Dim objExtraValue As Object = Nothing

        If objAutoFill Is Nothing OrElse objAutoFill.FillType = clsDBConstants.enumAutoFillTypes.NOTHING Then
            If Not objParent Is Nothing AndAlso _
            Not objParentMask Is Nothing AndAlso TypeOf objParentMask Is clsMaskFieldLink AndAlso _
            objMask.Field.IsForeignKey AndAlso _
            objMask.Field.FieldLink.IdentityTable.DatabaseName = objParent.Table.ExternalID AndAlso _
            objMask.Field.FieldLink.ID = CType(objParentMask, clsMaskFieldLink).FieldLink.ID AndAlso _
            Not objParent.ID = clsDBConstants.cintNULL Then
                objValue = objParent.ID
                strDisplay = objParent.ExternalID
                intLinkedSecurityID = objParent.SecurityID
                If objMask.Field.FormatType = clsDBConstants.enumFormatType.FKeyExtraField Then
                    Dim objExtraDT As DataTable = objTable.Database.GetItem( _
                        objMask.Field.FieldLink.IdentityTable.DatabaseName, objParent.ID)
                    objExtraValue = clsMaskField.GetExtraFieldValue(objMask, objExtraDT.Rows(0))

                    '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                    If objExtraDT IsNot Nothing Then
                        objExtraDT.Dispose()
                        objExtraDT = Nothing
                    End If
                End If
            End If

            If Not (objTable.TypeDependent AndAlso _
            objMask.Field.DatabaseName.ToUpper = clsDBConstants.Fields.cTYPEID.ToUpper) AndAlso _
            Not intID = clsDBConstants.cintNULL AndAlso _
            objMask.MaskType = clsTableMask.enumMaskType.ADD AndAlso _
            (objMask.IsReadOnly AndAlso objMask.IsMandatory = False) Then
                'We shouldn't clone readonly fields with no autofill
                objMask.Value1.InitializeValue(Nothing, Nothing, clsDBConstants.cintNULL)
            End If

            ' 2014-02-27 -- Peter Melisi -- Autofill NewSpace with AllocatedSpace if required -- Bug fix for #1400002604
            '[Begin]
            If (objParent IsNot Nothing) AndAlso
                (objParent.Table.DatabaseName = clsDBConstants.Tables.cMETADATAPROFILE) AndAlso
                (objTable.DatabaseName = clsDBConstants.Tables.cMOVEMENT) AndAlso
                (objMask.Field.DatabaseName = clsDBConstants.Fields.Movement.cNEWSPACEID) Then

                Dim colMasks As clsMaskFieldDictionary = CreateMaskCollection(objParent.Table, objParent.ID, True)

                If (colMasks.GetMaskValue(clsDBConstants.Fields.MetadataProfile.cALLOCATEDSPACEID) IsNot Nothing) AndAlso
                    (InStr(CStr(colMasks.GetMaskValue(clsDBConstants.Fields.MetadataProfile.cCURRENTLOCATION)), "Space") = 0) Then
                    objValue = colMasks.GetMaskValue(clsDBConstants.Fields.MetadataProfile.cALLOCATEDSPACEID)
                    strDisplay = colMasks(clsDBConstants.Fields.MetadataProfile.cALLOCATEDSPACEID).Value1.Display
                    intLinkedSecurityID = objParent.SecurityID
                End If
            End If
            '[End]
        Else
            If objTable.TypeDependent AndAlso _
            objMask.Field.DatabaseName.ToUpper = clsDBConstants.Fields.cTYPEID.ToUpper Then
                Return
            End If

            If objMask.MaskType = clsTableMask.enumMaskType.MODIFY OrElse _
            objMask.MaskType = clsTableMask.enumMaskType.MULTI_MODIFY OrElse _
            objAutoFill.FillType = clsDBConstants.enumAutoFillTypes.AUTO_NUMBER_FORMAT Then
                If objAutoFill.FillType = clsDBConstants.enumAutoFillTypes.AUTO_NUMBER_FORMAT AndAlso _
                Not objAutoFill.AutoNumberFormat Is Nothing Then
                    objMask.Value1.AutoNumber = objAutoFill.AutoNumberFormat
                End If
            Else
                Select Case objAutoFill.FillType
                    Case clsDBConstants.enumAutoFillTypes.LAST_MOVEMENT
                        If objParent IsNot Nothing AndAlso _
                        objParent.Table.DatabaseName = clsDBConstants.Tables.cMETADATAPROFILE AndAlso _
                        objMask.Field.IsForeignKey AndAlso _
                        objMask.Field.FieldLink.IdentityTable.DatabaseName = clsDBConstants.Tables.cMOVEMENT Then
                            Dim objDT As DataTable = objTable.Database.GetDataTableByField( _
                                clsDBConstants.Tables.cMOVEMENT, clsDBConstants.Fields.Movement.cMETADATAPROFILEID, objParent.ID)

                            Dim objMovement As clsDBObjBase
                            If objDT IsNot Nothing AndAlso objDT.Rows.Count > 0 Then
                                objDT.DefaultView.Sort = clsDBConstants.Fields.Movement.cMOVEDDATE & " DESC"

                                Dim objDRV As DataRowView = objDT.DefaultView.Item(0)

                                objMovement = New clsDBObject(objDRV.Row, objTable.Database)

                                objValue = objMovement.ID
                                strDisplay = objMovement.ExternalID
                                intLinkedSecurityID = objMovement.SecurityID

                                objMask.HasAutoFill = True
                            End If

                            If objDT IsNot Nothing Then objDT.Dispose()
                        End If

                    Case clsDBConstants.enumAutoFillTypes.LOGGED_IN_PERSON
                        objValue = objProfile.PersonID
                        strDisplay = objProfile.Person.ExternalID
                        intLinkedSecurityID = objProfile.SecurityID
                        objMask.HasAutoFill = True

                    Case clsDBConstants.enumAutoFillTypes.SELECTED_ITEM_EXTERNALID
                        If Not objParent Is Nothing Then
                            objValue = objParent.ExternalID
                        End If

                    Case clsDBConstants.enumAutoFillTypes.TODAY_DATE
                        Dim dtDate As Date = objMask.MaskFieldCollection.CurrentServerDate

                        If objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME Then 'localtimedate
                            dtDate = objMask.Database.Profile.ToLocalTime(dtDate) 'localdatetime
                        End If

                        If objMask.Field.DateType = clsDBConstants.enumDateTypes.DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Then 'localtimedate
                            objValue = dtDate.Date
                        Else
                            objValue = dtDate
                        End If
                        objMask.HasAutoFill = True

                    Case clsDBConstants.enumAutoFillTypes.WEEKS_FROM_DATE
                        Try
                            Dim dtDate As Date = objMask.MaskFieldCollection.CurrentServerDate

                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME Then 'localtimedate
                                dtDate = objMask.Database.Profile.ToLocalTime(dtDate) 'localdatetime
                            End If

                            Dim intWeeks As Integer = CType(objAutoFill.FillValue, Integer)
                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Then 'localtimedate
                                objValue = dtDate.Date.AddDays(7 * intWeeks)
                            Else
                                objValue = dtDate.AddDays(7 * intWeeks)
                            End If
                            objMask.HasAutoFill = True
                        Catch ex As Exception
                        End Try

                    Case clsDBConstants.enumAutoFillTypes.MONTHS_FROM_DATE
                        Try
                            Dim dtDate As Date = objMask.MaskFieldCollection.CurrentServerDate

                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME Then 'localtimedate
                                dtDate = objMask.Database.Profile.ToLocalTime(dtDate) 'localdatetime
                            End If

                            Dim intMonths As Integer = CType(objAutoFill.FillValue, Integer)
                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Then 'localtimedate
                                objValue = dtDate.Date.AddMonths(intMonths)
                            Else
                                objValue = dtDate.AddMonths(intMonths)
                            End If
                            objMask.HasAutoFill = True
                        Catch ex As Exception
                        End Try

                    Case clsDBConstants.enumAutoFillTypes.DAYS_FROM_DATE
                        Try
                            Dim dtDate As Date = objMask.MaskFieldCollection.CurrentServerDate

                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME Then
                                dtDate = objMask.Database.Profile.ToLocalTime(dtDate) 'localdatetime
                            End If

                            Dim intDays As Integer = CType(objAutoFill.FillValue, Integer)
                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Then 'localtimedate
                                objValue = dtDate.Date.AddDays(intDays)
                            Else
                                objValue = dtDate.AddDays(intDays)
                            End If
                            objMask.HasAutoFill = True
                        Catch ex As Exception
                        End Try

                    Case clsDBConstants.enumAutoFillTypes.YEARS_FROM_DATE
                        Try
                            Dim dtDate As Date = objMask.MaskFieldCollection.CurrentServerDate

                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME Then 'localtimedate
                                dtDate = objMask.Database.Profile.ToLocalTime(dtDate) 'localdatetime
                            End If

                            Dim intYears As Integer = CType(objAutoFill.FillValue, Integer)
                            If objMask.Field.DateType = clsDBConstants.enumDateTypes.DATE_ONLY Or objMask.Field.DateType = clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY Then 'localtimedate
                                objValue = dtDate.Date.AddYears(intYears)
                            Else
                                objValue = dtDate.AddYears(intYears)
                            End If
                            objMask.HasAutoFill = True
                        Catch ex As Exception
                        End Try

                    Case clsDBConstants.enumAutoFillTypes.SELECTED_ITEM_TABLE
                        If Not objParent Is Nothing AndAlso _
                        Not objParent.Table Is Nothing Then
                            objValue = objParent.Table.ID
                            strDisplay = objParent.Table.DatabaseName
                            intLinkedSecurityID = objParent.Table.SecurityID
                        End If

                    Case clsDBConstants.enumAutoFillTypes.LOGGED_PERSON_SECURITYID
                        If Not objProfile.NominatedSecurity Is Nothing Then
                            objValue = objProfile.NominatedSecurity.ID
                            strDisplay = objProfile.NominatedSecurity.ExternalID
                            intLinkedSecurityID = objProfile.NominatedSecurity.SecurityID
                        End If
                        objMask.HasAutoFill = True

                    Case clsDBConstants.enumAutoFillTypes.LOGGED_IN_USER_PROFILE
                        objValue = objProfile.ID
                        strDisplay = objProfile.ExternalID
                        intLinkedSecurityID = objProfile.SecurityID
                        objMask.HasAutoFill = True

                    Case clsDBConstants.enumAutoFillTypes.USER_VALUE
                        If Not objMask.Field.IsForeignKey Then
                            objValue = objAutoFill.FillValue
                            objMask.HasAutoFill = True
                        Else
                            Try
                                If objMask.AllowFreeTextEntry Then
                                    objMask.Value1.FreeText = objAutoFill.FillValue
                                Else
                                    Dim objDataTable As DataTable = objTable.Database.GetItem( _
                                        objMask.Field.FieldLink.IdentityTable.DatabaseName, CType(objAutoFill.FillValue, Integer))

                                    If Not objDataTable Is Nothing AndAlso objDataTable.Rows.Count = 1 Then
                                        strDisplay = CType(clsDB_Direct.NullValue( _
                                            objDataTable.Rows(0).Item(clsDBConstants.Fields.cEXTERNALID), ""), String)
                                        intLinkedSecurityID = CType(clsDB_Direct.NullValue( _
                                            objDataTable.Rows(0).Item(clsDBConstants.Fields.cSECURITYID), _
                                            clsDBConstants.cintNULL), Integer)
                                        objValue = objAutoFill.FillValue
                                        objExtraValue = clsMaskField.GetExtraFieldValue(objMask, objDataTable.Rows(0))
                                    End If

                                    '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                                    If objDataTable IsNot Nothing Then
                                        objDataTable.Dispose()
                                        objDataTable = Nothing
                                    End If
                                End If
                                objMask.HasAutoFill = True
                            Catch ex As Exception
                            End Try
                        End If
                End Select
            End If
        End If

        Try
            If Not objValue Is Nothing Then
                Select Case objMask.Field.DataType
                    Case SqlDbType.BigInt
                        objMask.Value1.InitializeValue(CType(objValue, Int64))

                    Case SqlDbType.Bit
                        objMask.Value1.InitializeValue(CType(objValue, Boolean))

                    Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.VarChar
                        Dim strText As String = CType(objValue, String)

                        If strText.Length > objMask.Field.Length Then
                            strText = strText.Substring(0, objMask.Field.Length)
                            objMask.Value1.InitializeValue(CType(strText, String))
                        Else
                            objMask.Value1.InitializeValue(CType(strText, String))
                        End If

                    Case SqlDbType.NText, SqlDbType.Text
                        objMask.Value1.InitializeValue(CType(objValue, String))

                    Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                        objMask.Value1.InitializeValue(CType(objValue, Date))

                    Case SqlDbType.Decimal, SqlDbType.Float, SqlDbType.Real, _
                    SqlDbType.Money, SqlDbType.SmallMoney
                        objMask.Value1.InitializeValue(CType(objValue, Double))

                    Case SqlDbType.Int
                        objMask.Value1.InitializeValue(CType(objValue, Integer))

                    Case SqlDbType.SmallInt
                        objMask.Value1.InitializeValue(CType(objValue, Int16))

                    Case SqlDbType.TinyInt
                        objMask.Value1.InitializeValue(CType(objValue, Byte))

                End Select

                If Not strDisplay Is Nothing AndAlso objMask.Field.IsForeignKey Then
                    objMask.Value1.Display = strDisplay
                    objMask.Value1.ObjectSecurityID = intLinkedSecurityID
                    objMask.Value1.ExtraFieldValue = objExtraValue
                End If

                objMask.HasAutoFill = True
            Else
                objMask.HasAutoFill = False
            End If

        Catch ex As Exception
            objMask.Value1.InitializeValue(Nothing)
        End Try
    End Sub

#End Region

#Region " Load Record "

    ''' <summary>
    ''' Fills the Mask Collection from a Record Using the ID provided
    ''' </summary>
    Private Shared Function LoadRecord(ByVal objDT As DataTable, _
    ByVal colMasks As clsMaskFieldDictionary, ByVal blnLoadLinkData As Boolean, _
    ByVal eMaskType As clsTableMask.enumMaskType) As clsMaskFieldDictionary
        Try
            Dim objTable As clsTable = colMasks.Table

            'Load the values from the database
            For Each objMask As clsMaskField In colMasks.Values
                If Not ((eMaskType = clsTableMask.enumMaskType.ADD OrElse _
                eMaskType = clsTableMask.enumMaskType.MULTI_ADD) AndAlso _
                objMask.Field.IsIdentityField) Then
                    objMask.Value1.InitializeValue(clsDB_Direct.DataRowValue( _
                        objDT.Rows(0), objMask.Field.DatabaseName, Nothing))
                End If

                'load linked record information
                If blnLoadLinkData Then
                    LoadLinkedData(objMask, objTable)
                End If

                If objMask.Field.DataType = SqlDbType.Image Then
                    objMask.Value1.InitializeValue(1)
                End If
            Next

            Return colMasks
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' If the field is a foreign key, this function will load associated data 
    ''' (linked item's ExternalID and security)
    ''' </summary>
    Public Shared Sub LoadLinkedData(ByVal objMask As clsMaskField, ByVal objTable As clsTable)

        If objMask.Field.IsForeignKey AndAlso
            objMask.Value1.Value IsNot Nothing Then

            Dim objLinkDT As DataTable = objTable.Database.GetItem(objMask.Field.FieldLink.IdentityTable.DatabaseName,
                                                                   CType(objMask.Value1.Value, Integer))

            LoadLinkedData(objMask, objLinkDT)

            If objMask.AllowFreeTextEntry Then
                objMask.Value1.FreeText = objMask.Value1.Display
            End If
        End If

    End Sub

    Public Shared Sub LoadLinkedData(ByVal objMask As clsMaskField, ByVal objDT As DataTable)
        If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
            objMask.Value1.Display = CType(clsDB_Direct.NullValue( _
                objDT.Rows(0).Item(clsDBConstants.Fields.cEXTERNALID), _
                clsDBConstants.cstrNULL), String)
            objMask.Value1.ObjectSecurityID = CType(clsDB_Direct.NullValue( _
                objDT.Rows(0).Item(clsDBConstants.Fields.cSECURITYID), _
                clsDBConstants.cintNULL), Integer)
            objMask.Value1.ExtraFieldValue = GetExtraFieldValue(objMask, objDT.Rows(0))
        End If
    End Sub

    Public Shared Function GetExtraFieldValue(ByVal objMask As clsMaskField, ByVal objDR As DataRow) As String
        Dim objExtraValue As Object = Nothing

        Try
            If objMask.Field.FormatType = clsDBConstants.enumFormatType.FKeyExtraField Then
                Dim objField As clsField = objMask.Database.SysInfo.Fields( _
                    objMask.Field.FieldLink.IdentityTable.ID & "_" & objMask.Field.FormatString)
                If objField IsNot Nothing AndAlso _
                objField.HasAccess AndAlso objField.Table.HasAccess Then
                    objExtraValue = clsDB_Direct.NullValue( _
                        objDR.Item(objMask.Field.FormatString), Nothing)

                    If objField.IsForeignKey AndAlso objExtraValue IsNot Nothing Then
                        'Get the externalID of the linked record
                        Dim objLinkDT As DataTable = objField.Database.GetItem( _
                            objField.FieldLink.IdentityTable.DatabaseName, _
                            CType(objExtraValue, Integer))

                        If objLinkDT IsNot Nothing AndAlso objLinkDT.Rows.Count = 1 Then
                            objExtraValue = CType(clsDB_Direct.NullValue( _
                                objLinkDT.Rows(0).Item(clsDBConstants.Fields.cEXTERNALID), _
                                clsDBConstants.cstrNULL), String)
                        End If

                        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                        If objLinkDT IsNot Nothing Then
                            objLinkDT.Dispose()
                            objLinkDT = Nothing
                        End If
                    End If

                    If objField.IsDateType Then
                        Select Case objField.DateType
                            Case clsDBConstants.enumDateTypes.DATE_AND_TIME, clsDBConstants.enumDateTypes.LOCAL_DATE_AND_TIME 'localtimedate
                                objExtraValue = CDate(objExtraValue).ToString("dd MMM yyyy hh:mm:ss tt")

                            Case clsDBConstants.enumDateTypes.DATE_ONLY, clsDBConstants.enumDateTypes.LOCAL_DATE_ONLY 'localtimedate
                                objExtraValue = CDate(objExtraValue).ToString("dd MMM yyyy")

                            Case clsDBConstants.enumDateTypes.TIME_ONLY
                                objExtraValue = CDate(objExtraValue).ToString("hh:mm:ss tt")
                        End Select
                    End If
                End If
            End If
        Catch ex As Exception
            objExtraValue = Nothing
        End Try

        Return CType(objExtraValue, String)
    End Function
#End Region

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeObject()
        m_objField = Nothing
        m_colMasks = Nothing

        If Not m_objValue1 Is Nothing Then
            m_objValue1.Dispose()
            m_objValue1 = Nothing
        End If

        If Not m_objValue2 Is Nothing Then
            m_objValue2.Dispose()
            m_objValue2 = Nothing
        End If

        If Not m_objMaskFieldLink Is Nothing Then
            m_objMaskFieldLink.Dispose()
            m_objMaskFieldLink = Nothing
        End If
    End Sub
#End Region

End Class
