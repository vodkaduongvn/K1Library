

'[Naing] 07/01/2014 Bug fix: 1400002579
' Made sure that when making changes to Audit table's fields, that auditing is by passed. This is only implemented for non type dependent fields of the audit table.
''' <summary>
''' Used to create or update a field using the DRM
''' </summary>
Public Class clsDRMField
    Inherits clsDRMBase

#Region " Members "

    '-- General Values
    Private m_objTable As clsTable
    Private m_strDatabaseName As String = clsDBConstants.cstrNULL
    Private m_eDataType As SqlDbType
    Private m_intLength As Integer = clsDBConstants.cintNULL
    Private m_blnIsMandatory As Boolean = False
    Private m_strCaption As String = clsDBConstants.cstrNULL
    Private m_eDateType As clsDBConstants.enumDateTypes = clsDBConstants.enumDateTypes.NOTHING
    Private m_intCaptionID As Integer = clsDBConstants.cintNULL
    Private m_blnIsEncrypted As Boolean = False
    Private m_blnIsVisible As Boolean = True
    Private m_intPrecision As Integer = clsDBConstants.cintNULL
    Private m_intScale As Integer = clsDBConstants.cintNULL
    Private m_blnReadOnly As Boolean = False
    Private m_intSortOrder As Integer = 5
    Private m_blnMultiLine As Boolean = False
    Private m_intNumberOfLines As Integer = 4
    Private m_blnIsExpanded As Boolean = False
    Private m_eFormat As clsDBConstants.enumFormatType = clsDBConstants.enumFormatType.None
    Private m_strFormatString As String

    '-- Autonumber Values
    Private m_eAutoFillType As clsDBConstants.enumAutoFillTypes = clsDBConstants.enumAutoFillTypes.NOTHING
    Private m_strAutoFillValue As String = clsDBConstants.cstrNULL
    Private m_intAutoNumberFormatID As Integer = clsDBConstants.cintNULL

    '-- System Values
    Private m_blnNullable As Boolean = True
    Private m_strDefaultValue As String = clsDBConstants.cstrNULL
    Private m_blnIsSystemEssential As Boolean = False
    Private m_blnIsSystemLocked As Boolean = False
    Private m_blnDisplayAsDropDown As Boolean = False
    Private m_blnIsFKeyExpanded As Boolean = False

    '-- UI Column Values
    'Private m_blnIsColumn As Boolean = False
    'Private m_intColumnHeadingCaptionID As Integer = clsDBConstants.cintNULL
    'Private m_strColHeading As String = clsDBConstants.cstrNULL
    'Private m_intColWidth As Integer = clsDBConstants.cintNULL
    'Private m_blnIsWidthPercentage As Boolean = True
    'TODO: allow width to be pixel or percentage

    '-- FK Values
    Private m_blnIsForeignKey As Boolean = False
    Private m_strForeignTable As String = clsDBConstants.cstrNULL
    Private m_strFieldLinkCaption As String = clsDBConstants.cstrNULL
    Private m_blnFieldLinkVisible As Boolean = False
    Private m_intFieldLinkSortOrder As Integer = 6

    '-- Multiple Sequence Fields
    Private m_blnIsMultipleSequenceField As Boolean = False
    Private m_blnDeterminesMultipleSequence As Boolean = False
    Private m_blnAllowFreeTextEntry As Boolean = False

    '-- Index Values
    'Private m_blnIsIndexed As Boolean = False
    'Private m_blnIsUniqueIndex As Boolean = False
    'Private m_blnIsClusteredIndex As Boolean = False
    Private m_colSecGroupIDs As FrameworkCollections.K1Collection(Of Integer)
#End Region

#Region " Enumerations "

    Public Enum enumStandardField
        ID = 1
        EXTERNALID = 2
        SECURITYID = 3
        TYPEID = 4
    End Enum
#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new field
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal objTable As clsTable, _
    ByVal strField As String, _
    ByVal eDataType As SqlDbType, _
    ByVal strCaption As String, _
    ByVal intSecurityID As Integer)
        MyBase.New(objDB, objTable.DatabaseName & "." & strField, intSecurityID, clsDBConstants.cintNULL)
        m_objTable = objTable
        m_strDatabaseName = strField
        m_eDataType = eDataType
        m_strCaption = strCaption
    End Sub
#End Region

#Region " New Standard Field (ID, ExternalID, TypeID, SecurityID) "

    ''' <summary>
    ''' Creates a new standard field (ID, ExternalID, TypeID, SecurityID)
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal objTable As clsTable, _
    ByVal eFieldType As enumStandardField, _
    ByVal intSecurityID As Integer)
        MyBase.New(objDB, clsDBConstants.cstrNULL, intSecurityID, clsDBConstants.cintNULL)
        m_objTable = objTable

        Select Case eFieldType
            Case enumStandardField.ID
                m_strDatabaseName = clsDBConstants.Fields.cID
                m_eDataType = SqlDbType.Int
                m_blnIsVisible = False
                m_blnReadOnly = True
                m_intSortOrder = 1
                m_blnIsSystemLocked = True
                m_blnIsSystemEssential = True
                m_blnNullable = False

            Case enumStandardField.EXTERNALID
                m_strDatabaseName = clsDBConstants.Fields.cEXTERNALID
                m_eDataType = SqlDbType.NVarChar
                m_intLength = 100
                m_intSortOrder = 2
                m_blnIsSystemEssential = True
                m_blnIsMandatory = True
                m_blnNullable = False

            Case enumStandardField.SECURITYID
                m_strDatabaseName = clsDBConstants.Fields.cSECURITYID
                m_eDataType = SqlDbType.Int
                m_blnIsMandatory = True
                m_intSortOrder = 100
                m_eAutoFillType = clsDBConstants.enumAutoFillTypes.LOGGED_PERSON_SECURITYID
                m_blnIsSystemEssential = True
                m_blnNullable = False
                m_blnDisplayAsDropDown = True
                MakeForeignKey(clsDBConstants.Tables.cSECURITY, False)

            Case enumStandardField.TYPEID
                m_strDatabaseName = clsDBConstants.Fields.cTYPEID
                m_eDataType = SqlDbType.Int
                m_intSortOrder = 3
                m_blnIsSystemEssential = True
                m_blnDisplayAsDropDown = True
                MakeForeignKey(clsDBConstants.Tables.cTYPE, False)

        End Select

        m_strCaption = m_strDatabaseName
        m_strExternalID = objTable.DatabaseName & "." & m_strDatabaseName
    End Sub
#End Region

#Region " New Linked Field "

    ''' <summary>
    ''' Creates a new linked field
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objTable As clsTable, ByVal strFieldName As String, _
    ByVal strIdentityTable As String, ByVal intSecurityID As Integer, ByVal blnLinkVisible As Boolean, _
    ByVal strFieldLinkCaption As String, ByVal intFieldLinkSortOrder As Integer, ByVal blnNullable As Boolean)
        MyBase.New(objDB, strFieldName, intSecurityID, clsDBConstants.cintNULL)
        m_objTable = objTable

        m_strDatabaseName = strFieldName
        m_eDataType = SqlDbType.Int
        m_blnIsMandatory = True
        m_strCaption = strFieldName
        m_intSortOrder = clsDBConstants.cintNULL
        m_blnIsSystemEssential = True
        m_blnNullable = blnNullable
        MakeForeignKey(strIdentityTable, blnLinkVisible, strFieldLinkCaption, intFieldLinkSortOrder)

        If m_objTable IsNot Nothing Then
            m_strExternalID = objTable.DatabaseName & "." & m_strDatabaseName
        End If
    End Sub
#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRM Field from an existing field database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objField As clsField)
        MyBase.New(objDB, objField)
        m_objTable = objField.Table
        m_strDatabaseName = objField.DatabaseName
        m_eDataType = objField.DataType
        m_intLength = objField.Length
        m_blnIsMandatory = objField.IsMandatory

        If objField.Caption IsNot Nothing Then
            m_strCaption = objField.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
            m_intCaptionID = objField.Caption.ID
        End If

        m_eDateType = objField.DateType
        m_blnIsEncrypted = objField.IsEncrypted
        m_blnIsVisible = objField.IsVisible
        m_intPrecision = objField.Length
        m_intScale = objField.Scale
        m_blnReadOnly = objField.IsReadOnly
        m_intSortOrder = objField.SortOrder
        m_blnMultiLine = objField.IsMultiLine
        m_blnIsExpanded = objField.IsExpanded
        m_intNumberOfLines = objField.NumberOfLines

        If Not objField.AutoFillInfo Is Nothing Then
            m_eAutoFillType = objField.AutoFillInfo.FillType
            m_strAutoFillValue = objField.AutoFillInfo.FillValue
            If Not objField.AutoFillInfo.AutoNumberFormat Is Nothing Then
                m_intAutoNumberFormatID = objField.AutoFillInfo.AutoNumberFormat.ID
            End If
        End If

        m_blnNullable = objField.IsNullable
        m_blnIsSystemEssential = objField.IsSystemEssential
        m_blnIsSystemLocked = objField.IsSystemLocked

        'If objField.IsColumn Then
        '    m_blnIsColumn = objField.IsColumn
        '    If objField.ColumnHeadingCaption IsNot Nothing Then
        '        m_strColHeading = objField.ColumnHeadingCaption.GetString( _
        '            objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
        '    End If
        '    m_intColWidth = objField.ColumnWidth
        '    If Not objField.ColumnHeadingCaption Is Nothing Then
        '        m_intColumnHeadingCaptionID = objField.ColumnHeadingCaption.ID
        '    End If
        'End If

        If Not objField.FieldLink Is Nothing Then
            m_blnIsForeignKey = True
            m_strForeignTable = objField.FieldLink.IdentityTable.DatabaseName
            m_strFieldLinkCaption = objField.FieldLink.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
            m_blnFieldLinkVisible = objField.FieldLink.IsVisible
            m_intFieldLinkSortOrder = objField.FieldLink.SortOrder
            m_blnDisplayAsDropDown = objField.FieldLink.DisplayAsDropDown
            m_blnIsFKeyExpanded = objField.FieldLink.IsExpanded
        End If

        m_blnIsMultipleSequenceField = objField.IsMultipleSequenceField
        m_blnDeterminesMultipleSequence = objField.DeterminesMultipleSequence
        m_blnAllowFreeTextEntry = objField.AllowFreeTextEntry

        m_eFormat = objField.FormatType
        m_strFormatString = objField.FormatString
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property Table() As clsTable
        Get
            Return m_objTable
        End Get
        Set(ByVal value As clsTable)
            m_objTable = value
            m_strExternalID = value.DatabaseName & "." & m_strDatabaseName
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            Return m_strDatabaseName
        End Get
        Set(ByVal value As String)
            m_strDatabaseName = value
        End Set
    End Property

    Public Property DataType() As SqlDbType
        Get
            Return m_eDataType
        End Get
        Set(ByVal value As SqlDbType)
            m_eDataType = value
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Return m_intLength
        End Get
        Set(ByVal value As Integer)
            m_intLength = value
        End Set
    End Property

    Public Property IsMandatory() As Boolean
        Get
            Return m_blnIsMandatory
        End Get
        Set(ByVal value As Boolean)
            m_blnIsMandatory = value
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

    Public Property DateType() As clsDBConstants.enumDateTypes
        Get
            Return m_eDateType
        End Get
        Set(ByVal value As clsDBConstants.enumDateTypes)
            m_eDateType = value
        End Set
    End Property

    Public Property IsEncrypted() As Boolean
        Get
            Return m_blnIsEncrypted
        End Get
        Set(ByVal value As Boolean)
            m_blnIsEncrypted = value
        End Set
    End Property

    Public Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
        Set(ByVal value As Boolean)
            m_blnIsVisible = value
        End Set
    End Property

    Public Property IsExpanded() As Boolean
        Get
            Return m_blnIsExpanded
        End Get
        Set(ByVal value As Boolean)
            m_blnIsExpanded = value
        End Set
    End Property

    Public Property Precision() As Integer
        Get
            Return m_intPrecision
        End Get
        Set(ByVal value As Integer)
            m_intPrecision = value
        End Set
    End Property

    Public Property Scale() As Integer
        Get
            Return m_intScale
        End Get
        Set(ByVal value As Integer)
            m_intScale = value
        End Set
    End Property

    Public Property IsReadOnly() As Boolean
        Get
            Return m_blnReadOnly
        End Get
        Set(ByVal value As Boolean)
            m_blnReadOnly = value
        End Set
    End Property

    Public Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
        Set(ByVal value As Integer)
            m_intSortOrder = value
        End Set
    End Property

    Public Property IsMultiLine() As Boolean
        Get
            Return m_blnMultiLine
        End Get
        Set(ByVal value As Boolean)
            m_blnMultiLine = value
        End Set
    End Property

    Public Property NumberOfLines() As Integer
        Get
            Return m_intNumberOfLines
        End Get
        Set(ByVal value As Integer)
            m_intNumberOfLines = value
        End Set
    End Property

    Public Property AutoFillType() As clsDBConstants.enumAutoFillTypes
        Get
            Return m_eAutoFillType
        End Get
        Set(ByVal value As clsDBConstants.enumAutoFillTypes)
            m_eAutoFillType = value
        End Set
    End Property

    Public Property AutoFillValue() As String
        Get
            Return m_strAutoFillValue
        End Get
        Set(ByVal value As String)
            m_strAutoFillValue = value
        End Set
    End Property

    Public Property AutoNumberFormatID() As Integer
        Get
            Return m_intAutoNumberFormatID
        End Get
        Set(ByVal value As Integer)
            m_intAutoNumberFormatID = value
        End Set
    End Property

    Public Property Nullable() As Boolean
        Get
            Return m_blnNullable
        End Get
        Set(ByVal value As Boolean)
            m_blnNullable = value
        End Set
    End Property

    Public Property DefaultValue() As String
        Get
            Return m_strDefaultValue
        End Get
        Set(ByVal value As String)
            m_strDefaultValue = value
        End Set
    End Property

    Public Property DisplayAsDropDown() As Boolean
        Get
            Return m_blnDisplayAsDropDown
        End Get
        Set(ByVal value As Boolean)
            m_blnDisplayAsDropDown = value
        End Set
    End Property

    Public Property IsFKeyExpanded() As Boolean
        Get
            Return m_blnIsFKeyExpanded
        End Get
        Set(ByVal value As Boolean)
            m_blnIsFKeyExpanded = value
        End Set
    End Property

    Public Property IsSystemEssential() As Boolean
        Get
            Return m_blnIsSystemEssential
        End Get
        Set(ByVal value As Boolean)
            m_blnIsSystemEssential = value
        End Set
    End Property

    Public Property IsSystemLocked() As Boolean
        Get
            Return m_blnIsSystemLocked
        End Get
        Set(ByVal value As Boolean)
            m_blnIsSystemLocked = value
        End Set
    End Property

    'Public Property IsColumn() As Boolean
    '    Get
    '        Return m_blnIsColumn
    '    End Get
    '    Set(ByVal value As Boolean)
    '        m_blnIsColumn = value
    '    End Set
    'End Property

    'Public Property ColHeading() As String
    '    Get
    '        Return m_strColHeading
    '    End Get
    '    Set(ByVal value As String)
    '        m_strColHeading = value
    '    End Set
    'End Property

    'Public Property ColWidth() As Integer
    '    Get
    '        Return m_intColWidth
    '    End Get
    '    Set(ByVal value As Integer)
    '        m_intColWidth = value
    '    End Set
    'End Property

    'Public Property IsWidthPercentage() As Boolean
    '    Get
    '        Return m_blnIsWidthPercentage
    '    End Get
    '    Set(ByVal value As Boolean)
    '        m_blnIsWidthPercentage = value
    '    End Set
    'End Property

    Public Property IsForeignKey() As Boolean
        Get
            Return m_blnIsForeignKey
        End Get
        Set(ByVal value As Boolean)
            m_blnIsForeignKey = value
        End Set
    End Property

    Public Property ForeignTable() As String
        Get
            Return m_strForeignTable
        End Get
        Set(ByVal value As String)
            m_strForeignTable = value
        End Set
    End Property

    Public Property FieldLinkCaption() As String
        Get
            Return m_strFieldLinkCaption
        End Get
        Set(ByVal value As String)
            m_strFieldLinkCaption = value
        End Set
    End Property

    Public Property FieldLinkVisible() As Boolean
        Get
            Return m_blnFieldLinkVisible
        End Get
        Set(ByVal value As Boolean)
            m_blnFieldLinkVisible = value
        End Set
    End Property

    Public Property FieldLinkSortOrder() As Integer
        Get
            Return m_intFieldLinkSortOrder
        End Get
        Set(ByVal value As Integer)
            m_intFieldLinkSortOrder = value
        End Set
    End Property

    Public Property Field() As clsField
        Get
            Return CType(m_objDBObj, clsField)
        End Get
        Set(ByVal value As clsField)
            m_objDBObj = value
        End Set
    End Property

    Public Property IsMultipleSequenceField() As Boolean
        Get
            Return m_blnIsMultipleSequenceField
        End Get
        Set(ByVal value As Boolean)
            m_blnIsMultipleSequenceField = value
        End Set
    End Property

    Public Property DeterminesMultipleSequence() As Boolean
        Get
            Return m_blnDeterminesMultipleSequence
        End Get
        Set(ByVal value As Boolean)
            m_blnDeterminesMultipleSequence = value
        End Set
    End Property

    Public Property AllowFreeTextEntry() As Boolean
        Get
            Return m_blnAllowFreeTextEntry
        End Get
        Set(ByVal value As Boolean)
            m_blnAllowFreeTextEntry = value
        End Set
    End Property

    Public Property FormatType() As clsDBConstants.enumFormatType
        Get
            Return m_eFormat
        End Get
        Set(ByVal value As clsDBConstants.enumFormatType)
            m_eFormat = value
        End Set
    End Property

    Public Property FormatString() As String
        Get
            Return m_strFormatString
        End Get
        Set(ByVal value As String)
            m_strFormatString = value
        End Set
    End Property

    Public Property SecurityGroupIDs() As FrameworkCollections.K1Collection(Of Integer)
        Get
            Return m_colSecGroupIDs
        End Get
        Set(ByVal value As FrameworkCollections.K1Collection(Of Integer))
            m_colSecGroupIDs = value
        End Set
    End Property


    Private ReadOnly Property IsAuditTableField() As Boolean
        Get
            '[Naing] Let's see if we need to create the audit trail record for this operation
            'Must ignore auditing when modifying and removing fields from Audit Table
            Return m_objTable.IsThisMe(clsDBConstants.Tables.cAUDITTRAIL)
        End Get
    End Property

#End Region

#Region " Methods "

#Region " Make System Field "

    Public Sub MakeSystemField(ByVal blnIsEssential As Boolean, ByVal blnIsLocked As Boolean)
        m_blnIsSystemEssential = blnIsEssential
        m_blnIsSystemLocked = blnIsLocked
    End Sub
#End Region

#Region " Make Foreign Key "

    ''' <summary>
    ''' Function to make a foreign key out of a new field
    ''' </summary>
    Public Sub MakeForeignKey(ByVal strForeignTable As String, _
    Optional ByVal blnLinkVisible As Boolean = False, _
    Optional ByVal strFieldLinkCaption As String = Nothing, _
    Optional ByVal intFieldLinkSortOrder As Integer = 6)
        If Not m_eDataType = SqlDbType.Int Then
            Throw New Exception("The data type must be ""integer"" to make the field a foreign key")
        End If

        If Not Field Is Nothing AndAlso Not Field.IsForeignKey Then
            Throw New Exception("Can't make an existing field a foreign key")
        End If

        If strFieldLinkCaption Is Nothing OrElse strFieldLinkCaption.Length = 0 Then
            strFieldLinkCaption = MakePlural(m_objTable.DatabaseName)
        End If

        m_blnIsForeignKey = True
        m_strForeignTable = strForeignTable
        m_blnFieldLinkVisible = blnLinkVisible
        m_intFieldLinkSortOrder = intFieldLinkSortOrder
        m_strFieldLinkCaption = strFieldLinkCaption
    End Sub
#End Region

#Region " Insert Update Fields "

    Public Sub InsertUpdate()

        Dim blnCreatedTransaction As Boolean = False
        Dim intCaptionID As Integer = clsDBConstants.cintNULL
        Dim objField As clsField
        Dim objFieldLink As clsFieldLink = Nothing
        Dim blnTableExists As Boolean = True

        Try
            m_strExternalID = m_objTable.DatabaseName & "." & m_strDatabaseName

            '================================================================================
            'First See if this is being called by the DRM Table Creation
            '================================================================================
            If m_objDB.SysInfo.Tables(m_objTable.ID) Is Nothing Then
                blnTableExists = False
            End If

            'Start a transaction
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            Dim blnCreateAuditTrailRecord = Not IsAuditTableField

            '================================================================================
            'Create the field and field heading captions
            '================================================================================
            If m_intID = clsDBConstants.cintNULL OrElse m_intCaptionID = clsDBConstants.cintNULL Then
                If Not String.IsNullOrEmpty(m_strCaption) Then
                    intCaptionID = CreateCaption("Field - " & m_strDatabaseName, m_strCaption, blnCreateAuditTrailRecord)
                End If
            Else
                intCaptionID = m_intCaptionID
                Dim strCaption As String = Field.Caption.GetString(m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)

                If Not String.IsNullOrEmpty(m_strCaption) AndAlso Not m_strCaption = strCaption Then
                    intCaptionID = CreateCaption(Field.Caption, m_strCaption, blnCreateAuditTrailRecord)
                End If
            End If

            '================================================================================
            'Create an auto-fill information object if necessary
            '================================================================================
            Dim objAutoFillInfo As clsAutoFillInfo = Nothing
            If Not m_eAutoFillType = clsDBConstants.cintNULL Then
                objAutoFillInfo = New clsAutoFillInfo(m_objDB, m_eAutoFillType,
                                                      m_strAutoFillValue,
                                                      m_intAutoNumberFormatID)
            End If

            '================================================================================
            'Insert the field
            '================================================================================
            objField = New clsField(m_objDB, m_intID,
                                    m_strExternalID, m_intSecurityID,
                                    m_strDatabaseName, m_objTable.ID,
                                    m_eDataType, intCaptionID,
                                    m_blnIsEncrypted, m_blnIsMandatory,
                                    m_blnIsVisible, m_blnReadOnly,
                                    m_blnMultiLine, m_blnIsSystemEssential,
                                    m_blnIsSystemLocked, m_blnNullable,
                                    m_intSortOrder, m_intLength,
                                    m_eDateType, m_intScale,
                                    objAutoFillInfo, m_blnIsMultipleSequenceField,
                                    m_blnDeterminesMultipleSequence, m_blnIsExpanded,
                                    m_intNumberOfLines, m_blnAllowFreeTextEntry,
                                    m_eFormat, m_strFormatString)

            objField.InsertUpdate(blnCreateAuditTrailRecord)

            If Not m_strDatabaseName = clsDBConstants.Fields.cID Then
                If m_intID = clsDBConstants.cintNULL Then
                    SystemDB.CreateField(m_objTable, objField, m_strDefaultValue)
                Else
                    If Not m_intLength = Field.Length OrElse Not m_intScale = Field.Scale OrElse Not m_eDataType = Field.DataType Then
                        SystemDB.UpdateField(objField, Nothing)
                    End If

                    If Not Field.IsNullable = m_blnNullable Then
                        If m_blnNullable Then
                            MakeNullable()
                        Else
                            MakeNotNullable()
                        End If
                    End If
                End If
            End If

            '================================================================================
            'Create the Field Link (If Necessary)
            '================================================================================
            If m_blnIsForeignKey Then
                Dim objDRMFieldLink As clsDRMFieldLink

                If m_intID = clsDBConstants.cintNULL Then
                    objDRMFieldLink = New clsDRMFieldLink(m_objDB, m_objTable, objField, m_intSecurityID, m_strForeignTable, _
                        m_strFieldLinkCaption, m_blnFieldLinkVisible, m_blnDisplayAsDropDown)
                Else
                    objDRMFieldLink = New clsDRMFieldLink(m_objDB, Field.FieldLink)
                    objDRMFieldLink.Caption = m_strFieldLinkCaption
                    objDRMFieldLink.IsVisible = m_blnFieldLinkVisible
                    objDRMFieldLink.DisplayAsDropDown = m_blnDisplayAsDropDown
                    objDRMFieldLink.Table = m_objTable
                End If
                If Not m_intFieldLinkSortOrder <= 0 Then
                    objDRMFieldLink.SortOrder = m_intFieldLinkSortOrder
                End If
                objDRMFieldLink.IsExpanded = m_blnIsFKeyExpanded

                objDRMFieldLink.InsertUpdate(blnCreateAuditTrailRecord)
            End If

            '================================================================================
            'Assign the field to the proper security group(s)
            '================================================================================
            If m_intID = clsDBConstants.cintNULL Then
                If m_colSecGroupIDs Is Nothing Then
                    If m_objDB.Profile.SecurityGroups Is Nothing AndAlso m_objDB.Profile.SecurityGroup IsNot Nothing Then
                        m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPFIELD,
                                                               clsDBConstants.Fields.LinkSecurityGroupField.cFIELDID,
                                                               objField.ID,
                                                               clsDBConstants.Fields.LinkSecurityGroupField.cSECURITYGROUPID,
                                                               m_objDB.Profile.SecurityGroup.ID)
                    Else
                        For Each intSecurityGroupID As Integer In m_objDB.Profile.LinkSecurityGroups.Values
                            m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPFIELD,
                                               clsDBConstants.Fields.LinkSecurityGroupField.cFIELDID,
                                               objField.ID,
                                               clsDBConstants.Fields.LinkSecurityGroupField.cSECURITYGROUPID,
                                               intSecurityGroupID)
                        Next
                    End If
                Else
                    For Each intID As Integer In m_colSecGroupIDs
                        m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPFIELD,
                                           clsDBConstants.Fields.LinkSecurityGroupField.cFIELDID,
                                           objField.ID,
                                           clsDBConstants.Fields.LinkSecurityGroupField.cSECURITYGROUPID,
                                           intID)
                    Next
                End If
            Else
                'copy type info over to new field
                For Each strKey As String In Field.TypeFieldInfos.Keys
                    objField.TypeFieldInfos(strKey) = Field.TypeFieldInfos(strKey)
                Next

                If objFieldLink IsNot Nothing Then
                    For Each strKey As String In Field.FieldLink.TypeFieldLinkInfos.Keys
                        objFieldLink.TypeFieldLinkInfos(strKey) = Field.FieldLink.TypeFieldLinkInfos(strKey)
                    Next
                End If
            End If

            m_objDB.SysInfo.DRMInsertUpdateField(m_objTable, objField)
            If objFieldLink IsNot Nothing Then
                m_objDB.SysInfo.DRMInsertUpdateFieldLink(m_objTable, objFieldLink)
            End If

            If blnTableExists Then
                SystemDB.CreateStandardSPs(m_objTable)
            End If

            m_intID = objField.ID
            m_intCaptionID = intCaptionID
            m_objDBObj = objField

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)

        Catch ex As Exception

            If blnCreatedTransaction Then m_objDB.EndTransaction(False)
            Throw

        End Try

    End Sub

#End Region

#Region " Full Text indexing Functions "

    Public Sub RemoveFullTextIndexes(ByVal blnForceDelete As Boolean)
        If Field IsNot Nothing Then
            If (m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cMETADATAPROFILE.ToUpper OrElse _
            m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cEDOC.ToUpper OrElse _
            m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cSPACE.ToUpper OrElse _
            (m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cTITLE.ToUpper AndAlso _
            (Field.DatabaseName = clsDBConstants.Fields.cEXTERNALID OrElse _
            Field.DatabaseName = clsDBConstants.Fields.Title.cCODE))) AndAlso _
            Field.IsTextType AndAlso (blnForceDelete OrElse Not Field.Length = m_intLength) Then
                Dim strTable As String = m_objTable.DatabaseName

                If clsDBConstants.Tables.cEDOC = strTable Then
                    strTable = clsDBConstants.Views.cEDOCIndexView
                    '-- Need to rebuild the view before modifying the field
                    Dim arrFields As New FrameworkCollections.K1Collection(Of String)

                    arrFields.Add(Field.DatabaseName)
                    m_objDB.RebuildEdocView(arrFields)
                End If

                If m_objDB.FullTextIndexHasColumn(strTable, Field.DatabaseName) Then
                    m_objDB.FullTextIndexDeleteColumn(strTable, Field.DatabaseName)
                End If
            End If
        End If

        If Field IsNot Nothing AndAlso Not m_strDatabaseName = Field.DatabaseName Then
            m_objDB.RenameField(m_objTable.DatabaseName, _
                Field.DatabaseName, m_strDatabaseName)
        End If
    End Sub

    Public Sub AddFullTextIndexes()
        If Field IsNot Nothing Then
            If (m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cMETADATAPROFILE.ToUpper OrElse _
            m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cEDOC.ToUpper OrElse _
            m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cSPACE.ToUpper OrElse _
            (m_objTable.DatabaseName.ToUpper = clsDBConstants.Tables.cTITLE.ToUpper AndAlso _
            (Field.DatabaseName = clsDBConstants.Fields.cEXTERNALID OrElse _
            Field.DatabaseName = clsDBConstants.Fields.Title.cCODE))) AndAlso _
            Field.IsTextType Then
                Dim strTable As String = m_objTable.DatabaseName

                If clsDBConstants.Tables.cEDOC = strTable Then
                    strTable = clsDBConstants.Views.cEDOCIndexView
                    '-- Need to rebuild the view before adding field to the index
                    m_objDB.RebuildEdocView()
                End If

                If Not m_objDB.FullTextIndexHasColumn(strTable, m_strDatabaseName) Then
                    m_objDB.FullTextIndexAddColumn(strTable, m_strDatabaseName)
                End If
            End If
        End If
    End Sub
#End Region

#Region " Delete "

    Public Sub Delete()

        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            If Field.FieldLink IsNot Nothing Then
                SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cFILTER & "] " & _
                    "WHERE [" & clsDBConstants.Fields.Filter.cINITFILTERFIELDLINKID & "] = " & Field.FieldLink.ID)

                SystemDB.DeleteForeignKey(Field, Field.FieldLink.IdentityTable)
                SystemDB.DeleteRecord(clsDBConstants.Tables.cFIELDLINK, Field.FieldLink.ID)

                Dim objCaption As clsCaption = Field.FieldLink.Caption
                m_objDB.SysInfo.DRMDeleteFieldLink(Field.FieldLink)

                DeleteCaption(objCaption, Not IsAuditTableField)
            End If

            SystemDB.DropDefaultConstraint(Field)
            clsTableIndex.DropIndexes(SystemDB, Field)
            SystemDB.DeleteField(Field)

            'remove any filters related to this field which won't get removed from the recursive delete below
            SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cFILTER & "] " & _
                "WHERE [" & clsDBConstants.Fields.Filter.cFILTERVALUEFIELDID & "] = " & Field.ID)
            SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cFILTER & "] " & _
                "WHERE [" & clsDBConstants.Fields.Filter.cINITFILTERFIELDID & "] = " & Field.ID)

            RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cFIELD, Field.ID, Not IsAuditTableField)

            DeleteCaption(Field.Caption, Not IsAuditTableField)

            m_objDB.SysInfo.DRMDeleteField(m_objTable, Field)
            SystemDB.CreateStandardSPs(m_objTable)

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)

        Catch ex As Exception

            If blnCreatedTransaction Then m_objDB.EndTransaction(False)
            m_objDB.RefreshSysInfo()
            Throw

        End Try

    End Sub

#End Region

#Region " Make Nullable\Not Nullable "

    Public Sub MakeNullable()
        Try
            m_blnNullable = True
            m_objDB.ExecuteSQL(String.Format("ALTER TABLE {0} ALTER COLUMN {1}", m_objTable.DatabaseName, GetFieldDeclaration()))
        Catch
            Throw
        End Try
    End Sub

    Public Sub MakeNotNullable()
        Try
            MakeNotNullable(False, Nothing)
        Catch
            Throw
        End Try
    End Sub

    Public Sub MakeNotNullable(ByVal blnDeleteNullValues As Boolean)
        Try
            MakeNotNullable(blnDeleteNullValues, Nothing)
        Catch
            Throw
        End Try
    End Sub

    Public Sub MakeNotNullable(ByVal blnDeleteNullValues As Boolean, ByVal strDefaultValue As String)
        Try
            If blnDeleteNullValues Then
                '-- Delete null values before making it not null
                m_objDB.ExecuteSQL(String.Format("DELETE FROM {0} WHERE {1} IS NULL", m_objTable.DatabaseName, m_strDatabaseName))
            End If

            '-- Get list of indexes that are bound to this field
            Dim colFieldIndexes As Generic.List(Of clsTableIndex) = _
                clsTableIndex.GetTableIndexes(m_objDB, m_objTable.DatabaseName, m_strDatabaseName)

            If colFieldIndexes.Count > 0 Then
                '-- Need to drop the field indexes before we can alter it,
                clsTableIndex.DropIndexes(m_objDB, m_objTable.DatabaseName, m_strDatabaseName)
            End If
            m_blnNullable = False
            m_strDefaultValue = strDefaultValue
            m_objDB.ExecuteSQL(String.Format("ALTER TABLE {0} ALTER COLUMN {1}", m_objTable.DatabaseName, GetFieldDeclaration()))

            '-- Recreate the dropped indexes
            For Each objIndex As clsTableIndex In colFieldIndexes
                objIndex.Create(m_objDB)
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region " Get Field Declaration "

    Private Function GetFieldDeclaration() As String
        Dim intLength As Integer = clsDBConstants.cintNULL

        If m_intLength > 0 Then
            intLength = m_intLength
        ElseIf m_intPrecision > 0 Then
            intLength = m_intPrecision
        End If

        Return Me.SystemDB.GetFieldDeclaration(m_strDatabaseName, m_eDataType, m_intLength, m_intScale, _
            m_blnNullable)
    End Function
#End Region

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
        m_objTable = Nothing
    End Sub
#End Region

End Class
