#Region " File Information "

'=====================================================================
' This class represents the table Field in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      6/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Imports System.Data.OleDb

Public Class clsField
    Inherits clsDBObjBase

#Region " Members "

    Private m_intTableID As Integer
    Private m_eDataType As SqlDbType
    Private m_intLength As Integer
    Private m_blnIsMandatory As Boolean
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_colLanguageCaptions As Hashtable
    Private m_colLanguageColHeadings As Hashtable
    Private m_blnIsEncrypted As Boolean
    Private m_eDateType As clsDBConstants.enumDateTypes
    Private m_blnIsVisible As Boolean = True
    Private m_strDatabaseName As String
    Private m_blnIsReadOnly As Boolean = False
    Private m_blnIsMultiLine As Boolean = False
    Private m_blnIsSystemEssential As Boolean = False
    Private m_blnIsSystemLocked As Boolean = False
    Private m_blnIsNullable As Boolean = False
    Private m_intSortOrder As Integer
    Private m_intScale As Integer
    Private m_colTypeFieldInfos As FrameworkCollections.K1Dictionary(Of clsTypeField)
    Private m_objAutoFillInfo As clsAutoFillInfo
    Private m_blnIsMultipleSequenceField As Boolean = False
    Private m_blnDeterminesMultipleSequence As Boolean = False
    Private m_eFormatType As clsDBConstants.enumFormatType
    Private m_strFormat As String
    Private m_objFieldLink As clsFieldLink  'This is only used when creating fields on the fly
    Private m_blnIsExpanded As Boolean
    Private m_intNumberOfLines As Integer
    Private m_blnAllowFreeTextEntry As Boolean = False
    Private m_colFilterCollection As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
    Private m_colParentFilters As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal intID As Integer)

        MyBase.New(Nothing, intID, intID.ToString(), clsDBConstants.cintNULL, clsDBConstants.cintNULL)
        Me.ComputeFunction = String.Empty
        Me.IsComputedField = False
    End Sub

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal strDatabaseName As String, _
    ByVal intTableID As Integer, _
    ByVal eDataType As SqlDbType, _
    ByVal intCaptionID As Integer, _
    ByVal blnIsEncrypted As Boolean, _
    ByVal blnIsMandatory As Boolean, _
    ByVal blnIsVisible As Boolean, _
    ByVal blnIsReadOnly As Boolean, _
    ByVal blnIsMultiLine As Boolean, _
    ByVal blnIsSystemEssential As Boolean, _
    ByVal blnIsSystemLocked As Boolean, _
    ByVal blnIsNullable As Boolean, _
    ByVal intSortOrder As Integer, _
    ByVal intLength As Integer, _
    ByVal eDateType As clsDBConstants.enumDateTypes, _
    ByVal intScale As Integer, _
    ByVal objAutoFillInfo As clsAutoFillInfo, _
    ByVal blnIsMultipleSequenceField As Boolean, _
    ByVal blnDeterminesMultipleSequence As Boolean, _
    ByVal blnIsExpanded As Boolean, _
    ByVal intNumberOfLines As Integer, _
    ByVal blnAllowFreeTextEntry As Boolean, _
    ByVal eFormat As clsDBConstants.enumFormatType, _
    ByVal strFormatString As String)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_strDatabaseName = strDatabaseName
        m_eDataType = eDataType
        m_intTableID = intTableID
        m_intCaptionID = intCaptionID
        m_objAutoFillInfo = objAutoFillInfo
        m_intLength = intLength
        m_eDateType = eDateType
        m_blnIsMandatory = blnIsMandatory
        m_blnIsEncrypted = blnIsEncrypted
        m_blnIsVisible = blnIsVisible
        m_blnIsReadOnly = blnIsReadOnly
        m_blnIsMultiLine = blnIsMultiLine
        m_blnIsSystemEssential = blnIsSystemEssential
        m_blnIsSystemLocked = blnIsSystemLocked
        m_blnIsNullable = blnIsNullable
        m_intSortOrder = intSortOrder
        m_intScale = intScale
        m_blnIsMultipleSequenceField = blnIsMultipleSequenceField
        m_blnDeterminesMultipleSequence = blnDeterminesMultipleSequence
        m_blnIsExpanded = blnIsExpanded
        m_intNumberOfLines = intNumberOfLines
        m_blnAllowFreeTextEntry = blnAllowFreeTextEntry
        m_eFormatType = eFormat
        m_strFormat = strFormatString

        Me.ComputeFunction = String.Empty
        Me.IsComputedField = False
    End Sub

#End Region

#Region " From Database "

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cTABLEID, clsDBConstants.cintNULL), Integer)
        m_eDataType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cDATATYPE, SqlDbType.Int), SqlDbType)
        m_intLength = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cLENGTH, clsDBConstants.cintNULL), Integer)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_blnIsMandatory = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISMANDATORY, False), Boolean)
        m_blnIsEncrypted = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISENCRYPTED, False), Boolean)
        m_eDateType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cDATETYPE, clsDBConstants.enumDateTypes.DATE_AND_TIME), clsDBConstants.enumDateTypes)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISVISIBLE, True), Boolean)
        m_strDatabaseName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cDATABASENAME, clsDBConstants.cstrNULL), String)
        m_blnIsReadOnly = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISREADONLY, False), Boolean)
        m_blnIsMultiLine = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISMULTILINE, False), Boolean)
        m_blnIsSystemEssential = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISSYSTEMESSENTIAL, False), Boolean)
        m_blnIsSystemLocked = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISSYSTEMLOCKED, False), Boolean)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cSORTORDER, clsDBConstants.cintNULL), Integer)
        m_intScale = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cSCALE, clsDBConstants.cintNULL), Integer)
        Dim eAutoFillType As clsDBConstants.enumAutoFillTypes
        Dim strAutoFillValue As String
        Dim intAutoNumberFormatID As Integer
        eAutoFillType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cAUTOFILLTYPE, clsDBConstants.cintNULL), clsDBConstants.enumAutoFillTypes)
        strAutoFillValue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cAUTOFILLVALUE, clsDBConstants.cstrNULL), String)
        intAutoNumberFormatID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cAUTONUMBERFORMATID, clsDBConstants.cintNULL), Integer)
        If Not eAutoFillType = clsDBConstants.enumAutoFillTypes.NOTHING Then
            m_objAutoFillInfo = New clsAutoFillInfo(m_objDB, eAutoFillType, strAutoFillValue, intAutoNumberFormatID)
        End If
        'TODO: add this field to the database
        m_blnIsNullable = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.Field.cISSYSTEMNULLABLE, True), Boolean)
        m_blnDeterminesMultipleSequence = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.Field.cDETERMINESMULTIPLESEQUENCE, False), Boolean)
        If m_blnDeterminesMultipleSequence Then
            m_blnIsMultipleSequenceField = True
        Else
            m_blnIsMultipleSequenceField = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.Field.cISMULTIPLESEQUENCEFIELD, False), Boolean)
        End If
        m_eFormatType = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.Field.cFORMATTYPE, clsDBConstants.enumFormatType.None), clsDBConstants.enumFormatType)
        m_strFormat = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cFORMATSTRING, clsDBConstants.cstrNULL), String)
        m_blnIsExpanded = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cISEXPANDED, False), Boolean)
        m_intNumberOfLines = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cNUMBEROFLINES, clsDBConstants.cintNULL), Integer)
        If m_intNumberOfLines < 1 Then
            m_intNumberOfLines = 4
        ElseIf m_intNumberOfLines > 100 Then
            m_intNumberOfLines = 100
        End If
        m_blnAllowFreeTextEntry = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Field.cALLOWFREETEXTENTRY, False), Boolean)
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB, ByVal blnWSImproved As Boolean)
        m_intID = CType(clsDB_Direct.DataRowValue(objDR, "A", clsDBConstants.cintNULL), Integer)
        m_intSecurityID = CType(clsDB_Direct.DataRowValue(objDR, "B", clsDBConstants.cintNULL), Integer)
        m_strExternalID = CType(clsDB_Direct.DataRowValue(objDR, "C", clsDBConstants.cstrNULL), String)
        m_intTypeID = CType(clsDB_Direct.DataRowValue(objDR, "D", clsDBConstants.cintNULL), Integer)
        m_objDB = objDB
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, "E", clsDBConstants.cintNULL), Integer)
        m_eDataType = CType(clsDB_Direct.DataRowValue(objDR, "F", SqlDbType.Int), SqlDbType)
        m_intLength = CType(clsDB_Direct.DataRowValue(objDR, "G", clsDBConstants.cintNULL), Integer)
        m_blnIsMandatory = CType(clsDB_Direct.DataRowValue(objDR, "H", False), Boolean)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, "I", clsDBConstants.cintNULL), Integer)
        m_eDateType = CType(clsDB_Direct.DataRowValue(objDR, "J", clsDBConstants.enumDateTypes.DATE_AND_TIME), clsDBConstants.enumDateTypes)
        m_blnIsEncrypted = CType(clsDB_Direct.DataRowValue(objDR, "K", False), Boolean)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, "L", True), Boolean)
        m_strDatabaseName = CType(clsDB_Direct.DataRowValue(objDR, "M", clsDBConstants.cstrNULL), String)
        m_intScale = CType(clsDB_Direct.DataRowValue(objDR, "N", clsDBConstants.cintNULL), Integer)
        m_blnIsReadOnly = CType(clsDB_Direct.DataRowValue(objDR, "O", False), Boolean)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, "P", clsDBConstants.cintNULL), Integer)
        m_blnIsMultiLine = CType(clsDB_Direct.DataRowValue(objDR, "Q", False), Boolean)
        m_blnIsSystemEssential = CType(clsDB_Direct.DataRowValue(objDR, "R", False), Boolean)
        m_blnIsSystemLocked = CType(clsDB_Direct.DataRowValue(objDR, "S", False), Boolean)
        Dim eAutoFillType As clsDBConstants.enumAutoFillTypes
        Dim strAutoFillValue As String
        Dim intAutoNumberFormatID As Integer
        eAutoFillType = CType(clsDB_Direct.DataRowValue(objDR, "T", clsDBConstants.cintNULL), clsDBConstants.enumAutoFillTypes)
        strAutoFillValue = CType(clsDB_Direct.DataRowValue(objDR, "U", clsDBConstants.cstrNULL), String)
        intAutoNumberFormatID = CType(clsDB_Direct.DataRowValue(objDR, "V", clsDBConstants.cintNULL), Integer)
        If Not eAutoFillType = clsDBConstants.enumAutoFillTypes.NOTHING Then
            m_objAutoFillInfo = New clsAutoFillInfo(m_objDB, eAutoFillType, strAutoFillValue, intAutoNumberFormatID)
        End If
        'TODO: add this field to the database
        m_blnIsNullable = CType(clsDB.DataRowValue(objDR, "W", True), Boolean)
        m_eFormatType = CType(clsDB.DataRowValue(objDR, "X", clsDBConstants.enumFormatType.None), clsDBConstants.enumFormatType)
        m_strFormat = CType(clsDB_Direct.DataRowValue(objDR, "Y", clsDBConstants.cstrNULL), String)
        m_blnDeterminesMultipleSequence = CType(clsDB.DataRowValue(objDR, "C1", False), Boolean)
        If m_blnDeterminesMultipleSequence Then
            m_blnIsMultipleSequenceField = True
        Else
            m_blnIsMultipleSequenceField = CType(clsDB.DataRowValue(objDR, "Z", False), Boolean)
        End If

        m_blnIsExpanded = CType(clsDB_Direct.DataRowValue(objDR, "C2", False), Boolean)
        m_intNumberOfLines = CType(clsDB_Direct.DataRowValue(objDR, "C3", clsDBConstants.cintNULL), Integer)
        If m_intNumberOfLines < 1 Then
            m_intNumberOfLines = 4
        ElseIf m_intNumberOfLines > 100 Then
            m_intNumberOfLines = 100
        End If
        m_blnAllowFreeTextEntry = CType(clsDB_Direct.DataRowValue(objDR, "C4", False), Boolean)
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property DataType() As SqlDbType
        Get
            Return m_eDataType
        End Get
    End Property

    Public ReadOnly Property Length() As Integer
        Get
            Return m_intLength
        End Get
    End Property

    Public ReadOnly Property IsMandatory() As Boolean
        Get
            Return m_blnIsMandatory
        End Get
    End Property

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property CaptionText() As String
        Get
            Return m_objDB.SysInfo.GetCaptionString(m_intCaptionID).ToString()
        End Get
    End Property

    Public ReadOnly Property Caption() As clsCaption
        Get
            If m_objCaption Is Nothing Then
                If Not m_intCaptionID = clsDBConstants.cintNULL Then
                    m_objCaption = clsCaption.GetItem(m_intCaptionID, Me.Database)
                End If
            End If

            Return m_objCaption
        End Get
    End Property

    Public ReadOnly Property CaptionID() As Integer
        Get
            Return m_intCaptionID
        End Get
    End Property

    Public ReadOnly Property IsTextType() As Boolean
        Get
            If m_eDataType = SqlDbType.Char OrElse _
            m_eDataType = SqlDbType.NChar OrElse _
            m_eDataType = SqlDbType.NText OrElse _
            m_eDataType = SqlDbType.NVarChar OrElse _
            m_eDataType = SqlDbType.Text OrElse _
            m_eDataType = SqlDbType.VarChar Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsNumericType() As Boolean
        Get
            If m_eDataType = SqlDbType.BigInt OrElse _
            m_eDataType = SqlDbType.Bit OrElse _
            m_eDataType = SqlDbType.Decimal OrElse _
            m_eDataType = SqlDbType.Float OrElse _
            m_eDataType = SqlDbType.Int OrElse _
            m_eDataType = SqlDbType.Money OrElse _
            m_eDataType = SqlDbType.Real OrElse _
            m_eDataType = SqlDbType.SmallInt OrElse _
            m_eDataType = SqlDbType.SmallMoney OrElse _
            m_eDataType = SqlDbType.TinyInt Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsDateType() As Boolean
        Get
            If m_eDataType = SqlDbType.Date OrElse _
             m_eDataType = SqlDbType.DateTime OrElse _
             m_eDataType = SqlDbType.DateTime2 OrElse _
             m_eDataType = SqlDbType.SmallDateTime Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsBinaryType() As Boolean
        Get
            If m_eDataType = SqlDbType.Binary OrElse _
            m_eDataType = SqlDbType.VarBinary OrElse _
            m_eDataType = SqlDbType.Image Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsEncrypted() As Boolean
        Get
            Return m_blnIsEncrypted
        End Get
    End Property

    Public ReadOnly Property DateType() As clsDBConstants.enumDateTypes
        Get
            Return m_eDateType
        End Get
    End Property

    Public ReadOnly Property IsIdentityField() As Boolean
        Get
            If UCase(m_strDatabaseName.Trim) = UCase(clsDBConstants.Fields.cID) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
    End Property

    Public ReadOnly Property IsExpanded() As Boolean
        Get
            Return m_blnIsExpanded
        End Get
    End Property

    Public ReadOnly Property DatabaseName() As String
        Get
            Return m_strDatabaseName
        End Get
    End Property

    Public ReadOnly Property IsReadOnly() As Boolean
        Get
            Return m_blnIsReadOnly
        End Get
    End Property

    Public ReadOnly Property IsMultiLine() As Boolean
        Get
            Return m_blnIsMultiLine
        End Get
    End Property

    Public ReadOnly Property NumberOfLines() As Integer
        Get
            Return m_intNumberOfLines
        End Get
    End Property

    Public ReadOnly Property IsSystemEssential() As Boolean
        Get
            Return m_blnIsSystemEssential
        End Get
    End Property

    Public ReadOnly Property IsSystemLocked() As Boolean
        Get
            Return m_blnIsSystemLocked
        End Get
    End Property

    Public ReadOnly Property IsNullable() As Boolean
        Get
            Return m_blnIsNullable
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property

    Public ReadOnly Property Scale() As Integer
        Get
            Return m_intScale
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

    Public ReadOnly Property IsInvalidListType() As Boolean
        Get
            If IsBinaryType OrElse _
            m_eDataType = SqlDbType.Variant OrElse _
            m_eDataType = SqlDbType.UniqueIdentifier OrElse _
            m_blnIsEncrypted = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsAutoNumberField() As Boolean
        Get
            If m_objAutoFillInfo Is Nothing OrElse _
            m_objAutoFillInfo.AutoNumberFormat Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property TypeFieldInfos() As FrameworkCollections.K1Dictionary(Of clsTypeField)
        Get
            If m_colTypeFieldInfos Is Nothing Then
                m_colTypeFieldInfos = New FrameworkCollections.K1Dictionary(Of clsTypeField) 'clsTypeField.GetList(Me, m_objDB)
            End If
            Return m_colTypeFieldInfos
        End Get
    End Property

    Public ReadOnly Property Filters(ByVal intTypeID As Integer) As List(Of clsFilter)
        Get
            Dim colFilters As List(Of clsFilter) = Nothing

            If Not intTypeID = clsDBConstants.cintNULL Then
                colFilters = m_objDB.SysInfo.Filters("F_" & m_intID & "_" & intTypeID)
            End If

            If colFilters Is Nothing Then
                colFilters = m_objDB.SysInfo.Filters("F_" & m_intID)
            End If

            Return colFilters
        End Get
    End Property

    Public ReadOnly Property AutoFillInfo() As clsAutoFillInfo
        Get
            Return m_objAutoFillInfo
        End Get
    End Property

    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objDB.SysInfo.Tables(m_intTableID)
        End Get
    End Property

    Public ReadOnly Property IsForeignKey() As Boolean
        Get
            If FieldLink Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public Property FieldLink() As clsFieldLink
        Get
            If m_objFieldLink Is Nothing Then
                Return m_objDB.SysInfo.FieldLinks(CStr(m_intID))
            Else
                Return m_objFieldLink
            End If
        End Get
        Set(ByVal value As clsFieldLink)
            m_objFieldLink = value
        End Set
    End Property

    Public ReadOnly Property KeyName() As String
        Get
            Return CStr(m_intTableID) & "_" & m_strDatabaseName
        End Get
    End Property

    Public ReadOnly Property IsRangeField() As Boolean
        Get
            If m_eDataType = SqlDbType.Bit OrElse _
            IsBinaryType OrElse _
            m_eDataType = SqlDbType.NText OrElse _
            m_eDataType = SqlDbType.Text OrElse _
            IsForeignKey Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property FormatType() As clsDBConstants.enumFormatType
        Get
            Return m_eFormatType
        End Get
    End Property

    Public ReadOnly Property FormatString() As String
        Get
            Return m_strFormat
        End Get
    End Property

    Public ReadOnly Property IsSortable() As Boolean
        Get
            If IsBinaryType OrElse _
            m_eDataType = SqlDbType.NText OrElse _
            m_eDataType = SqlDbType.Text Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property AllowFreeTextEntry() As Boolean
        Get
            Return m_blnAllowFreeTextEntry
        End Get
    End Property

    Public Property ParentOfFilters() As FrameworkCollections.K1Dictionary(Of List(Of clsFilter))
        Get
            Return m_colParentFilters
        End Get
        Set(ByVal value As FrameworkCollections.K1Dictionary(Of List(Of clsFilter)))
            m_colParentFilters = value
        End Set
    End Property

    ''' <summary>
    ''' Used when the field does not map to a physical column but to a scalar function in the database 
    ''' i.e. select *, {tablename}_ComputedField( [{tablename}].[ID] ) as CaptionText from {tablename}
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsComputedField() As Boolean

    ''' <summary>
    ''' Used when the field does not map to a physical column but to a scalar function in the database 
    ''' i.e. select *, {tablename}_ComputedField( [{tablename}].[ID] ) as CaptionText from {tablename}
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ComputeFunction() As String


#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsField
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cFIELD, intID)

            Return New clsField(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
        Dim objDT As DataTable
        Dim strSP As String
        Dim colFields As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)

        Try
            strSP = clsDBConstants.Tables.cFIELD & clsDBConstants.StoredProcedures.cGETLIST

            objDT = objDB.GetDataTable(strSP)

            colFields = New FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objField As New clsField(objDataRow, objDB)
                colFields.Add(objField.KeyName, objField)
            Next

            Return colFields
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList2(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
        Dim objDT As DataTable
        Dim strSP As String
        Dim colFields As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)

        Try
            strSP = "SELECT [Field].[ID] AS 'A', [Field].[SecurityID] AS 'B', " &
                "[Field].[ExternalID] AS 'C', [Field].[TypeID] AS 'D', [Field].[TableID] AS 'E', " &
                "[Field].[DataType] AS 'F', [Field].[Length] AS 'G', [Field].[isMandatory] AS 'H', " &
                "[Field].[CaptionID] AS 'I', [Field].[DateType] AS 'J', [Field].[isEncrypted] AS 'K', " &
                "[Field].[isVisible] AS 'L', [Field].[DatabaseName] AS 'M', [Field].[Scale] AS 'N', " &
                "[Field].[isReadOnly] AS 'O', [Field].[SortOrder] AS 'P', [Field].[IsMultiLine] AS 'Q', " &
                "[Field].[isSystemEssential] AS 'R', [Field].[isSystemLocked] AS 'S', [Field].[AutoFillType] AS 'T', " &
                "[Field].[AutoFillValue] AS 'U', [Field].[AutoNumberFormatID] AS 'V', [Field].[isSystemNullable] AS 'W', " &
                "[Field].[FormatType] AS 'X', [Field].[FormatString] AS 'Y', [Field].[isMultipleSequenceField] AS 'Z', " &
                "[Field].[DeterminesMultipleSequence] AS 'C1', [Field].[isExpanded] AS 'C2', [Field].[NumberOfLines] AS 'C3', " &
                "[Field].[AllowFreeTextEntry] AS 'C4' FROM [Field]"

            objDT = objDB.GetDataTableBySQL(strSP)

            colFields = New FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objField As New clsField(objDataRow, objDB, True)
                colFields.Add(objField.KeyName, objField)
            Next

            Return colFields
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function isInType(intType As Integer) As Boolean
        If Me.DatabaseName = "TypeID" Then
            Return True
        End If

        Return (From tfi In TypeFieldInfos
                Where tfi.Value.AppliesToTypeID = intType And tfi.Value.IsVisible
                Select tfi.Value.AppliesToTypeID).Count > 0
    End Function


#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrailRecord As Boolean = False)

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(m_objDB.SysInfo.Tables(clsDBConstants.Tables.cFIELD), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cDATABASENAME, m_strDatabaseName)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cDATATYPE, m_eDataType)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cTABLEID, m_intTableID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cCAPTIONID, m_intCaptionID)
        'colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cCOLUMNHEADINGCAPTIONID, m_intColHeadingCaptionID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cLENGTH, m_intLength)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cDATETYPE, m_eDateType)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISMANDATORY, m_blnIsMandatory)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISENCRYPTED, m_blnIsEncrypted)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISVISIBLE, m_blnIsVisible)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISREADONLY, m_blnIsReadOnly)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISMULTILINE, m_blnIsMultiLine)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISSYSTEMESSENTIAL, m_blnIsSystemEssential)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISSYSTEMLOCKED, m_blnIsSystemLocked)
        'colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISCOLUMN, m_blnIsColumn)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISSYSTEMNULLABLE, m_blnIsNullable)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cSORTORDER, m_intSortOrder)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cSCALE, m_intScale)
        'colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cCOLUMNWIDTH, m_intColumnWidth)
        'colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISWIDTHPERCENTAGE, m_blnIsWidthPercentage)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISMULTIPLESEQUENCEFIELD, m_blnIsMultipleSequenceField)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cDETERMINESMULTIPLESEQUENCE, m_blnDeterminesMultipleSequence)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cISEXPANDED, m_blnIsExpanded)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cNUMBEROFLINES, m_intNumberOfLines)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cALLOWFREETEXTENTRY, m_blnAllowFreeTextEntry)

        If Not m_eFormatType = clsDBConstants.enumFormatType.None Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cFORMATTYPE, m_eFormatType)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cFORMATSTRING, m_strFormat)
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cFORMATTYPE, Nothing)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cFORMATSTRING, Nothing)
        End If

        If Not m_objAutoFillInfo Is Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLTYPE, m_objAutoFillInfo.FillType)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLVALUE, m_objAutoFillInfo.FillValue)
            If Not m_objAutoFillInfo.AutoNumberFormat Is Nothing Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTONUMBERFORMATID, _
                    m_objAutoFillInfo.AutoNumberFormat.ID)
            End If
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLTYPE, Nothing)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLVALUE, Nothing)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTONUMBERFORMATID, Nothing)
        End If

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB, blnCreateAuditTrailRecord)
        Else
            colMasks.Update(m_objDB, blnCreateAuditTrailRecord)
        End If

    End Sub

#End Region

#Region " Security "

    Public Function HasAccess() As Boolean
        If m_objDB.Profile.HasAccess(m_intSecurityID) AndAlso _
        m_objDB.Profile.LinkFields(CStr(m_intID)) IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_objCaption Is Nothing Then
                m_objCaption.Dispose()
                m_objCaption = Nothing
            End If

            If m_colTypeFieldInfos IsNot Nothing Then
                m_colTypeFieldInfos.Dispose()
                m_colTypeFieldInfos = Nothing
            End If

            If Not m_objAutoFillInfo Is Nothing Then
                m_objAutoFillInfo.Dispose()
                m_objAutoFillInfo = Nothing
            End If

            If Not m_objFieldLink Is Nothing Then
                m_objFieldLink.Dispose()
                m_objFieldLink = Nothing
            End If

            If Not m_colFilterCollection Is Nothing Then
                For Each colList As List(Of clsFilter) In m_colFilterCollection.Values
                    colList.Clear()
                    colList = Nothing
                Next
                m_colFilterCollection.Dispose()
                m_colFilterCollection = Nothing
            End If

            If Not m_colParentFilters Is Nothing Then
                For Each colList As List(Of clsFilter) In m_colParentFilters.Values
                    colList.Clear()
                    colList = Nothing
                Next
                m_colParentFilters.Dispose()
                m_colParentFilters = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
