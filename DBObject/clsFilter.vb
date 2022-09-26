#Region " File Information "

'=====================================================================
' This class represents the table Filter in the Database.
' It is used to filter the result list on search pages when coming
' from designated tables.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Kevin      15/10/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsFilter
    Inherits clsDBObjBase

#Region " Members "

    Private m_intInitFilterFieldID As Integer
    Private m_intInitFilterFieldLinkID As Integer
    Private m_intFilterByFieldID As Integer
    Private m_eFilterType As clsDBConstants.enumFilterTypes
    Private m_strFilterValue As String
    Private m_intValueFieldID As Integer
    Private m_intAppliesToTypeID As Integer
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intInitFilterFieldID As Integer, _
    ByVal intInitFilterFieldLinkID As Integer, _
    ByVal intFilterByFieldID As Integer, _
    ByVal eFilterType As clsDBConstants.enumFilterTypes, _
    ByVal strFilterValue As String, _
    ByVal intValueFieldID As Integer, _
    ByVal intAppliesToTypeID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intFilterByFieldID = intFilterByFieldID
        m_intValueFieldID = intValueFieldID
        m_intInitFilterFieldID = intInitFilterFieldID
        m_intInitFilterFieldLinkID = intInitFilterFieldLinkID
        m_eFilterType = eFilterType
        m_strFilterValue = strFilterValue
        m_intAppliesToTypeID = intAppliesToTypeID
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intFilterByFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cFILTERBYFIELDID, clsDBConstants.cintNULL), Integer)
        m_intValueFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cFILTERVALUEFIELDID, clsDBConstants.cintNULL), Integer)
        m_eFilterType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cFILTERTYPE, clsDBConstants.cintNULL), clsDBConstants.enumFilterTypes)
        m_intInitFilterFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cINITFILTERFIELDID, clsDBConstants.cintNULL), Integer)
        m_intInitFilterFieldLinkID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cINITFILTERFIELDLINKID, clsDBConstants.cintNULL), Integer)
        m_strFilterValue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cFILTERVALUE, clsDBConstants.cstrNULL), String)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Filter.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property FilterByFieldID() As Integer
        Get
            Return m_intFilterByFieldID
        End Get
    End Property

    Public ReadOnly Property FilterByField() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intFilterByFieldID)
        End Get
    End Property

    Public ReadOnly Property FilterType() As clsDBConstants.enumFilterTypes
        Get
            Return m_eFilterType
        End Get
    End Property

    Public ReadOnly Property InitFilterFieldID() As Integer
        Get
            Return m_intInitFilterFieldID
        End Get
    End Property

    Public ReadOnly Property InitFilterField() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intInitFilterFieldID)
        End Get
    End Property

    Public ReadOnly Property InitFilterFieldLinkID() As Integer
        Get
            Return m_intInitFilterFieldLinkID
        End Get
    End Property

    Public ReadOnly Property FilterValue() As String
        Get
            Return m_strFilterValue
        End Get
    End Property

    Public ReadOnly Property FilterValueFieldID() As Integer
        Get
            Return m_intValueFieldID
        End Get
    End Property

    Public ReadOnly Property FilterValueField() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intValueFieldID)
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property

    Public ReadOnly Property KeyName() As String
        Get
            Dim strKey As String
            If Not m_intInitFilterFieldID = clsDBConstants.cintNULL Then
                'Filter on a field link
                If m_intAppliesToTypeID = clsDBConstants.cintNULL Then
                    strKey = "F_" & m_intInitFilterFieldID
                Else
                    strKey = "F_" & m_intInitFilterFieldID & "_" & m_intAppliesToTypeID
                End If
            Else
                'Filter on a field
                If m_intAppliesToTypeID = clsDBConstants.cintNULL Then
                    strKey = "L_" & m_intInitFilterFieldLinkID
                Else
                    strKey = "L_" & m_intInitFilterFieldLinkID & "_" & m_intAppliesToTypeID
                End If
            End If
            Return strKey
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsFilter
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cFILTER, intID)

            Return New clsFilter(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As List(Of clsFilter)
        Dim colObjects As List(Of clsFilter)

        Try
            Dim strSP As String = clsDBConstants.Tables.cFILTER & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colObjects = New List(Of clsFilter)
            For Each objDR As DataRow In objDT.Rows
                colObjects.Add(New clsFilter(objDR, objDB))
            Next

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objFilter As clsFilter, ByVal objDB As clsDB) As List(Of clsFilter)
        Dim colObjects As List(Of clsFilter)

        Try
            Dim colParams As New clsDBParameterDictionary

            If objFilter.InitFilterFieldID = clsDBConstants.cintNULL Then
                colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.Filter.cINITFILTERFIELDLINKID), _
                    objFilter.InitFilterFieldLinkID))
            Else
                colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.Filter.cINITFILTERFIELDID), _
                    objFilter.InitFilterFieldID))
            End If

            If Not objFilter.AppliesToTypeID = clsDBConstants.cintNULL Then
                colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.Filter.cAPPLIESTOTYPEID), _
                    objFilter.AppliesToTypeID))
            End If

            Dim objDT As DataTable = objDB.GetDataTable(clsDBConstants.Tables.cFILTER & _
                clsDBConstants.StoredProcedures.cGETLIST, colParams)

            colObjects = New List(Of clsFilter)
            For Each objDR As DataRow In objDT.Rows
                colObjects.Add(New clsFilter(objDR, objDB))
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Shared Function ConvertToFilter(ByVal objFieldLink As clsFieldLink,
                                           Optional ByVal colFields As clsMaskFieldDictionary = Nothing,
                                           Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As clsSearchFilter

        If intTypeID = clsDBConstants.cintNULL AndAlso
            colFields IsNot Nothing AndAlso
            objFieldLink.IdentityTable.TypeDependent Then

            intTypeID = CInt(colFields.GetMaskValue(clsDBConstants.Fields.cTYPEID, clsDBConstants.cintNULL))

        End If

        Dim objTable As clsTable

        If objFieldLink.ForeignKeyTable Is Nothing OrElse
            objFieldLink.ForeignKeyTable.IsLinkTable Then

            objTable = objFieldLink.LinkedTable

        Else

            objTable = objFieldLink.ForeignKeyTable

        End If


        Return CreateSearchFilter(objFieldLink.IdentityTable,
                                  objTable,
                                  objFieldLink.Filters(intTypeID),
                                  colFields)

    End Function

    Public Shared Function ConvertToFilter(ByVal objField As clsField,
                                           Optional ByVal colFields As clsMaskFieldDictionary = Nothing,
                                           Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As clsSearchFilter

        If intTypeID = clsDBConstants.cintNULL AndAlso
            colFields IsNot Nothing AndAlso
            objField.Table.TypeDependent Then

            intTypeID = CInt(colFields.GetMaskValue(clsDBConstants.Fields.cTYPEID,
                                                    clsDBConstants.cintNULL))

        End If

        Return CreateSearchFilter(objField.Table,
                                  objField.FieldLink.IdentityTable,
                                  objField.Filters(intTypeID),
                                  colFields)
    End Function

    Private Shared Function CreateSearchFilter(ByVal objFromTable As clsTable,
                                               ByVal objFilteredTable As clsTable,
                                               ByVal colFilters As List(Of clsFilter),
                                               Optional ByVal colFields As clsMaskFieldDictionary = Nothing) As clsSearchFilter

        Dim objSF As clsSearchFilter = Nothing
        Dim intDefaultTypeID As Integer = 0
        Dim intTypeID As Integer = 0

        If Not colFilters Is Nothing Then
            Dim objSG As New clsSearchGroup(clsSearchFilter.enumOperatorType.NONE)
            Dim colSOs As New List(Of clsSearchObjBase)
            Dim objDB As clsDB = objFromTable.Database

            objSG.SearchObjs = colSOs

            For Each objFilter As clsFilter In colFilters
                Dim objSE As clsSearchElement = Nothing
                Dim eOpType As clsSearchFilter.enumOperatorType
                Dim objFilterField As clsField

                If colSOs.Count = 0 Then
                    eOpType = clsSearchFilter.enumOperatorType.NONE
                Else
                    eOpType = clsSearchFilter.enumOperatorType.AND
                End If

                objFilterField = objFilter.FilterByField

                If objFilterField Is Nothing OrElse _
                objFilter.FilterType = clsDBConstants.cintNULL Then Continue For

                Select Case objFilter.FilterType
                    Case clsDBConstants.enumFilterTypes.PARENT_FIELD
                        Dim objParentField As clsField

                        objParentField = objFilter.FilterValueField
                        If objParentField Is Nothing OrElse _
                        colFields Is Nothing Then Continue For

                        objSE = New clsSearchElement(eOpType, _
                            objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                            clsSearchFilter.enumComparisonType.EQUAL, _
                            colFields.GetMaskValue(objParentField.DatabaseName))
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.USER_VALUE
                        If objFilterField.IsForeignKey AndAlso objFilter.FilterValue.IndexOf(",") >= 0 Then
                            objSE = New clsSearchElement(eOpType, _
                                objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                                clsSearchFilter.enumComparisonType.IN, objFilter.FilterValue)
                        Else
                            objSE = New clsSearchElement(eOpType, _
                                objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                                clsSearchFilter.enumComparisonType.EQUAL, objFilter.FilterValue)
                        End If
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.TABLE
                        objSE = New clsSearchElement(eOpType, _
                            objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                            clsSearchFilter.enumComparisonType.EQUAL, objFromTable.ID)
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.IS_NULL, _
                    clsDBConstants.enumFilterTypes.IS_NOT_NULL
                        If objFilter.FilterType = clsDBConstants.enumFilterTypes.IS_NOT_NULL Then
                            If eOpType = clsSearchFilter.enumOperatorType.NONE Then
                                eOpType = clsSearchFilter.enumOperatorType.NOT
                            Else
                                eOpType = clsSearchFilter.enumOperatorType.ANDNOT
                            End If
                        End If

                        objSE = New clsSearchElement(eOpType, _
                            objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                            clsSearchFilter.enumComparisonType.EQUAL, Nothing)
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.LOGGED_IN_PERSON
                        objSE = New clsSearchElement(eOpType, _
                            objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                            clsSearchFilter.enumComparisonType.EQUAL, objDB.Profile.PersonID)
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.LOGGED_IN_USERPROFILE
                        objSE = New clsSearchElement(eOpType, _
                            objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                            clsSearchFilter.enumComparisonType.EQUAL, objDB.Profile.ID)
                        colSOs.Add(objSE)

                    Case clsDBConstants.enumFilterTypes.TODAY
                        Dim objRangeSG As clsSearchGroup = clsSearchGroup.CreateRangeGroup(eOpType, _
                            objFilterField, Today, Today.AddDays(1).AddSeconds(-1), _
                            objFilterField.Table.DatabaseName)
                        colSOs.Add(objRangeSG)

                    Case clsDBConstants.enumFilterTypes.CURRENT_MONTH
                        Dim dtDate As Date = CType("1 " & Today.ToString("MMMM yyyy HH:mm:ss"), Date)
                        Dim objRangeSG As clsSearchGroup = clsSearchGroup.CreateRangeGroup(eOpType, _
                            objFilterField, dtDate, dtDate.AddMonths(1).AddSeconds(-1), _
                            objFilterField.Table.DatabaseName)
                        colSOs.Add(objRangeSG)

                    Case clsDBConstants.enumFilterTypes.CURRENT_YEAR
                        Dim dtDate As Date = CType("1 January " & Today.ToString("yyyy HH:mm:ss"), Date)
                        Dim objRangeSG As clsSearchGroup = clsSearchGroup.CreateRangeGroup(eOpType, _
                            objFilterField, dtDate, dtDate.AddYears(1).AddSeconds(-1), _
                            objFilterField.Table.DatabaseName)
                        colSOs.Add(objRangeSG)

                    Case clsDBConstants.enumFilterTypes.IS_DOCUMENT_PROFILE, _
                    clsDBConstants.enumFilterTypes.IS_FILE_FOLDER, _
                    clsDBConstants.enumFilterTypes.IS_ARCHIVE_BOX
                        Dim colIDs As Hashtable = Nothing

                        Select Case objFilter.FilterType
                            Case clsDBConstants.enumFilterTypes.IS_DOCUMENT_PROFILE
                                colIDs = GetTypeIDs(objDB, clsDBConstants.enumMDPTypeCodes.DocumentProfile)
                                intDefaultTypeID = objDB.SysInfo.K1Groups.GetDefaultType(clsDBConstants.enumMDPTypeCodes.DocumentProfile)
                            Case clsDBConstants.enumFilterTypes.IS_FILE_FOLDER
                                colIDs = GetTypeIDs(objDB, clsDBConstants.enumMDPTypeCodes.FileFolder)
                                intDefaultTypeID = objDB.SysInfo.K1Groups.GetDefaultType(clsDBConstants.enumMDPTypeCodes.FileFolder)
                            Case clsDBConstants.enumFilterTypes.IS_ARCHIVE_BOX
                                colIDs = GetTypeIDs(objDB, clsDBConstants.enumMDPTypeCodes.ArchiveBox)
                                intDefaultTypeID = objDB.SysInfo.K1Groups.GetDefaultType(clsDBConstants.enumMDPTypeCodes.ArchiveBox)
                        End Select

                        If colIDs IsNot Nothing AndAlso colIDs.Count > 0 Then
                            If colIDs.Count = 1 Then
                                intTypeID = GetSingleValue(colIDs)
                                objSE = New clsSearchElement(eOpType, _
                                    objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                                    clsSearchFilter.enumComparisonType.EQUAL, intTypeID)
                                colSOs.Add(objSE)
                            Else
                                objSE = New clsSearchElement(eOpType, _
                                     objFilterField.Table.DatabaseName & "." & objFilterField.DatabaseName, _
                                     clsSearchFilter.enumComparisonType.IN, colIDs)
                                colSOs.Add(objSE)
                            End If
                        End If

                End Select
            Next

            If objSG.SearchObjs.Count > 0 Then
                objSF = New clsSearchFilter(objDB, objSG, objFilteredTable.DatabaseName)
                objSF.TypeID = intTypeID
                objSF.DefaultTypeID = intDefaultTypeID
            End If

        End If

        Return objSF
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cFILTER), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cFILTERBYFIELDID, m_intFilterByFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cFILTERVALUEFIELDID, m_intValueFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cFILTERTYPE, m_eFilterType)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cFILTERVALUE, m_strFilterValue)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cINITFILTERFIELDID, m_intInitFilterFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cINITFILTERFIELDLINKID, m_intInitFilterFieldLinkID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Filter.cAPPLIESTOTYPEID, m_intAppliesToTypeID)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

End Class
