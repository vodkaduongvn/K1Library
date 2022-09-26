Public Class clsFieldLink
    Inherits clsDBObjBase

#Region " Members "

    Private m_intPrimaryFieldID As Integer
    Private m_intForeignFieldID As Integer
    Private m_blnIsVisible As Boolean
    Private m_intSortOrder As Integer
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_colTypeFieldLinks As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
    Private m_blnDisplayAsDropDown As Boolean = False
    Private m_blnIsExpanded As Boolean = False
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intPrimaryFieldID As Integer, _
    ByVal intForeignFieldID As Integer, _
    ByVal blnIsVisible As Boolean, _
    ByVal intSortOrder As Integer, _
    ByVal intCaptionID As Integer, _
    ByVal blnDisplayAsDropDown As Boolean, _
    ByVal blnExpanded As Boolean)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intPrimaryFieldID = intPrimaryFieldID
        m_intForeignFieldID = intForeignFieldID
        m_blnIsVisible = blnIsVisible
        m_intSortOrder = intSortOrder
        m_intCaptionID = intCaptionID
        m_blnDisplayAsDropDown = blnDisplayAsDropDown
        m_blnIsExpanded = blnExpanded
    End Sub
#End Region

#Region " From Database "

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intPrimaryFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cPRIMARYFIELDID, clsDBConstants.cintNULL), Integer)
        m_intForeignFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cFOREIGNFIELDID, clsDBConstants.cintNULL), Integer)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cISVISIBLE, False), Boolean)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cSORTORDER, clsDBConstants.cintNULL), Integer)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_blnDisplayAsDropDown = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cDISPLAYASDROPDOWN, False), Boolean)
        m_blnIsExpanded = CType(clsDB.DataRowValue(objDR, clsDBConstants.Fields.FieldLink.cISEXPANDED, False), Boolean)
    End Sub

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB, ByVal blnWSImproved As Boolean)
        m_intID = CType(clsDB_Direct.DataRowValue(objDR, "A", clsDBConstants.cintNULL), Integer)
        m_strExternalID = CType(clsDB_Direct.DataRowValue(objDR, "C", clsDBConstants.cstrNULL), String)
        m_intSecurityID = CType(clsDB_Direct.DataRowValue(objDR, "B", clsDBConstants.cintNULL), Integer)
        m_intTypeID = CType(clsDB_Direct.DataRowValue(objDR, "D", clsDBConstants.cintNULL), Integer)
        m_objDB = objDB
        m_intPrimaryFieldID = CType(clsDB_Direct.DataRowValue(objDR, "E", clsDBConstants.cintNULL), Integer)
        m_intForeignFieldID = CType(clsDB_Direct.DataRowValue(objDR, "F", clsDBConstants.cintNULL), Integer)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, "G", False), Boolean)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, "H", clsDBConstants.cintNULL), Integer)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, "I", clsDBConstants.cintNULL), Integer)
        m_blnDisplayAsDropDown = CType(clsDB.DataRowValue(objDR, "J", False), Boolean)
        m_blnIsExpanded = CType(clsDB.DataRowValue(objDR, "K", False), Boolean)
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property IdentityFieldID() As Integer
        Get
            Return m_intPrimaryFieldID
        End Get
    End Property

    Public ReadOnly Property IdentityField() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intPrimaryFieldID)
        End Get
    End Property

    Public ReadOnly Property ForeignKeyFieldID() As Integer
        Get
            Return m_intForeignFieldID
        End Get
    End Property

    Public ReadOnly Property ForeignKeyField() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intForeignFieldID)
        End Get
    End Property

    Public ReadOnly Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property

    Public ReadOnly Property CaptionText() As String
        Get
            Return m_objDB.SysInfo.GetCaptionString(m_intCaptionID)
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

    ''' <summary>
    ''' The Table the Foreign Key field is pointing to.
    ''' </summary>
    Public ReadOnly Property IdentityTable() As clsTable
        Get
            Return m_objDB.SysInfo.Fields(m_intPrimaryFieldID).Table
        End Get
    End Property

    Public ReadOnly Property IsLinkTableRelated() As Boolean
        Get
            Dim objTable As clsTable = ForeignKeyTable

            If objTable IsNot Nothing AndAlso objTable.IsLinkTable Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>
    ''' This will return the other field link (foreign key) of the link table
    ''' </summary>
    Public ReadOnly Property LinkTableOppositeFieldLink() As clsFieldLink
        Get
            Dim objTable As clsTable = ForeignKeyTable

            If objTable IsNot Nothing AndAlso objTable.IsLinkTable Then
                Dim objForeignFieldLink As clsFieldLink = Nothing

                For Each objFieldLink As clsFieldLink In objTable.FieldLinks.Values
                    If Not objFieldLink.ID = Me.ID Then
                        objForeignFieldLink = objFieldLink
                        Exit For
                    End If
                Next

                If objForeignFieldLink IsNot Nothing Then
                    Return objForeignFieldLink
                End If
            End If

            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' If the foreign key table is a link table link, 
    ''' this will return the other table in the many-to-many relationship
    ''' </summary>
    Public Overridable ReadOnly Property LinkedTable() As clsTable
        Get
            Dim objFieldLink As clsFieldLink = LinkTableOppositeFieldLink

            If objFieldLink IsNot Nothing Then
                Return objFieldLink.IdentityTable
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' This will be the table that contains the foreign key.
    ''' </summary>
    Public ReadOnly Property ForeignKeyTable() As clsTable
        Get
            Dim objTable As clsTable = Nothing
            Dim objField As clsField

            objField = m_objDB.SysInfo.Fields(m_intForeignFieldID)

            If Not objField Is Nothing Then
                objTable = objField.Table
            End If

            Return objTable
        End Get
    End Property

    Public ReadOnly Property TypeFieldLinkInfos() As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
        Get
            If m_colTypeFieldLinks Is Nothing Then
                m_colTypeFieldLinks = New FrameworkCollections.K1Dictionary(Of clsTypeFieldLink) 'clsTypeFieldLink.GetList(Me, m_objDB)
            End If
            Return m_colTypeFieldLinks
        End Get
    End Property

    Public ReadOnly Property DisplayAsDropDown() As Boolean
        Get
            Return m_blnDisplayAsDropDown
        End Get
    End Property

    Public ReadOnly Property IsExpanded() As Boolean
        Get
            Return m_blnIsExpanded
        End Get
    End Property

    Public ReadOnly Property Filters(ByVal intTypeID As Integer) As List(Of clsFilter)
        Get
            Dim colFilters As List(Of clsFilter) = Nothing

            If Not intTypeID = clsDBConstants.cintNULL Then
                colFilters = m_objDB.SysInfo.Filters("L_" & m_intID & "_" & intTypeID)
            End If

            If colFilters Is Nothing Then
                colFilters = m_objDB.SysInfo.Filters("L_" & m_intID)
            End If

            Return colFilters
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsFieldLink
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cFIELDLINK, intID)

            Return New clsFieldLink(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
        Dim objDT As DataTable
        Dim strSP As String
        Dim colFieldLinks As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)

        Try
            strSP = clsDBConstants.Tables.cFIELDLINK & clsDBConstants.StoredProcedures.cGETLIST

            objDT = objDB.GetDataTable(strSP)

            colFieldLinks = New FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objFieldLink As New clsFieldLink(objDataRow, objDB)
                colFieldLinks.Add(CStr(objFieldLink.ForeignKeyFieldID), objFieldLink)
            Next

            Return colFieldLinks
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList2(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
        Dim objDT As DataTable
        Dim strSP As String
        Dim colFieldLinks As FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)

        Try
            strSP = "SELECT [FieldLink].[ID] AS 'A', [FieldLink].[SecurityID] AS 'B', " & _
                "[FieldLink].[ExternalID] AS 'C', [FieldLink].[TypeID] AS 'D', [FieldLink].[PrimaryFieldID] AS 'E', " & _
                "[FieldLink].[ForeignFieldID] AS 'F', [FieldLink].[IsVisible] AS 'G', [FieldLink].[SortOrder] AS 'H', " & _
                "[FieldLink].[CaptionID] AS 'I', [FieldLink].[DisplayAsDropDown] AS 'J', [FieldLink].[isExpanded] AS 'K'" & _
                "FROM [FieldLink]"

            objDT = objDB.GetDataTableBySQL(strSP)

            colFieldLinks = New FrameworkCollections.K1DualKeyDictionary(Of clsFieldLink, Integer)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objFieldLink As New clsFieldLink(objDataRow, objDB, True)
                colFieldLinks.Add(CStr(objFieldLink.ForeignKeyFieldID), objFieldLink)
            Next

            Return colFieldLinks
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate(Optional blnCreateAuditTrail As Boolean = True)

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(m_objDB.SysInfo.Tables(clsDBConstants.Tables.cFIELDLINK),
                                                                                   m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cPRIMARYFIELDID, m_intPrimaryFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cFOREIGNFIELDID, m_intForeignFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cISVISIBLE, m_blnIsVisible)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cSORTORDER, m_intSortOrder)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cCAPTIONID, m_intCaptionID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cDISPLAYASDROPDOWN, m_blnDisplayAsDropDown)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.FieldLink.cISEXPANDED, m_blnIsExpanded)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB, blnCreateAuditTrail)
        End If

    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_objCaption IsNot Nothing Then
                m_objCaption.Dispose()
                m_objCaption = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
