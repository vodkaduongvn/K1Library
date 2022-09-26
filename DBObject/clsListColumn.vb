Public Class clsListColumn
    Inherits clsDBObjBase

#Region " Members "

    Private m_intFieldID As Integer
    Private m_intAppliesToTypeID As Integer
    Private m_intWidth As Integer
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_intSortOrder As Integer
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intFieldID As Integer, _
    ByVal intAppliesToTypeID As Integer, _
    ByVal intWidth As Integer, _
    ByVal intCaptionID As Integer, _
    ByVal intSortOrder As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intFieldID = intFieldID
        m_intAppliesToTypeID = intAppliesToTypeID
        m_intWidth = intWidth
        m_intCaptionID = intCaptionID
        m_intSortOrder = intSortOrder
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ListColumn.cFIELDID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ListColumn.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
        m_intWidth = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ListColumn.cWIDTH, clsDBConstants.cintNULL), Integer)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ListColumn.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ListColumn.cSORTORDER, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property FieldID() As Integer
        Get
            Return m_intFieldID
        End Get
    End Property

    Public ReadOnly Property Field() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intFieldID)
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property

    Public ReadOnly Property Width() As Integer
        Get
            Return m_intWidth
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

    Public ReadOnly Property CaptionID() As Integer
        Get
            Return m_intCaptionID
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsListColumn
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cLISTCOLUMN, intID)

            Return New clsListColumn(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsListColumn)
        Dim colItems As FrameworkCollections.K1Dictionary(Of clsListColumn)
        Dim objItem As clsListColumn

        Try
            Dim strSP As String = clsDBConstants.Tables.cLISTCOLUMN & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colItems = New FrameworkCollections.K1Dictionary(Of clsListColumn)
            For Each objDR As DataRow In objDT.Rows
                objItem = New clsListColumn(objDR, objDB)
                colItems.Add(CStr(objItem.ID), objItem)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colItems
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cLISTCOLUMN), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ListColumn.cFIELDID, m_intFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ListColumn.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ListColumn.cWIDTH, m_intWidth)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ListColumn.cCAPTIONID, m_intCaptionID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ListColumn.cSORTORDER, m_intSortOrder)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
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
