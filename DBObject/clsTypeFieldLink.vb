#Region " File Information "

'=====================================================================
' This class represents the table TypeFieldLinkInfo in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date            Description
'---------------------------------------------------------------------
' KD        07/09/2005      Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsTypeFieldLink
    Inherits clsDBObjBase

#Region " Members "

    Private m_intFieldLinkID As Integer
    Private m_blnIsVisible As Boolean
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_intAppliesToTypeID As Integer
    Private m_intSortOrder As Integer
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_blnIsVisible = True
        m_intCaptionID = clsDBConstants.cintNULL
        m_intAppliesToTypeID = clsDBConstants.cintNULL
        m_intSortOrder = 5
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldLinkInfo.cISVISIBLE, True), Boolean)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldLinkInfo.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldLinkInfo.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldLinkInfo.cSORTORDER, clsDBConstants.cintNULL), Integer)
        m_intFieldLinkID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldLinkInfo.cFIELDLINKID, clsDBConstants.cintNULL), Integer)
    End Sub

    Public Sub New(ByVal objFieldLink As clsFieldLink, ByVal intID As Integer, ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, ByVal intAppliesToTypeID As Integer, ByVal blnVisible As Boolean, _
    ByVal intCaptionID As Integer, ByVal intSortOrder As Integer)
        MyBase.New(objFieldLink.Database, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intAppliesToTypeID = intAppliesToTypeID
        m_intCaptionID = intCaptionID
        m_intFieldLinkID = objFieldLink.ID
        m_intSortOrder = intSortOrder
        m_blnIsVisible = blnVisible
    End Sub
#End Region

#Region " Properties "

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

    Public ReadOnly Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property

    Public ReadOnly Property FieldLinkID() As Integer
        Get
            Return m_intFieldLinkID
        End Get
    End Property
#End Region

#Region "Methods"

    Public Function isInType(intType As Integer) As Boolean

        Return intType = AppliesToTypeID()
    End Function

#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsTypeFieldLink
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cTYPEFIELDLINKINFO, intID)

            Return New clsTypeFieldLink(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
        Dim strSP As String

        Try
            strSP = clsDBConstants.Tables.cTYPEFIELDLINKINFO & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
            For Each objDR As DataRow In objDT.Rows
                Dim objTypeFieldLinkInfo As New clsTypeFieldLink(objDR, objDB)
                colObjects.Add(objTypeFieldLinkInfo.FieldLinkID & "_" & CType(objTypeFieldLinkInfo.m_intAppliesToTypeID, String), objTypeFieldLinkInfo)
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

    Public Shared Function GetList(ByVal objFieldLink As clsFieldLink, _
    ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cTYPEFIELDLINKINFO, _
                clsDBConstants.Fields.TypeFieldLinkInfo.cFIELDLINKID, objFieldLink.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTypeFieldLink)
            For Each objDR As DataRow In objDT.Rows
                Dim objTypeFieldLinkInfo As New clsTypeFieldLink(objDR, objDB)
                colObjects.Add(CType(objTypeFieldLinkInfo.m_intAppliesToTypeID, String), objTypeFieldLinkInfo)
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

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTYPEFIELDLINKINFO), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldLinkInfo.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldLinkInfo.cFIELDLINKID, m_intFieldLinkID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldLinkInfo.cISVISIBLE, m_blnIsVisible)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldLinkInfo.cSORTORDER, m_intSortOrder)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldLinkInfo.cCAPTIONID, m_intCaptionID)

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
            If Not m_objCaption Is Nothing Then
                m_objCaption.Dispose()
                m_objCaption = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
