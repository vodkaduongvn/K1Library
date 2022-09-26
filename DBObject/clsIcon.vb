#Region " File Information "

'=====================================================================
' This class represents the table Icon in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsIcon
	Inherits clsDBObjBase

#Region " Members "

    Private m_intUpEDOCID As Integer
    Private m_objUpEDOC As clsEDOC
    Private m_intHoverEDOCID As Integer
    Private m_objHoverEDOC As clsEDOC
    Private m_intDownEDOCID As Integer
    Private m_objDownEDOC As clsEDOC
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intUpEDOCID As Integer, _
    ByVal intHoverEDOCID As Integer, _
    ByVal intDownEDOCID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intUpEDOCID = intUpEDOCID
        m_intHoverEDOCID = intHoverEDOCID
        m_intDownEDOCID = intDownEDOCID
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intUpEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Icon.cEDOCID, clsDBConstants.cintNULL), Integer)
        m_intHoverEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Icon.cOVEREDOCID, clsDBConstants.cintNULL), Integer)
        m_intDownEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Icon.cDOWNEDOCID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property UpEDOC() As clsEDOC
        Get
            If m_objUpEDOC Is Nothing Then
                If Not m_intUpEDOCID = clsDBConstants.cintNULL Then
                    m_objUpEDOC = clsEDOC.GetItem(m_intUpEDOCID, Me.Database)
                End If
            End If
            Return m_objUpEDOC
        End Get
    End Property

    Public ReadOnly Property HoverEDOC() As clsEDOC
        Get
            If m_objHoverEDOC Is Nothing Then
                If Not m_intHoverEDOCID = clsDBConstants.cintNULL Then
                    m_objHoverEDOC = clsEDOC.GetItem(m_intHoverEDOCID, Me.Database)
                End If
            End If
            Return m_objHoverEDOC
        End Get
    End Property

    Public ReadOnly Property DownEDOC() As clsEDOC
        Get
            If m_objDownEDOC Is Nothing Then
                If Not m_intDownEDOCID = clsDBConstants.cintNULL Then
                    m_objDownEDOC = clsEDOC.GetItem(m_intDownEDOCID, Me.Database)
                End If
            End If
            Return m_objDownEDOC
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsIcon
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cICON, intID)

            Return New clsIcon(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsIcon)
        Dim objDT As DataTable
        Dim colObjs As FrameworkCollections.K1Dictionary(Of clsIcon)

        Try
            objDT = objDB.GetDataTable(clsDBConstants.Tables.cICON & clsDBConstants.StoredProcedures.cGETLIST)

            colObjs = New FrameworkCollections.K1Dictionary(Of clsIcon)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objItem As New clsIcon(objDataRow, objDB)
                colObjs.Add(CType(objItem.ID, String), objItem)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjs
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cICON), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Icon.cEDOCID, m_intUpEDOCID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Icon.cDOWNEDOCID, m_intDownEDOCID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Icon.cOVEREDOCID, m_intHoverEDOCID)

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
            If m_objUpEDOC IsNot Nothing Then
                m_objUpEDOC.Dispose()
                m_objUpEDOC = Nothing
            End If

            If m_objHoverEDOC IsNot Nothing Then
                m_objHoverEDOC.Dispose()
                m_objHoverEDOC = Nothing
            End If

            If m_objDownEDOC IsNot Nothing Then
                m_objDownEDOC.Dispose()
                m_objDownEDOC = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
