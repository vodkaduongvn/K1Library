#Region " File Information "

'=====================================================================
' This class represents the table Warning in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' KD        1/04/2008    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsWarningMessage
    Inherits clsDBObjBase

#Region " Members "

    Private m_strStoredProcedure As String
    Private m_intTableID As Integer
    Private m_eWarningType As enumWarningType
    Private m_strSQL As String
#End Region

#Region " Enumerations "

    Public Enum enumWarningType
        INSERT_ONLY = 0
        MODIFY_ONLY = 1
        INSERT_AND_MODIFY = 2
        DELETE_ONLY = 3
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal strStoredProcedure As String, _
    ByVal strSQL As String, _
    ByVal eWarningType As enumWarningType, _
    ByVal intTableID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_strStoredProcedure = strStoredProcedure
        m_strSQL = strSQL
        m_eWarningType = eWarningType
        m_intTableID = intTableID
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_strStoredProcedure = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Warning.cSTOREDPROCEDURENAME, clsDBConstants.cstrNULL), String)
        m_strSQL = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Warning.cSQL, clsDBConstants.cstrNULL), String)
        m_eWarningType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Warning.cWARNINGTYPE, enumWarningType.INSERT_AND_MODIFY), enumWarningType)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Warning.cTABLEID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property StoredProcedure() As String
        Get
            Return m_strStoredProcedure
        End Get
    End Property

    Public ReadOnly Property SQL() As String
        Get
            Return m_strSQL
        End Get
    End Property

    Public ReadOnly Property WarningType() As enumWarningType
        Get
            Return m_eWarningType
        End Get
    End Property

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objDB.SysInfo.Tables(m_intTableID)
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsWarningMessage
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cWARNING, intID)

            Return New clsWarningMessage(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objTable As clsTable, ByVal objDB As clsDB) As FrameworkCollections.K1Collection(Of clsWarningMessage)
        Dim colObjects As FrameworkCollections.K1Collection(Of clsWarningMessage)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cWARNING, _
                clsDBConstants.Fields.Warning.cTABLEID, objTable.ID)

            colObjects = New FrameworkCollections.K1Collection(Of clsWarningMessage)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsWarningMessage(objDR, objDB)
                colObjects.Add(objItem)
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

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsWarningMessage)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsWarningMessage)

        Try
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cWARNING)
            Dim objDT As DataTable = objDB.GetDataTable(objTable)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsWarningMessage)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsWarningMessage(objDR, objDB)
                colObjects.Add(CStr(objItem.StoredProcedure), objItem)
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
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cWARNING), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Warning.cSTOREDPROCEDURENAME, m_strStoredProcedure)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Warning.cSQL, m_strSQL)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Warning.cTABLEID, m_intTableID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Warning.cWARNINGTYPE, m_eWarningType)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

End Class
