#Region " File Information "

'=====================================================================
' This class represents the table ErrorMessage in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsErrorMessage
	Inherits clsDBObjBase

#Region " Members "

	Private m_intStringID As Integer
    Private m_objString As clsString
    Private m_intUIID As Integer
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intStringID As Integer, _
    ByVal intUIID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intStringID = intStringID
        m_intUIID = intUIID
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intStringID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ErrorMessage.cSTRINGID, clsDBConstants.cintNULL), Integer)
        m_intUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ErrorMessage.cUIID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property StringObj() As clsString
        Get
            If m_objString Is Nothing Then
                If Not m_intStringID = clsDBConstants.cintNULL Then
                    m_objString = clsString.GetItem(m_intStringID, Me.Database)
                End If
            End If
            Return m_objString
        End Get
    End Property

    Public ReadOnly Property UIID() As Integer
        Get
            Return m_intUIID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsErrorMessage
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cERRORMESSAGE, intID)

            Return New clsErrorMessage(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsErrorMessage)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsErrorMessage)

        Try
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cERRORMESSAGE)
            Dim objDT As DataTable = objDB.GetDataTable(objTable)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsErrorMessage)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsErrorMessage(objDR, objDB)
                colObjects.Add(CStr(objItem.UIID), objItem)
            Next

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Shared Function GetItemByUIID(ByVal intUIID As Integer, ByVal objDB As clsDB) As clsErrorMessage
        Dim objDT As DataTable = objDB.GetDataTableByField( _
            clsDBConstants.Tables.cERRORMESSAGE, clsDBConstants.Fields.ErrorMessage.cUIID, intUIID)

        If objDT.Rows.Count = 1 Then
            Return clsErrorMessage.GetItem(CType(objDT.Rows.Item(0).Item(clsDBConstants.Fields.cID), Integer), objDB)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetString(ByVal intLanguageID As Integer, _
    ByVal intDefaultLanguageID As Integer) As String
        Dim objStringObj As clsString = StringObj

        If Not objStringObj Is Nothing Then
            Return StringObj.GetLanguageString(intLanguageID, intDefaultLanguageID, False)
        Else
            Return ""
        End If
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cERRORMESSAGE), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ErrorMessage.cSTRINGID, m_intStringID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.ErrorMessage.cUIID, m_intUIID)

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
            If Not m_objString Is Nothing Then
                m_objString.Dispose()
                m_objString = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
