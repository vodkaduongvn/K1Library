#Region " File Information "

'=====================================================================
' This class represents the table Caption in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsCaption
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
        m_intStringID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Caption.cSTRINGID, clsDBConstants.cintNULL), Integer)
        m_intUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Caption.cUIID, clsDBConstants.cintNULL), Integer)
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

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsCaption
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cCAPTION, intID)

            If objDT.Rows.Count = 1 Then
                Return New clsCaption(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Function GetString(ByVal intLanguageID As Integer, _
    ByVal intDefaultLanguageID As Integer) As String
        Dim objStringObj As clsString = StringObj

        If Not objStringObj Is Nothing Then
            Return StringObj.GetLanguageString(intLanguageID, intDefaultLanguageID, True)
        Else
            Return ""
        End If
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrail As Boolean = True)
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cCAPTION), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Caption.cSTRINGID, m_intStringID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Caption.cUIID, m_intUIID)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB, blnCreateAuditTrail:=blnCreateAuditTrail)
        Else
            colMasks.Update(m_objDB, blnCreateAuditTrail:=blnCreateAuditTrail)
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
