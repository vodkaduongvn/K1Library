#Region " File Information "

'=====================================================================
' This class represents the table String in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsString
	Inherits clsDBObjBase

#Region " Members "

    Private m_colLinkLanguageString As FrameworkCollections.K1Dictionary(Of clsLanguageString)
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property LinkLanguageStrings() As FrameworkCollections.K1Dictionary(Of clsLanguageString)
        Get
            If m_colLinkLanguageString Is Nothing Then
                m_colLinkLanguageString = clsLanguageString.GetList(Me, Me.Database)
            End If
            Return m_colLinkLanguageString
        End Get
    End Property
#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsString
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cSTRING, intID)

            Return New clsString(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Function GetLanguageString(ByVal intLanguageID As Integer, _
    ByVal intDefaultLanguageID As Integer, Optional ByVal blnStoreInMemory As Boolean = False) As String
        Dim strText As String = ""

        If m_colLinkLanguageString Is Nothing Then
            m_colLinkLanguageString = clsLanguageString.GetList( _
                Me, Me.Database, blnStoreInMemory)
        End If

        Dim objLanguageString As clsLanguageString
        objLanguageString = m_colLinkLanguageString(CType(intLanguageID, String))

        If Not objLanguageString Is Nothing Then
            strText = objLanguageString.Text
        Else
            objLanguageString = m_colLinkLanguageString(CType(intDefaultLanguageID, String))
            If Not objLanguageString Is Nothing Then
                strText = objLanguageString.Text
            End If
        End If

        Return strText
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrail As Boolean = True)
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cSTRING), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)

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
            If m_colLinkLanguageString IsNot Nothing Then
                m_colLinkLanguageString.Dispose()
                m_colLinkLanguageString = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
