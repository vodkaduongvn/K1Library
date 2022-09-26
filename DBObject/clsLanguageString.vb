#Region " File Information "

'=====================================================================
' This class represents the table LinkLanguageString in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsLanguageString
    Inherits clsDBObjBase

#Region " Members "

    Private m_intStringID As Integer
    Private m_intLanguageID As Integer
    Private m_strText As String
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intLanguageID As Integer, _
    ByVal intStringID As Integer, _
    ByVal strText As String)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intStringID = intStringID
        m_intLanguageID = intLanguageID
        m_strText = strText
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB, ByVal blnPermanent As Boolean)
        MyBase.New(objDR, objDB)
        m_intStringID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.LanguageString.cSTRINGID, clsDBConstants.cintNULL), Integer)
        m_intLanguageID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.LanguageString.cLANGUAGEID, clsDBConstants.cintNULL), Integer)

        If blnPermanent Then
            m_strText = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.LanguageString.cSTRING, clsDBConstants.cstrNULL), String)
            If m_strText.Length > 0 Then
                m_strText = m_strText.Substring(0, Math.Min(100, m_strText.Length))
            End If
        End If
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Text() As String
        Get
            If m_strText Is Nothing Then
                Dim objDT As DataTable = m_objDB.GetItem(clsDBConstants.Tables.cLANGUAGESTRING, m_intID)
                If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
                    Return ""
                Else
                    Return CType(clsDB_Direct.DataRowValue(objDT.Rows(0), _
                        clsDBConstants.Fields.LanguageString.cSTRING, clsDBConstants.cstrNULL), String)
                End If
            Else
                Return m_strText
            End If
        End Get
    End Property

    Public ReadOnly Property StringID() As Integer
        Get
            Return m_intStringID
        End Get
    End Property

    Public ReadOnly Property LanguageID() As Integer
        Get
            Return m_intLanguageID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB, _
    Optional ByVal blnPermanent As Boolean = False) As clsLanguageString
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cLANGUAGESTRING, intID)

            Return New clsLanguageString(objDT.Rows(0), objDB, blnPermanent)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objString As clsString, ByVal objDB As clsDB, _
    Optional ByVal blnPermanent As Boolean = False) As FrameworkCollections.K1Dictionary(Of clsLanguageString)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsLanguageString)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cLANGUAGESTRING, _
                clsDBConstants.Fields.LanguageString.cSTRINGID, objString.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsLanguageString)
            For Each objDR As DataRow In objDT.Rows
                Dim objLinkLanguageString As New clsLanguageString(objDR, objDB, blnPermanent)
                colObjects.Add(CType(objLinkLanguageString.m_intLanguageID, String), objLinkLanguageString)
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

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrail As Boolean = True)
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cLANGUAGESTRING), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.LanguageString.cLANGUAGEID, m_intLanguageID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.LanguageString.cSTRINGID, m_intStringID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.LanguageString.cSTRING, m_strText)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB, blnCreateAuditTrail:=blnCreateAuditTrail)
        Else
            colMasks.Update(m_objDB, blnCreateAuditTrail:=blnCreateAuditTrail)
        End If

        m_strText = Nothing 'clear the string from memory
    End Sub
#End Region

End Class
