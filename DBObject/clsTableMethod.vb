#Region " File Information "

'=====================================================================
' This class represents the table LinkTableMethod in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsTableMethod
    Inherits clsDBObjBase

#Region " Members "

    Private m_intTableID As Integer
    Private m_intMethodID As Integer
    Private m_blnAudit As Boolean
    Private m_blnAuditData As Boolean
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_intTableID = clsDBConstants.cintNULL
        m_intMethodID = clsDBConstants.cintNULL
        m_blnAudit = False
        m_blnAuditData = False
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TableMethod.cTABLEID, clsDBConstants.cintNULL), Integer)
        m_intMethodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TableMethod.cMETHODID, clsDBConstants.cintNULL), Integer)
        m_blnAudit = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TableMethod.cAUDIT, False), Boolean)
        m_blnAuditData = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TableMethod.cAUDITDATA, False), Boolean)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Audit() As Boolean
        Get
            Return m_blnAudit
        End Get
    End Property

    Public ReadOnly Property AuditData() As Boolean
        Get
            Return m_blnAuditData
        End Get
    End Property

    Public ReadOnly Property MethodID() As Integer
        Get
            Return m_intMethodID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsTableMethod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cTABLEMETHOD, intID)

            Return New clsTableMethod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objTable As clsTable, ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTableMethod)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTableMethod)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cTABLEMETHOD, _
                clsDBConstants.Fields.TableMethod.cTABLEID, objTable.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTableMethod)
            For Each objDR As DataRow In objDT.Rows
                Dim objLinkTableMethod As New clsTableMethod(objDR, objDB)
                colObjects.Add(CType(objLinkTableMethod.m_intMethodID, String), objLinkTableMethod)
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

End Class
