#Region " File Information "

'=====================================================================
' This class represents the table Security in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsSecurity
	Inherits clsDBObjBase

#Region " Members "

    Private m_blnIsPublic As Boolean

#End Region

#Region " Constructors "

	Public Sub New()
		MyBase.New()
	End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_blnIsPublic = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Security.cISPUBLIC, False), Boolean)
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property IsPublic() As Boolean
        Get
            Return m_blnIsPublic
        End Get
    End Property

#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsSecurity
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cSECURITY, intID)

            Return New clsSecurity(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsSecurity)
        Dim colSecurities As FrameworkCollections.K1Dictionary(Of clsSecurity)

        Try
            Dim strSP As String = clsDBConstants.Tables.cSECURITY & _
                clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colSecurities = New FrameworkCollections.K1Dictionary(Of clsSecurity)
            For Each objDR As DataRow In objDT.Rows
                Dim objSecurity As New clsSecurity(objDR, objDB)
                colSecurities.Add(CStr(objSecurity.ID), objSecurity)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colSecurities
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class
