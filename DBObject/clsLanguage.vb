#Region " File Information "

'=====================================================================
' This class represents the table Language in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsLanguage
	Inherits clsDBObjBase

#Region " Constructors "

	Public Sub New()
		MyBase.New()
	End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
    End Sub
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsLanguage
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cLANGUAGE, intID)

            Return New clsLanguage(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsLanguage)
        Dim colLanguages As FrameworkCollections.K1Dictionary(Of clsLanguage)
        Dim objLanguage As clsLanguage

        Try
            Dim strSP As String = clsDBConstants.Tables.cLANGUAGE & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colLanguages = New FrameworkCollections.K1Dictionary(Of clsLanguage)
            For Each objDR As DataRow In objDT.Rows
                objLanguage = New clsLanguage(objDR, objDB)
                colLanguages.Add(CStr(objLanguage.ID), objLanguage)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colLanguages
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
