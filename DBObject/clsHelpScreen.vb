Public Class clsHelpScreen
    Inherits clsDBObjBase

#Region " Members "

    Private m_intLanguageID As Integer
    Private m_strHelpFile As String
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_strHelpFile = clsDBConstants.cstrNULL
        m_intLanguageID = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_strHelpFile = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.HelpScreen.cHELPFILE, clsDBConstants.cstrNULL), String)
        m_intLanguageID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.HelpScreen.cLANGUAGEID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property HelpFile() As String
        Get
            Return m_strHelpFile
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsHelpScreen
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cHELPSCREEN, intID)

            Return New clsHelpScreen(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsHelpScreen)
        Dim objDT As DataTable
        Dim colObjs As FrameworkCollections.K1Dictionary(Of clsHelpScreen)

        Try
            objDT = objDB.GetDataTable(clsDBConstants.Tables.cHELPSCREEN & clsDBConstants.StoredProcedures.cGETLIST)

            colObjs = New FrameworkCollections.K1Dictionary(Of clsHelpScreen)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objHS As New clsHelpScreen(objDataRow, objDB)
                colObjs.Add(CType(objHS.m_intLanguageID, String), objHS)
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

End Class
