Public Class clsDRMMethod
    Inherits clsDBObjBase

#Region " Members "
    Dim m_intDRMMethodID As Integer
    Dim m_intMethodID As Integer
#End Region

#Region " Properties "
    Public ReadOnly Property DRMMethodID() As Integer
        Get
            Return m_intDRMMethodID
        End Get
    End Property

    Public ReadOnly Property MethodID() As Integer
        Get
            Return m_intMethodID
        End Get
    End Property
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub


    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intDRMMethodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.LinkUserProfileDRMMethod.cDRMMETHODID, clsDBConstants.cintNULL), Integer)
        m_intMethodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TableMethod.cMETHODID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsDRMMethod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cDRMMethod, intID)

            Return New clsDRMMethod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsDRMMethod)
        Dim colDRMMethods As FrameworkCollections.K1Dictionary(Of clsDRMMethod)

        Try
            Dim strSP As String = clsDBConstants.Tables.cDRMMethod & _
                clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colDRMMethods = New FrameworkCollections.K1Dictionary(Of clsDRMMethod)

            For Each objDR As DataRow In objDT.Rows
                Dim objDRMMethod As New clsDRMMethod(objDR, objDB)
                colDRMMethods.Add(CStr(objDRMMethod.ID), objDRMMethod)
            Next

            Return colDRMMethods
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
