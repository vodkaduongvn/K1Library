Public Class clsDBObject
    Inherits clsDBObjBase

#Region " Members "

    Private m_intTableID As Integer
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
    End Sub

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB, _
    ByVal objTable As clsTable)
        MyBase.New(objDR, objDB)
        m_intTableID = objTable.ID
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal intID As Integer, _
    ByVal strExternalID As String, ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer, ByVal objTable As clsTable)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, intTypeID)
        m_intTableID = objTable.ID
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objDB.SysInfo.Tables(m_intTableID)
        End Get
    End Property
#End Region

#Region " Methods "

#Region " GetItem "

    Public Shared Function GetItem(ByVal strTable As String, ByVal intID As Integer, ByVal objDB As clsDB) As clsDBObject
        Try
            Dim objDT As DataTable = objDB.GetItem(strTable, intID)

            If objDT.Rows.Count = 1 Then
                Return New clsDBObject(objDT.Rows(0), objDB)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#End Region

End Class
