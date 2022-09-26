#Region " File Information "

'=====================================================================
' This class represents the table Period in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsPeriod
    Inherits clsDBObjBase

#Region " Members "

    Private m_strCode As String
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_strCode = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Period.cCODE, clsDBConstants.cstrNULL), String)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Code() As String
        Get
            Return m_strCode
        End Get
    End Property
#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsPeriod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cPERIOD, intID)

            Return New clsPeriod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
