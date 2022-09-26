#Region " File Information "

'=====================================================================
' This class represents the table Background in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsBackground
	Inherits clsDBObjBase

#Region " Members "

	Private m_intEDOCID As Integer
	Private m_objEDOC As clsEDOC
#End Region

#Region " Constructors "

	Public Sub New()
		MyBase.New()
        m_intEDOCID = clsDBConstants.cintNULL
	End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.new(objDR, objDB)
        m_intEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Background.cEDOCID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property EDOC() As clsEDOC
        Get
            If m_objEDOC Is Nothing Then
                If Not m_intEDOCID = clsDBConstants.cintNULL Then
                    m_objEDOC = clsEDOC.GetItem(m_intEDOCID, Me.Database)
                End If
            End If
            Return m_objEDOC
        End Get
    End Property

#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsBackground
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cBACKGROUND, intID)

            Return New clsBackground(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_objEDOC IsNot Nothing Then
                m_objEDOC.Dispose()
                m_objEDOC = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
