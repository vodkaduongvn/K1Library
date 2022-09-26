#Region " File Information "

'=====================================================================
' This class represents the table Button in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsButton
	Inherits clsDBObjBase

#Region " Members "

	Private m_intStringID As Integer
	Private m_objString As clsString
    Private m_intUpEDOCID As Integer
    Private m_objUpEDOC As clsEDOC
    Private m_intHoverEDOCID As Integer
    Private m_objHoverEDOC As clsEDOC
    Private m_intDownEDOCID As Integer
    Private m_objDownEDOC As clsEDOC
    Private m_intUIID As Integer
#End Region

#Region " Constructors "

	Public Sub New()
		MyBase.New()
        m_intStringID = clsDBConstants.cintNULL
        m_intUpEDOCID = clsDBConstants.cintNULL
        m_intHoverEDOCID = clsDBConstants.cintNULL
        m_intDownEDOCID = clsDBConstants.cintNULL
        m_intUIID = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intStringID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Button.cSTRINGID, clsDBConstants.cintNULL), Integer)
        m_intUpEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Button.cEDOCID, clsDBConstants.cintNULL), Integer)
        m_intHoverEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Button.cOVEREDOCID, clsDBConstants.cintNULL), Integer)
        m_intDownEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Button.cDOWNEDOCID, clsDBConstants.cintNULL), Integer)
        m_intUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Button.cUIID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property StringObj() As clsString
        Get
            If m_objString Is Nothing Then
                If Not m_intStringID = clsDBConstants.cintNULL Then
                    m_objString = clsString.GetItem(m_intStringID, Me.Database)
                End If
            End If
            Return m_objString
        End Get
    End Property

    Public ReadOnly Property UpEDOC() As clsEDOC
        Get
            If m_objUpEDOC Is Nothing Then
                If Not m_intUpEDOCID = clsDBConstants.cintNULL Then
                    m_objUpEDOC = clsEDOC.GetItem(m_intUpEDOCID, Me.Database)
                End If
            End If
            Return m_objUpEDOC
        End Get
    End Property

    Public ReadOnly Property HoverEDOC() As clsEDOC
        Get
            If m_objHoverEDOC Is Nothing Then
                If Not m_intHoverEDOCID = clsDBConstants.cintNULL Then
                    m_objHoverEDOC = clsEDOC.GetItem(m_intHoverEDOCID, Me.Database)
                End If
            End If
            Return m_objHoverEDOC
        End Get
    End Property

    Public ReadOnly Property DownEDOC() As clsEDOC
        Get
            If m_objDownEDOC Is Nothing Then
                If Not m_intDownEDOCID = clsDBConstants.cintNULL Then
                    m_objDownEDOC = clsEDOC.GetItem(m_intDownEDOCID, Me.Database)
                End If
            End If
            Return m_objDownEDOC
        End Get
    End Property

    Public ReadOnly Property UIID() As Integer
        Get
            Return m_intUIID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsButton
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cBUTTON, intID)

            Return New clsButton(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As List(Of clsButton)
        Dim objDT As DataTable
        Dim strSP As String
        Dim colButtons As List(Of clsButton)

        Try
            strSP = clsDBConstants.Tables.cBUTTON & clsDBConstants.StoredProcedures.cGETLIST

            objDT = objDB.GetDataTable(strSP)

            colButtons = New List(Of clsButton)
            For Each objDataRow As DataRow In objDT.Rows
                colButtons.Add(New clsButton(objDataRow, objDB))
            Next

            Return colButtons
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

    Public Function GetString(ByVal intLanguageID As Integer, _
    ByVal intDefaultLanguageID As Integer) As String
        Dim objStringObj As clsString = StringObj

        If objStringObj IsNot Nothing Then
            Return StringObj.GetLanguageString(intLanguageID, intDefaultLanguageID, True)
        Else
            Return ""
        End If
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_objString IsNot Nothing Then
                m_objString.Dispose()
                m_objString = Nothing
            End If

            If m_objUpEDOC IsNot Nothing Then
                m_objUpEDOC.Dispose()
                m_objUpEDOC = Nothing
            End If

            If m_objHoverEDOC IsNot Nothing Then
                m_objHoverEDOC.Dispose()
                m_objHoverEDOC = Nothing
            End If

            If m_objDownEDOC IsNot Nothing Then
                m_objDownEDOC.Dispose()
                m_objDownEDOC = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
