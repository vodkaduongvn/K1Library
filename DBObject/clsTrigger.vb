Public Class clsTrigger
    Inherits clsDBObjBase

#Region " Members "

    Private m_strDatabaseName As String
    Private m_intTableID As Integer
    Private m_strAction As String
    Private m_blnOnInsert As Boolean
    Private m_blnOnUpdate As Boolean
    Private m_blnOnDelete As Boolean
    Private m_strSQL As String
#End Region

#Region " Enumerations "

    Public Enum enumTriggerAction
        [FOR] = 0
        INSTEAD_OF = 1
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_strDatabaseName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cDATABASENAME, clsDBConstants.cstrNULL), String)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cTABLEID, clsDBConstants.cintNULL), Integer)
        m_strAction = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cTRIGGERACTION, clsDBConstants.cstrNULL), String)
        m_blnOnInsert = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cONINSERT, False), Boolean)
        m_blnOnUpdate = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cONUPDATE, False), Boolean)
        m_blnOnDelete = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cONDELETE, False), Boolean)
        m_strSQL = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Trigger.cSQL, clsDBConstants.cstrNULL), String)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property DatabaseName() As String
        Get
            Return m_strDatabaseName
        End Get
    End Property

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objDB.SysInfo.Tables(m_intTableID)
        End Get
    End Property

    Public ReadOnly Property Action() As enumTriggerAction
        Get
            If m_strAction IsNot Nothing AndAlso m_strAction = "I" Then
                Return enumTriggerAction.INSTEAD_OF
            Else
                Return enumTriggerAction.FOR
            End If
        End Get
    End Property

    Public ReadOnly Property OnInsert() As Boolean
        Get
            Return m_blnOnInsert
        End Get
    End Property

    Public ReadOnly Property OnUpdate() As Boolean
        Get
            Return m_blnOnUpdate
        End Get
    End Property

    Public ReadOnly Property OnDelete() As Boolean
        Get
            Return m_blnOnDelete
        End Get
    End Property

    Public ReadOnly Property SQL() As String
        Get
            Return m_strSQL
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsTrigger
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cTRIGGER, intID)

            Return New clsTrigger(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTrigger)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTrigger)

        Try
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cTRIGGER)
            Dim objDT As DataTable = objDB.GetDataTable(objTable)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTrigger)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsTrigger(objDR, objDB)
                If Not colObjects.ContainsKey(CStr(objItem.DatabaseName)) Then
                    colObjects.Add(CStr(objItem.DatabaseName), objItem)
                End If
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
