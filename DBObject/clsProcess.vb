Public Class clsProcess
    Inherits clsDBObjBase

#Region " Members "

    Private m_intTableID As Integer
    Private m_eProcessType As enumProcessType
    Private m_strProcess As String
    Private m_blnRequiresInput As Boolean
    Private m_eInputFormat As enumInputFormat
    Private m_intProcessFileID As Integer
    Private m_intPersonID As Integer
#End Region

#Region " Constructors "

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cTABLEID, clsDBConstants.cintNULL), Integer)
        m_eProcessType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cPROCESSTYPE, enumProcessType.EXECUTABLE), enumProcessType)
        m_strProcess = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cPROCESS, ""), String)
        m_blnRequiresInput = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cREQUIRESINPUT, False), Boolean)
        m_eInputFormat = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cINPUTTYPE, enumInputFormat.XML), enumInputFormat)
        m_intProcessFileID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cPROCESSFILEID, clsDBConstants.cintNULL), Integer)
        m_intPersonID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Process.cPERSONID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Enumerations "

    Public Enum enumProcessType
        EXECUTABLE = 0
        STORED_PROCEDURE = 1
    End Enum

    Public Enum enumInputFormat
        XML = 0
        CSV = 1
    End Enum
#End Region

#Region " Properties "

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property ProcessType() As enumProcessType
        Get
            Return m_eProcessType
        End Get
    End Property

    Public ReadOnly Property Process() As String
        Get
            Return m_strProcess
        End Get
    End Property

    Public ReadOnly Property RequiresInput() As Boolean
        Get
            Return m_blnRequiresInput
        End Get
    End Property

    Public ReadOnly Property InputFormat() As enumInputFormat
        Get
            Return m_eInputFormat
        End Get
    End Property

    Public ReadOnly Property ProcessFileID() As Integer
        Get
            Return m_intProcessFileID
        End Get
    End Property

    Public ReadOnly Property PersonID() As Integer
        Get
            Return m_intPersonID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsProcess
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cPROCESS, intID)

            Return New clsProcess(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objTable As clsTable, ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsProcess)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsProcess)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cPROCESS, _
                clsDBConstants.Fields.TableMethod.cTABLEID, objTable.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsProcess)
            For Each objDR As DataRow In objDT.Rows
                Dim objRec As New clsProcess(objDR, objDB)
                colObjects.Add(CStr(objRec.ID), objRec)
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
