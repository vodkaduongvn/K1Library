Public Class clsDRMWarningMessage
    Inherits clsDRMBase

#Region " Members "

    Private m_strStoredProcedure As String
    Private m_strSQL As String
    Private m_intTableID As Integer
    Private m_eWarningType As clsWarningMessage.enumWarningType
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer, _
    ByVal strStoredProcedure As String, _
    ByVal strSQL As String, _
    ByVal intTableID As Integer, _
    ByVal eWarningType As clsWarningMessage.enumWarningType)
        MyBase.New(objDB, strExternalID, intSecurityID, intTypeID)
        m_strStoredProcedure = strStoredProcedure
        m_strSQL = strSQL
        m_intTableID = intTableID
        m_eWarningType = eWarningType
    End Sub
#End Region

#Region " From Existing "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal objMsg As clsWarningMessage, _
    ByVal strSQL As String)
        MyBase.New(objDB, objMsg)
        m_strStoredProcedure = objMsg.StoredProcedure
        m_strSQL = strSQL
        m_intTableID = objMsg.TableID
        m_eWarningType = objMsg.WarningType
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property WarningMessage() As clsWarningMessage
        Get
            Return CType(m_objDBObj, clsWarningMessage)
        End Get
    End Property

    Public Property WarningType() As clsWarningMessage.enumWarningType
        Get
            Return m_eWarningType
        End Get
        Set(ByVal value As clsWarningMessage.enumWarningType)
            m_eWarningType = value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            m_objDB.DeleteStoredProcedure(m_strStoredProcedure)
            m_objDB.CreateStoredProcedure(m_strStoredProcedure, "@Record XML", m_strSQL)

            Dim intID As Integer = clsDBConstants.cintNULL
            If WarningMessage IsNot Nothing Then
                intID = WarningMessage.ID
            End If

            m_objDBObj = New clsWarningMessage(m_objDB, intID, m_strExternalID, _
                m_intSecurityID, m_strStoredProcedure, m_strSQL, m_eWarningType, m_intTableID)
            WarningMessage.InsertUpdate()

            m_intID = WarningMessage.ID

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            m_objDB.DeleteStoredProcedure(m_strStoredProcedure)
            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cWARNING, WarningMessage.ID)

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
    End Sub
#End Region

End Class
