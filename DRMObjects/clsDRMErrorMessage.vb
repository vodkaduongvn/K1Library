Public Class clsDRMErrorMessage
    Inherits clsDRMBase

#Region " Members "

    Private m_intStringID As Integer
    Private m_intUIID As Integer
    Private m_strText As String
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer, _
    ByVal intUIID As Integer, _
    ByVal strText As String)
        MyBase.New(objDB, strExternalID, intSecurityID, intTypeID)
        m_intUIID = intUIID
        m_strText = strText
    End Sub
#End Region

#Region " From Existing "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal objMsg As clsErrorMessage, _
    ByVal strText As String)
        MyBase.New(objDB, objMsg)
        If Not objMsg.StringObj Is Nothing Then
            m_intStringID = objMsg.StringObj.ID
        End If
        m_intUIID = objMsg.UIID
        m_strText = strText
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property ErrorMessage() As clsErrorMessage
        Get
            Return CType(m_objDBObj, clsErrorMessage)
        End Get
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

            Dim objStringObj As clsString = Nothing
            Dim intLanguageStringID As Integer = clsDBConstants.cintNULL
            Dim objLanguageString As clsLanguageString

            If Not ErrorMessage Is Nothing Then
                objStringObj = ErrorMessage.StringObj

                If Not objStringObj Is Nothing Then
                    objLanguageString = ErrorMessage.StringObj.LinkLanguageStrings( _
                        CType(m_objDB.Profile.LanguageID, String))
                    intLanguageStringID = objLanguageString.ID
                End If
            End If

            If objStringObj Is Nothing Then
                objStringObj = New clsString(m_objDB, clsDBConstants.cintNULL, m_strExternalID, m_intSecurityID)
                objStringObj.InsertUpdate()
            End If

            Dim intID As Integer = clsDBConstants.cintNULL
            If ErrorMessage IsNot Nothing Then
                intID = ErrorMessage.ID
            End If
            m_objDBObj = New clsErrorMessage(m_objDB, intID, m_strExternalID, m_intSecurityID, _
                objStringObj.ID, m_intUIID)
            ErrorMessage.InsertUpdate()

            objLanguageString = New clsLanguageString(m_objDB, _
                intLanguageStringID, m_strExternalID & " (" & m_objDB.Profile.Language.ExternalID & ")", _
                m_intSecurityID, m_objDB.Profile.LanguageID, objStringObj.ID, m_strText)

            objLanguageString.InsertUpdate()
            objStringObj.LinkLanguageStrings(CStr(objLanguageString.LanguageID)) = objLanguageString

            m_intID = ErrorMessage.ID

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

            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cERRORMESSAGE, ErrorMessage.ID)

            If ErrorMessage.StringObj IsNot Nothing Then
                SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLANGUAGESTRING & "] " & _
                    "WHERE [" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = " & ErrorMessage.StringObj.ID)
                SystemDB.DeleteRecord(clsDBConstants.Tables.cSTRING, ErrorMessage.StringObj.ID)
            End If

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
