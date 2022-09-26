Friend Class clsDRMCaption
    Inherits clsDRMBase

#Region " Members "

    Private m_intStringID As Integer
    Private m_intUIID As Integer
    Private m_strText As String
#End Region

#Region " Constructors "

#Region " New Field "

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
    ByVal objCaption As clsCaption, _
    ByVal strText As String)
        MyBase.New(objDB, objCaption)
        If Not objCaption.StringObj Is Nothing Then
            m_intStringID = objCaption.StringObj.ID
        End If
        m_intUIID = objCaption.UIID
        m_strText = strText
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property Caption() As clsCaption
        Get
            Return CType(m_objDBObj, clsCaption)
        End Get
    End Property
#End Region

#Region " Methods "

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrailRecord As Boolean = True)
        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            Dim objStringObj As clsString = Nothing
            Dim intLanguageStringID As Integer = clsDBConstants.cintNULL
            Dim objLanguageString As clsLanguageString

            If Not m_objDBObj Is Nothing Then
                objStringObj = Caption.StringObj

                If Not objStringObj Is Nothing Then
                    objLanguageString = Caption.StringObj.LinkLanguageStrings(CType(m_objDB.Profile.LanguageID, String))
                    If Not objLanguageString Is Nothing Then
                        intLanguageStringID = objLanguageString.ID
                    End If
                End If
            End If

            If objStringObj Is Nothing Then
                objStringObj = New clsString(m_objDB, clsDBConstants.cintNULL, m_strExternalID, m_intSecurityID)
                objStringObj.InsertUpdate(blnCreateAuditTrailRecord)
            End If

            If Caption Is Nothing Then
                m_objDBObj = New clsCaption(m_objDB, clsDBConstants.cintNULL, m_strExternalID, m_intSecurityID, _
                    objStringObj.ID, m_intUIID)
                Caption.InsertUpdate(blnCreateAuditTrailRecord)
            End If

            objLanguageString = New clsLanguageString(m_objDB, _
                intLanguageStringID, m_strExternalID & " (" & m_objDB.Profile.Language.ExternalID & ")", _
                m_intSecurityID, m_objDB.Profile.LanguageID, objStringObj.ID, m_strText)

            objLanguageString.InsertUpdate(blnCreateAuditTrailRecord)
            objStringObj.LinkLanguageStrings(CStr(objLanguageString.LanguageID)) = objLanguageString
            m_objDB.SysInfo.DRMInsertUpdateCaption(Caption)

            m_intID = Caption.ID

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)

        Catch ex As Exception

            If blnCreatedTransaction Then m_objDB.EndTransaction(False)
            Throw

        End Try

    End Sub

    Public Sub Delete(Optional ByVal blnCreateAuditTrailRecord As Boolean = True)
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cCAPTION, Caption.ID, blnCreateAuditTrailRecord)

            If Caption.StringObj IsNot Nothing Then
                SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLANGUAGESTRING & "] " & _
                    "WHERE [" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = " & Caption.StringObj.ID)
                SystemDB.ExecuteSQL("UPDATE  [" & clsDBConstants.Tables.cCAPTION & "] " & _
                                    "SET [" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = NULL " & _
                                    "WHERE [" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = " & Caption.StringObj.ID)
                SystemDB.DeleteRecord(clsDBConstants.Tables.cSTRING, Caption.StringObj.ID)

                m_objDB.SysInfo.DRMDeleteCaption(Caption)
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
