' Created: Unknown
' Created By: Unknown

'Updated: 23 August 2013
'Updated By: Naing
'Comment: As part of Recfind 2.6 release

''' <summary>
''' This Utility class is responsible for logging AuidtTrail data. It is somewhat linked to the clsMaskFieldDictionary class.
''' </summary>
''' <remarks></remarks>
Public Class clsAuditTrail

#Region " Methods "

    ''' <summary>
    ''' Creates audit trail records for each ID in the data table if we are auditing the specified table and method
    ''' </summary>
    ''' <param name="objDb">The database object</param>
    ''' <param name="eMethod">The method to audit</param>
    ''' <param name="objTable">The table to audit</param>
    ''' <param name="objDt">This should be a data table containing a single column of the record IDs we are auditing</param>
    Public Shared Sub CreateAuditTrailRecords(ByVal objDb As clsDB,
                                              ByVal eMethod As clsMethod.enumMethods,
                                              ByVal objTable As clsTable,
                                              ByVal objDt As DataTable)

        If objDt Is Nothing OrElse objDt.Rows.Count = 0 Then
            Return
        End If

        'first just double check that we are auditing this table and method
        Dim objTableMethod As clsTableMethod
        Dim objMethod As clsMethod = objDb.SysInfo.Methods(CStr(eMethod))

        If Not objMethod Is Nothing Then
            objTableMethod = objTable.TableMethods(CStr(objMethod.ID))

            If Not objTableMethod Is Nothing AndAlso objTableMethod.Audit Then
                For Each objRow As DataRow In objDt.Rows

                    Dim intId As Integer = CInt(objRow(0))

                    CreateTableMethodRecord(objDb, eMethod, objTable, intId, objDb.GetRecordExternalID(objTable, intId))
                Next
            End If
        End If
    End Sub

    Public Shared Function AuditTableMethodData(ByVal objDb As clsDB, ByVal eMethod As clsMethod.enumMethods, ByVal objTable As clsTable) As Boolean
        Dim objTableMethod As clsTableMethod
        Dim objMethod As clsMethod = objDb.SysInfo.Methods(CStr(eMethod))

        If Not objMethod Is Nothing Then
            objTableMethod = objTable.TableMethods(CStr(objMethod.ID))

            If Not objTableMethod Is Nothing AndAlso objTableMethod.Audit AndAlso objTableMethod.AuditData Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Shared Sub CreateUnsuccessfulLoginRecord(ByVal objDb As clsDB, ByVal strUserName As String)

        If objDb.SysInfo.K1Configuration.AuditUnsuccessfulLogins Then
            Dim strNewExternalId As String = String.Format("Unsuccessful login attempt for UserID '{0}' from '{1}'", strUserName, Environment.MachineName)
            LogAuditTrail(objDb, strNewExternalId)
        End If

    End Sub

    Public Shared Sub CreateLoginRecord(ByVal objDb As clsDB)

        If objDb.SysInfo.K1Configuration.AuditLogins Then
            Dim strExternalId = String.Format("{0} logged in", objDb.Profile.Person.ExternalID)
            LogAuditTrail(objDb, strExternalId)
        End If

    End Sub

    Public Shared Sub CreateLogoffRecord(ByVal objDb As clsDB)

        If objDb.SysInfo.K1Configuration.AuditLogoffs Then
            Dim strExternalId = String.Format("{0} logged off", objDb.Profile.Person.ExternalID)
            LogAuditTrail(objDb, strExternalId)
        End If

    End Sub

    Public Shared Sub CreateLogoffRecord(ByVal objDb As clsDB,
                                         ByVal intUserProfileId As Integer,
                                         Optional ByVal strCreatedByApplication As String = Nothing)

        If objDb.SysInfo.K1Configuration.AuditLogoffs Then

            Dim objUp As clsUserProfileBase = clsUserProfileBase.GetItem(intUserProfileId, objDb)
            Dim strExternalId As String = String.Format("{0} logged off", objUp.Person.ExternalID)
            LogAuditTrail(objDb, strExternalId, objUp.PersonID, strCreatedByApplication)

        End If

    End Sub

    Public Shared Sub CreateSessionTimeoutRecord(ByVal objDb As clsDB,
                                                 ByVal intUserProfileId As Integer,
                                                 Optional ByVal strCreatedByApplication As String = Nothing)

        If objDb.SysInfo.K1Configuration.AuditLogoffs Then
            Dim objUp As clsUserProfileBase = clsUserProfileBase.GetItem(intUserProfileId, objDb)
            Dim strExternalId = String.Format("The session has timed out for user {0}", objUp.Person.ExternalID)
            LogAuditTrail(objDb, strExternalId, objUp.PersonID, strCreatedByApplication)
        End If

    End Sub

    Public Shared Sub CreateTableMethodRecord(ByVal objDb As clsDB,
                                              ByVal eMethod As clsMethod.enumMethods,
                                              ByVal objTable As clsTable,
                                              Optional ByVal intId As Integer = clsDBConstants.cintNULL,
                                              Optional ByVal strExternalId As String = clsDBConstants.cstrNULL,
                                              Optional ByVal objRecordData As clsAuditTrailRecordData = Nothing)
        Dim objTableMethod As clsTableMethod
        Dim objMethod As clsMethod = objDb.SysInfo.Methods(CStr(eMethod))

        'Update the session time when ever we do any activity which checks auditing
        objDb.SessionUpdate()

        If Not objMethod Is Nothing Then
            objTableMethod = objTable.TableMethods(CStr(objMethod.ID))

            If Not objTableMethod Is Nothing AndAlso objTableMethod.Audit Then
                Dim strNewExternalId As String

                If objDb.Profile Is Nothing Then
                    strNewExternalId = My.Application.Info.ProductName &
                        " [" & objTable.DatabaseName & "." & objMethod.ExternalID & "]"
                Else
                    strNewExternalId = objDb.Profile.Person.ExternalID &
                        " [" & objTable.DatabaseName & "." & objMethod.ExternalID & "]"
                End If

                If Not intId = clsDBConstants.cintNULL AndAlso Not strExternalId = clsDBConstants.cstrNULL Then
                    strNewExternalId &= ": Record #" & intId & " ('" & strExternalId & "')"
                End If

                LogAuditTrail(objDb, strNewExternalId, objMethod, objTable, intId, strExternalId, objRecordData)
            End If
        End If
    End Sub

    Private Shared Sub LogAuditTrail(ByVal objDb As clsDB, ByVal strAuditExternalId As String)

        LogAuditTrail(objDb, strAuditExternalId, Nothing, Nothing, clsDBConstants.cintNULL, Nothing, Nothing)

    End Sub

    Private Shared Sub LogAuditTrail(ByVal objDb As clsDB,
                                     ByVal strExternalId As String,
                                     ByVal intPersonId As Integer,
                                     ByVal strCreatedByApplication As String)

        LogAuditTrail(objDb, strExternalId, Nothing, Nothing, clsDBConstants.cintNULL, Nothing, Nothing, intPersonId, strCreatedByApplication)

    End Sub

    'ToDo [Naing] Large methods like this are really hard to maintain, must re-factor into smaller pieces.
    Private Shared Sub LogAuditTrail(ByVal objDb As clsDB,
                                              ByVal strAuditExternalId As String,
                                              ByVal objMethod As clsMethod,
                                              ByVal objTable As clsTable,
                                              ByVal intId As Integer,
                                              ByVal strExternalId As String,
                                              ByVal objRecordData As clsAuditTrailRecordData,
                                              Optional ByVal intPersonId As Integer = clsDBConstants.cintNULL,
                                              Optional ByVal strCreatedByApplication As String = Nothing)
        Dim blnBeginTran As Boolean = False

        Try

            blnBeginTran = Not objDb.HasTransaction

            If blnBeginTran Then
                objDb.BeginTransaction()
            End If

            '[Naing] Use the repository pattern to separate concerns.
            Dim repo = New clsAuditTrailRepository(objDb)

            '[Naing] Extra logic was added to ensure that only Image enabled tables like EDOC will use this optimized Auditing logic.
            If (objRecordData IsNot Nothing AndAlso
                objTable IsNot Nothing AndAlso
                objTable.IsImageEnabledTable()) Then

                '[Naing] Try get temporary path of the file that is attached to the original record.
                Dim strTempFileFullName = String.Empty
                '[Naing] Try get the original name of the file that is attached to the original record.
                Dim strFileName = String.Empty

                If (objRecordData.OriginalRecord IsNot Nothing) Then

                    Dim objOriginalMaskField As clsMaskField = objRecordData.OriginalRecord.MaskFieldCollection(clsDBConstants.Fields.EDOC.cIMAGE)

                    If (objOriginalMaskField IsNot Nothing) Then

                        '[Naing] This will be a value of "Path.GetTempPath()" \ "Path.GetRandomFileName()", except for Inserts. 
                        'See clsMaskFieldDictionary [Insert]/[Update]/[Delete] methods for details of the implementation.
                        strTempFileFullName = objOriginalMaskField.Value1.FileName

                        '[Naing] Only save the file name if it was necessary to save a copy in the audit trail. If not ignore it. See clsAuditTrailRepository --> Insert method.
                        If (Not String.IsNullOrEmpty(strTempFileFullName)) Then
                            '[Naing] Try get the original file name
                            strFileName = objRecordData.OriginalRecord.MaskFieldCollection.
                                GetMaskValue(clsDBConstants.Fields.EDOC.cFILENAME, String.Empty).
                                ToString()
                        End If

                    End If

                End If

                '[Naing] First we create Audit Record.
                Dim auditRecordId = repo.Insert(strAuditExternalId, strCreatedByApplication,
                                                intPersonId, objMethod,
                                                objTable, intId,
                                                strExternalId, strFileName)

                '[Naing] Here for readability's sake.
                ' ReSharper disable ConvertToConstant.Local
                Dim blnIncludeBinaryData = False
                ' ReSharper restore ConvertToConstant.Local
                '[Naing] We stream the XML file to the database. Unfortunately this will exclude the Thumbnail data for the EDOC table.
                repo.InsertAuditXmlData(objRecordData, auditRecordId, blnIncludeBinaryData)

                '[Naing] We stream the Blob file to the database.
                If (Not String.IsNullOrEmpty(strTempFileFullName) AndAlso File.Exists(strTempFileFullName)) Then
                    repo.InsertAuditBlob(auditRecordId, strTempFileFullName)
                End If

            Else '[Naing] This will execute the non optimized standard Audit logic. To ensure that I didn't break anything.

                '[Naing] First we create Audit Record.
                Dim auditRecordId = repo.Insert(strAuditExternalId, strCreatedByApplication,
                                                intPersonId, objMethod,
                                                objTable, intId,
                                                strExternalId)

                If (objRecordData IsNot Nothing) Then
                    '[Naing] We stream the XML file to the database.
                    repo.InsertAuditXmlData(objRecordData, auditRecordId)
                End If

            End If

            If blnBeginTran Then
                objDb.EndTransaction(True)
            End If

        Catch ex As Exception

            If blnBeginTran Then
                objDb.EndTransaction(False)
            End If

            Throw New Exception("Failed to complete audit trail log.", ex)

        End Try

    End Sub

#End Region

End Class