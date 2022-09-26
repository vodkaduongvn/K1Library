
Class clsAuditTrailRepository

    Private ReadOnly m_objDb As clsDB

    Sub New(ByVal objDb As clsDB)
        m_objDb = objDb
    End Sub

    ''' <summary>
    ''' Creates a new Audit Record
    ''' </summary>    
    ''' <param name="strAuditExternalId"></param>
    ''' <param name="strCreatedByApplication"></param>
    ''' <param name="intPersonId"></param>
    ''' <param name="objMethod"></param>
    ''' <param name="objTable"></param>
    ''' <param name="intId"></param>
    ''' <param name="strExternalId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert(ByVal strAuditExternalId As String,
                           ByVal strCreatedByApplication As String,
                           ByVal intPersonId As Integer,
                           ByVal objMethod As clsMethod,
                           ByVal objTable As clsTable, ByVal intId As Integer,
                           ByVal strExternalId As String,
                           Optional ByVal fileName As String = "") As Integer

        Dim objAuditTable As clsTable = m_objDb.SysInfo.Tables(clsDBConstants.Tables.cAUDITTRAIL)

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(objAuditTable)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, strAuditExternalId)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cDATE, Now)
        If strCreatedByApplication Is Nothing Then
            If m_objDb.DataAccessType = clsDB.enumDataAccessType.WEB_SESSION Then
                Dim objDbWs As clsDB_Web = CType(m_objDb, clsDB_Web)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cCREATEDBYAPPLICATION, objDbWs.Product)
            Else
                colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cCREATEDBYAPPLICATION, My.Application.Info.ProductName)
            End If
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cCREATEDBYAPPLICATION, strCreatedByApplication)
        End If
        If m_objDb.Profile IsNot Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cPERSONID, m_objDb.Profile.PersonID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_objDb.Profile.NominatedSecurityID)
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_objDb.SysInfo.K1Configuration.DRMDefaultSecurityID)
        End If
        If Not intPersonId = clsDBConstants.cintNULL Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cPERSONID, intPersonId)
        End If
        If Not objMethod Is Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cMETHODID, objMethod.ID)
        End If
        If Not objTable Is Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cTABLEID, objTable.ID)
        End If
        If Not intId = clsDBConstants.cintNULL Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cRECORDID, intId)
        End If
        If Not strExternalId = clsDBConstants.cstrNULL Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cRECORDEXTERNALID, strExternalId)
        End If
        If (Not String.IsNullOrEmpty(fileName)) Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.AuditTrail.cFileName, Path.GetFileName(fileName))
        End If

        Dim intNewId As Integer = colMasks.Insert(m_objDb, True, False)

        Return intNewId

    End Function

    ''' <summary>
    ''' Inserts the Serialized Audit Data to Database.
    ''' </summary>
    ''' <param name="objRecordData"></param>
    ''' <param name="auditRecordId"></param>
    ''' <remarks></remarks>
    Public Sub InsertAuditXmlData(ByVal objRecordData As clsAuditTrailRecordData,
                                  ByVal auditRecordId As Integer,
                                  Optional ByVal includeBinaryData As Boolean = True)

        '[Naing] We serialize the data to xml and save it in a temp file
        Dim auditDataXmlFileName As String = objRecordData.SerializeToXml(includeBinaryData)

        If Not (String.IsNullOrEmpty(auditDataXmlFileName)) Then

            Using objFileStream As New FileStream(auditDataXmlFileName, FileMode.Open, FileAccess.Read)
                Dim intLength As Integer = CInt(objFileStream.Length)
                Dim intOffsetFs As Integer = 0
                Dim intOffsetDb As Integer = 0
                Dim intChunkSize As Integer = 524288 '512 kb like in clsDB and Web Services 'old value 8000 bytes

                m_objDb.RaiseFileTransferInit(clsDB.enumDataAccessType.AUDIT_TRAIL,
                                              CInt(Math.Ceiling(intLength / intChunkSize)),
                                              clsDB.enumTransferType.UPLOAD)

                m_objDb.ExecuteSQL("UPDATE [" & clsDBConstants.Tables.cAUDITTRAIL & "]" & vbCrLf &
                                   "SET [" & clsDBConstants.Fields.AuditTrail.cRECORDDATA & "] = ''" & vbCrLf &
                                   "WHERE [" & clsDBConstants.Fields.cID & "] = " & auditRecordId)

                While intOffsetFs < intLength
                    If intOffsetFs + intChunkSize >= intLength Then
                        intChunkSize = intLength - intOffsetFs
                    End If

                    If m_objDb.ThreadedOperationCancelled Then
                        ' Bug Fix for #1300002434 and 1300002436 change exception message to show what is actually happening -- 2013-06-05 -- Peter Melisi
                        Throw New clsK1Exception("The user has cancelled the operation.")
                    End If

                    Dim arrBytes(intChunkSize) As Byte
                    objFileStream.Read(arrBytes, 0, intChunkSize)

                    '[Naing] 04/10/2012 For some reason this step here introduces an invalid xml character at the end of the xml file.
                    'This invalid character is only detected in SQL Server xml data type which causes a problem when converting the audit 
                    'trail record data into xml type from nvarchar type. To rectify a method has been added to detect unicode characters 
                    'which are invalid in xml spec version 1
                    Dim strData As String
                    'detect the last iteration
                    If (intOffsetFs + intChunkSize) >= intLength Then
                        strData = RemoveUnsupportedCharacters(Text.Encoding.UTF8.GetString(arrBytes))
                    Else
                        strData = Text.Encoding.UTF8.GetString(arrBytes)
                        '[Naing] 07/02/2013 Bug Fix#1300002399 
                        If (strData.Length <= intChunkSize) Then
                            strData = RemoveUnsupportedCharacters(strData)
                        End If
                    End If

                    m_objDb.ExecuteSQL("UPDATE [" & clsDBConstants.Tables.cAUDITTRAIL & "]" & vbCrLf & _
                                       "SET [" & clsDBConstants.Fields.AuditTrail.cRECORDDATA & "]" & _
                                       ".WRITE('" & clsDB.SQLString(strData) & "'" & _
                                       "," & intOffsetDb & "," & intChunkSize & ")" & vbCrLf & _
                                       "WHERE [" & clsDBConstants.Fields.cID & "] = " & auditRecordId)

                    m_objDb.RaiseFileTransferStep()

                    '[Naing] 07/02/2013 Bug Fix#1300002399 
                    'For some reason this line of code, Text.Encoding.UTF8.GetString(arrBytes),
                    'ignores some bytes from the file-stream so the offset size is mismatched 
                    'between the database and the file stream. This was added to fix the 
                    'mismatch if it occurs because this mismatch doesn't always occur.
                    intOffsetDb += Math.Min(intChunkSize, strData.Length)

                    intOffsetFs += intChunkSize
                End While
            End Using

            m_objDb.RaiseFileTransferEnd()
            File.Delete(auditDataXmlFileName)

        End If

    End Sub

    ''' <summary>
    ''' Insert the image file as a blob into the database.
    ''' </summary>
    ''' <param name="auditRecordId"></param>
    ''' <param name="fullFileName"></param>
    ''' <remarks></remarks>
    Public Sub InsertAuditBlob(ByVal auditRecordId As Integer, fullFileName As String)

        If (String.IsNullOrEmpty(fullFileName) AndAlso Not File.Exists(fullFileName)) Then
            Throw New ArgumentException("File does not exist.")
        End If
        
        Dim objFileInfo = New FileInfo(fullFileName)

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@BinaryData", New Byte() {0}))
        colParams.Add(New clsDBParameter("@Id", auditRecordId))
        m_objDb.ExecuteSQL(String.Format("UPDATE [{0}] SET [{1}] = @BinaryData WHERE ID = @Id",
                                         clsDBConstants.Tables.cAUDITTRAIL,
                                         clsDBConstants.Fields.AuditTrail.cImage),
                           colParams)

        m_objDb.WriteBLOB(clsDBConstants.Tables.cAUDITTRAIL,
                          clsDBConstants.Fields.AuditTrail.cImage,
                          SqlDbType.Image,
                          CInt(objFileInfo.Length),
                          auditRecordId,
                          objFileInfo.FullName,
                          True)

    End Sub

    ''' <summary>
    ''' Cleans out unsupported Unicode char from string
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns>Cleaned out string valid for building xml document</returns>
    ''' <remarks></remarks>
    Private Shared Function RemoveUnsupportedCharacters(data As String) As String
        Dim badChars = data.ToCharArray().Where(Function(c) IsBadXmlChar(c)).ToList()
        badChars.ForEach(
            Sub(c)
                Dim index = data.IndexOf(c)
                data = data.Remove(index, 1)
            End Sub)
        Return data
    End Function

    ''' <summary>
    ''' Detects chars not supported in xml char set
    ''' </summary>
    ''' <param name="cc"></param>
    ''' <returns>True if char is NOT supported by xml character set</returns>
    ''' <remarks></remarks>
    Private Shared Function IsBadXmlChar(cc As Char) As Boolean
        Dim c As Integer = AscW(cc)
        If ((c = &H9) OrElse (c = &HA) OrElse (c = &HD) OrElse _
            ((c >= &H20) AndAlso (c <= &HD7FF)) OrElse _
            ((c >= &HE000) AndAlso (c <= &HFFFD)) OrElse _
            ((c >= &H10000) AndAlso (c <= &H10FFFF))) Then
            Return False
        End If
        Return True
    End Function

End Class