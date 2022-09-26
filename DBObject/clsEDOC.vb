'=====================================================================
' This class represents the table EDOC in the Database.
'=====================================================================

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      6/07/2004    Implemented.
' Naing     10/06/2013   Extended with Repository Pattern.
'=====================================================================

Imports Aspose.Email
Imports K1Library.FrameworkCollections

Public Class clsEdocStatusCodesRepository

    Private m_objDbContext As clsDB
    Private m_objTypeTable As clsTable
    Private m_objCodeTable As clsTable
    Private m_intTypeId As Integer

    Sub New(objDb As clsDB)
        m_objDbContext = objDb
        m_objTypeTable = m_objDbContext.SysInfo.Tables(clsDBConstants.Tables.cTYPE)
        m_objCodeTable = m_objDbContext.SysInfo.Tables(clsDBConstants.Tables.cCODES)
        Dim m_objTypeDt = m_objDbContext.GetDataTableByField(m_objTypeTable.DatabaseName, clsDBConstants.Fields.cEXTERNALID, "EDOC Status")
        m_intTypeId = CInt(m_objTypeDt.Rows.Item(0)(clsDBConstants.Fields.cID))
    End Sub

    Public Function GetStatusCode(ByVal intCodeId As Integer) As String
        Dim objCodesDt = m_objDbContext.GetDataTableByField(m_objCodeTable.DatabaseName, clsDBConstants.Fields.cTYPEID, m_intTypeId)
        Return CType((From dr As DataRow In objCodesDt.Rows.OfType(Of DataRow)() Where (CInt(dr(clsDBConstants.Fields.cID)) = intCodeId)
                      Select dr(clsDBConstants.Fields.Codes.cCODE)).FirstOrDefault(), String)
    End Function

    Public Function GetStatusExternalId(ByVal intCodeId As Integer) As String
        Dim objCodesDt = m_objDbContext.GetDataTableByField(m_objCodeTable.DatabaseName, clsDBConstants.Fields.cTYPEID, m_intTypeId)
        Return CType((From dr As DataRow In objCodesDt.Rows.OfType(Of DataRow)() Where (CInt(dr(clsDBConstants.Fields.cID)) = intCodeId)
                      Select dr(clsDBConstants.Fields.cEXTERNALID)).FirstOrDefault(), String)
    End Function

    Public Function GetStatusId(ByVal strCode As String) As Integer
        Dim objCodesDt = m_objDbContext.GetDataTableByField(m_objCodeTable.DatabaseName, clsDBConstants.Fields.cTYPEID, m_intTypeId)
        Return (From dr As DataRow In objCodesDt.Rows.OfType(Of DataRow)() Where (dr(clsDBConstants.Fields.Codes.cCODE).ToString() = strCode)
                Select CInt(dr(clsDBConstants.Fields.cID))).FirstOrDefault()
    End Function

End Class

Public Class clsEDOC
    Inherits clsDBObjBase

#Region " Members "

    Private m_strContentType As String
    Private m_intSize As Integer
    Private m_strFileName As String
    Private m_strFileExtension As String
    Private m_strStatusId As Integer
    'Private m_intWidth As Integer = clsDBConstants.cintNULL
    'Private m_intHeight As Integer = clsDBConstants.cintNULL

#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB,
    ByVal intID As Integer,
    ByVal strExternalID As String,
    ByVal intSecurityID As Integer,
    ByVal strFileName As String)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_strFileName = strFileName
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_strContentType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.EDOC.cCONTENTTYPE, clsDBConstants.cstrNULL), String)
        m_intSize = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.EDOC.cSIZE, clsDBConstants.cintNULL), Integer)
        m_strFileName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.EDOC.cFILENAME, clsDBConstants.cstrNULL), String)
        m_strFileExtension = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.EDOC.cSUFFIX, clsDBConstants.cstrNULL), String)
        m_strStatusId = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.EDOC.cEDOCSTATUSID, clsDBConstants.cintNULL), Integer)

    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property ContentType() As String
        Get
            Return m_strContentType
        End Get
    End Property

    Public ReadOnly Property Size() As Integer
        Get
            Return m_intSize
        End Get
    End Property

    Public ReadOnly Property FileName() As String
        Get
            Return m_strFileName
        End Get
    End Property

    Public ReadOnly Property FileExtension() As String
        Get
            Return m_strFileExtension
        End Get
    End Property

    Public ReadOnly Property FullFileExtension() As String
        Get
            If m_strFileExtension.Length > 0 Then
                Return "." & m_strFileExtension
            End If
            Return m_strFileExtension
        End Get
    End Property

    Public ReadOnly Property StatusId As Integer
        Get
            Return m_strStatusId
        End Get
    End Property

    Private strStatusCode As String = String.Empty
    Public ReadOnly Property StatusCode() As String
        Get
            If (String.IsNullOrEmpty(strStatusCode)) Then
                strStatusCode = GetEdocStatusCode(Me.StatusId, m_objDB)
            End If
            Return strStatusCode
        End Get
    End Property

    Private strStatusExternalId As String = String.Empty
    Public ReadOnly Property StatusExternalId() As String
        Get
            If String.IsNullOrEmpty(strStatusExternalId) Then
                strStatusExternalId = GetEdocStatusCodeExternalId(Me.StatusId, m_objDB)
            End If
            Return strStatusExternalId
        End Get
    End Property

    Public Shared Function GetEdocStatusCode(codeId As Integer, objDb As clsDB) As String
        'Get the database value
        Dim repo = New clsEdocStatusCodesRepository(objDb)
        Dim strCode = repo.GetStatusCode(codeId)
        Return strCode
    End Function

    Public Shared Function GetEdocStatusCodeExternalId(id As Integer, objDb As clsDB) As String
        'Get the database value
        Dim repo = New clsEdocStatusCodesRepository(objDb)
        Dim strCode = repo.GetStatusExternalId(id)
        Return strCode
    End Function

#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsEDOC
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cEDOC, intID)
            Return New clsEDOC(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cEDOC), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)

        If m_intID = clsDBConstants.cintNULL Then
            Dim objMask As clsMaskField = colMasks(clsDBConstants.Fields.EDOC.cIMAGE)
            objMask.Value1.Value = New Byte() {0}
            objMask.Value1.FileName = m_strFileName

            UpdateEDOCFields(m_objDB, colMasks, m_strFileName, False, clsDBConstants.cintNULL)
            m_strFileName = CType(colMasks.GetMaskValue(clsDBConstants.Fields.EDOC.cFILENAME, m_strFileName), String)
            m_intSize = CInt(colMasks.GetMaskValue(clsDBConstants.Fields.EDOC.cSIZE, 0))

            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

#Region " Business Logic "

    Public Sub SaveImageToFile(ByVal strFile As String)
        Database.ReadBLOB(clsDBConstants.Tables.cEDOC, clsDBConstants.Fields.EDOC.cIMAGE, ID, strFile)
    End Sub

    Public Shared Sub UpdateEDOCFields(ByVal objDB As clsDB,
                                       ByVal colMasks As clsMaskFieldDictionary,
                                       ByVal strFileName As String,
                                       ByVal blnFromScanner As Boolean,
                                       ByVal intPreviousEDOCID As Integer)

        If Not File.Exists(strFileName) Then
            Return
        End If

        Dim objFileInfo As New FileInfo(strFileName)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cFILENAME, IO.Path.GetFileName(strFileName))
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCONTENTTYPE, GetMIMEType(strFileName))
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSIZE, objFileInfo.Length)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSUFFIX, Path.GetExtension(strFileName).Replace(".", ""))

        'Changed Published Date to be Modified Date of file instead of Creation Date
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cPUBLISHEDDATE, objFileInfo.LastWriteTime)
        'colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cPUBLISHEDDATE, objFileInfo.CreationTime)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cORIGINALCOPY, "O")
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCHECKEDOUT, False)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cISLATESTVERSION, True)

        Dim intVersion As Integer = 1
        If intPreviousEDOCID <> clsDBConstants.cintNULL Then
            Dim objRecDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cEDOC, intPreviousEDOCID)
            intVersion = CInt(clsDB.NullValue(objRecDT.Rows(0)(clsDBConstants.Fields.EDOC.cVERSIONNUMBER), 1)) + 1

            Dim objMask As clsMaskField = colMasks(clsDBConstants.Fields.EDOC.cPREVIOUSEDOCID)
            objMask.Value1.Value = intPreviousEDOCID
            clsMaskField.LoadLinkedData(objMask, objRecDT)

            Dim intParentMDP As Integer = CInt(clsDB.NullValue(objRecDT.Rows(0)(clsDBConstants.Fields.EDOC.cMETADATAPROFILEID), clsDBConstants.cintNULL))
            If Not intParentMDP = clsDBConstants.cintNULL AndAlso objMask.Field.IsForeignKey Then
                objMask = colMasks(clsDBConstants.Fields.EDOC.cMETADATAPROFILEID)
                objMask.Value1.Value = intParentMDP
                clsMaskField.LoadLinkedData(objMask, objMask.Field.FieldLink.IdentityTable)
            End If
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cPREVIOUSEDOCID, Nothing)
        End If

        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cVERSIONNUMBER, intVersion)

        If blnFromScanner Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cORIGINALPATH,
                "[" & System.Environment.MachineName & "] (Scanned Image)")
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cORIGINALPATH,
                "[" & System.Environment.MachineName & "] " & IO.Path.GetDirectoryName(strFileName))
        End If

        Dim objFileSummary As New clsFileSummary(strFileName)
        If objFileSummary IsNot Nothing Then
            If Not String.IsNullOrEmpty(objFileSummary.Author) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cAUTHOR, objFileSummary.Author)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Category) Then
                colMasks.UpdateMaskObj("Category", objFileSummary.Category)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Comments) Then
                colMasks.UpdateMaskObj("Comments", objFileSummary.Comments)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.KeyWords) Then
                colMasks.UpdateMaskObj("KeyWords", objFileSummary.KeyWords)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Subject) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSUBJECT, objFileSummary.Subject)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Title) Then
                colMasks.UpdateMaskObj("Title", objFileSummary.Title)
            End If
        End If

        'Aspose can handle .msg and .eml files
        If Path.GetExtension(strFileName).IsEmailExtension() Then
            UpdateEmailEDOCFields(colMasks, strFileName)
        End If
    End Sub

    Public Shared Sub UpdateEmailEDOCFields(ByRef colMasks As clsMaskFieldDictionary, ByVal strFileName As String)
        Dim objLicense As Aspose.Email.License = New Aspose.Email.License()
        'Pass only the name of the license file embedded in the assembly
        objLicense.SetLicense("Aspose.Email.lic")

        'Load MSG file into MailMessage object
        Using objMessage = MailMessage.Load(strFileName)
            'set published date to date received
            If IsDate(objMessage.Headers("Date")) AndAlso (CDate(objMessage.Headers("Date")) <> Date.MinValue) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cPUBLISHEDDATE, CDate(objMessage.Headers("Date")).ToString("MMM dd, yyyy HH:mm:ss"))
            End If

            If objMessage.From IsNot Nothing Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSENDER, objMessage.From.DisplayName & " <" & objMessage.From.Address & ">")
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cAUTHOR, objMessage.From.DisplayName)
            End If

            'build string of recipients
            Dim strRecipient As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.To
                strRecipient &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strRecipient.Equals(String.Empty) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cRECIPIENT, strRecipient)
            End If

            If Not objMessage.Subject.Equals(String.Empty) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSUBJECT, objMessage.Subject)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cABSTRACT, objMessage.Subject)
            End If

            'build string of CC
            Dim strCC As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.CC
                strCC &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strCC.Equals(String.Empty) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCCLIST, strCC)
            End If

            'build string of BCC
            Dim strBCC As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.Bcc
                strBCC &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strBCC.Equals(String.Empty) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cBCCLIST, strBCC)
            End If
        End Using
    End Sub

    Public Shared Function GetEmailSentDate(strFileName As String) As Date
        Try
            Dim objLicense As Aspose.Email.License = New Aspose.Email.License()
            'Pass only the name of the license file embedded in the assembly
            objLicense.SetLicense("Aspose.Email.lic")

            'Load MSG file into MailMessage object
            Using objMessage = MailMessage.Load(strFileName)
                'set published date to date received
                If IsDate(objMessage.Headers("Date")) Then
                    Dim dt = CDate(objMessage.Headers("Date"))
                    If dt = Date.MinValue Then
                        Return File.GetLastWriteTime(strFileName)
                    Else
                        Return dt
                    End If
                Else
                    Return File.GetLastWriteTime(strFileName)
                End If
            End Using
        Catch
            Return SqlTypes.SqlDateTime.MinValue.Value
        End Try
    End Function

#End Region

End Class

Public Class clsEdocRepository
    Implements IDisposable

    Private m_objDbContext As clsDB
    Private m_tempFileLocation As String
    Private m_colRecordLocks As K1Dictionary(Of clsRecordLock)
    Private m_objTable As clsTable

    Sub New(objDbContext As clsDB, tempFileLocation As String, Optional colRecordLocks As K1Dictionary(Of clsRecordLock) = Nothing)
        m_objDbContext = objDbContext
        m_objTable = m_objDbContext.SysInfo.Tables(clsDBConstants.Tables.cEDOC)

        m_tempFileLocation = tempFileLocation

        If (colRecordLocks Is Nothing) Then
            m_colRecordLocks = New K1Dictionary(Of clsRecordLock)
        Else
            m_colRecordLocks = colRecordLocks
        End If
    End Sub

    Public Function GetEdoc(edocId As Integer) As clsEDOC
        Try
            Return clsEDOC.GetItem(edocId, m_objDbContext)
        Catch ex As Exception
            Throw New DataException("Could not retrieve EDOC object.", ex)
        End Try
    End Function

    Public Function GetEdocSafely(edocId As Integer) As clsEDOC
        Try
            Return clsEDOC.GetItem(edocId, m_objDbContext)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' This will fetch the image file from the database and store it in a temp location.
    ''' </summary>
    ''' <param name="objEdoc">EDOC object that represents an EDOC record.</param>
    ''' <returns>temp file path and name</returns>
    ''' <remarks>Please please delete the temp file once you are done with it. Do not leave it behind.</remarks>
    Public Function GetEdocFile(objEdoc As clsEDOC) As String
        Dim fileName = GetFullQualifiedTempFileLocation(objEdoc)
        If (File.Exists(fileName)) Then
            File.SetAttributes(fileName, FileAttributes.Normal)
            Try
                File.Delete(fileName)
            Catch ex As Exception
                Throw New FileLoadException("Temp file already exists " & fileName & " and could not be over written.")
            End Try
        End If
        objEdoc.SaveImageToFile(fileName)
        Return fileName
    End Function

    ''' <summary>
    ''' This method updates the blob and related image natural metadata fields.
    ''' </summary>
    ''' <param name="edocId">EDOC Id that represents an EDOC record ID.</param>
    ''' <param name="objFileInfo">imaging document file to convert to blob.</param>
    ''' <returns>True for success</returns>
    ''' <remarks>Unlocking of the EDOC has to be done by the calling code if blnUseLocking is True</remarks>
    Public Function UpdateEdocFile(edocId As Integer,
                                   objFileInfo As FileInfo,
                                   Optional blnUseLocking As Boolean = False,
                                   Optional blnShowProgress As Boolean = False) As Boolean

        If (Not blnUseLocking AndAlso CheckRecordLock(edocId)) Then
            Return UpdateEdocFileInt(edocId, objFileInfo, blnShowProgress)
        End If

        If (blnUseLocking AndAlso LockRecord(edocId)) Then
            Return UpdateEdocFileInt(edocId, objFileInfo, blnShowProgress)
        End If

        Throw New DBConcurrencyException("EDOC record is locked by another session.")

    End Function

    ''' <summary>
    ''' Clears the EDOC image field
    ''' </summary>
    ''' <param name="edocId"></param>
    ''' <param name="blnUseLocking"></param>
    ''' <remarks>Unlocking of the EDOC has to be done by the calling code if blnUseLocking is True</remarks>
    Public Sub ClearEdocImage(edocId As Integer,
                              Optional blnUseLocking As Boolean = False)

        If (Not blnUseLocking AndAlso CheckRecordLock(edocId)) Then
            ClearEdocImageInt(edocId)
            Return
        End If

        If (blnUseLocking AndAlso LockRecord(edocId)) Then
            ClearEdocImageInt(edocId)
            UnlockRecord(edocId)
            Return
        End If

        Throw New DBConcurrencyException("EDOC record is locked by another user's session.")

    End Sub

    Public Function LockRecord(ByVal intRecordID As Integer) As Boolean
        Return String.IsNullOrEmpty(clsRecordLock.GetLock(m_objDbContext, m_colRecordLocks, m_objTable.ID, intRecordID))
    End Function

    Public Function CheckRecordLock(ByVal intRecordID As Integer) As Boolean
        Return clsRecordLock.CheckLock(m_objDbContext, m_colRecordLocks, m_objTable.ID, intRecordID)
    End Function

    Public Sub UnlockRecord(ByVal intRecordID As Integer)
        clsRecordLock.ReleaseLock(m_objDbContext, m_colRecordLocks, m_objTable.ID, intRecordID)
    End Sub

    Protected Sub ClearEdocImageInt(edocId As Integer)
        Dim intImageSize As Integer = m_objDbContext.ExecuteScalar("SELECT datalength([Image]) FROM EDOC WHERE ID=" & CStr(edocId))
        If intImageSize <= 1 Then
            Dim intDB As Integer = m_objDbContext.ExecuteScalar("SELECT DBID FROM K1Archive_Link WHERE EDOCID=" & CStr(edocId))

            If intDB > 0 Then
                Dim objArchiveDT As DataTable = m_objDbContext.GetDataTableBySQL("SELECT ExternalID FROM K1Archive WHERE ID=" & CStr(intDB))
                If objArchiveDT IsNot Nothing AndAlso objArchiveDT.Rows.Count = 1 Then
                    Dim objEncryption As New clsEncryption(True)
                    Dim objConnection As New SqlClient.SqlConnection(objEncryption.Decrypt(CStr(objArchiveDT.Rows(0)(0))))
                    objConnection.Open()
                    m_objDbContext.ExecuteSQL("DELETE FROM K1Archive_Link WHERE EDOCID=" & CStr(edocId))

                    Dim cmdDB As SqlClient.SqlCommand = objConnection.CreateCommand()
                    cmdDB.CommandText = "DELETE FROM EDOC WHERE ID=" & CStr(edocId)
                    cmdDB.CommandType = CommandType.Text
                    cmdDB.ExecuteNonQuery()
                End If
            End If
        Else
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@BinaryData", New Byte() {0}))
            colParams.Add(New clsDBParameter("@Id", edocId))
            Dim strCmd = String.Format("Update {0} set {1} = @BinaryData, {2} = @BinaryData where ID = @Id",
                                       m_objTable.DatabaseName,
                                       clsDBConstants.Fields.EDOC.cIMAGE,
                                       clsDBConstants.Fields.EDOC.cTHUMBNAIL)
            m_objDbContext.ExecuteSQL(strCmd, colParams)

            'Just in case the EDOC is sitting in the queue to be archived when deletd
            m_objDbContext.ExecuteSQL("DELETE FROM K1Archive_Queue WHERE EDOCID=" & CStr(edocId))
        End If
    End Sub

    Protected Function UpdateEdocFileInt(edocId As Integer,
                                         objFileInfo As FileInfo,
                                         Optional blnShowProgress As Boolean = False) As Boolean
        If (objFileInfo.Exists) Then
            File.SetAttributes(objFileInfo.FullName, FileAttributes.Normal)
            Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(m_objDbContext.SysInfo.Tables(clsDBConstants.Tables.cEDOC), edocId)
            UpdateEdocFileRelatedFields(colMasks, objFileInfo)
            colMasks.Update(m_objDbContext, False)

            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@BinaryData", New Byte() {0}))
            colParams.Add(New clsDBParameter("@Id", edocId))
            m_objDbContext.ExecuteSQL("UPDATE " & clsDBConstants.Tables.cEDOC & " SET " & clsDBConstants.Fields.EDOC.cIMAGE & " = @BinaryData WHERE ID = @Id",
                                      colParams)

            Dim objTable = m_objDbContext.SysInfo.Tables.Values.FirstOrDefault(Function(t) t.DatabaseName = clsDBConstants.Tables.cEDOC)
            Dim objMaskField = objTable.Fields.Values.FirstOrDefault(Function(f) f.DatabaseName = clsDBConstants.Fields.EDOC.cIMAGE)
            m_objDbContext.WriteBLOB(objTable, objMaskField, edocId, objFileInfo.FullName, blnShowProgress)
            Return True
        End If
        Return False
    End Function

    Protected Sub UpdateEdocFileRelatedFields(ByVal colMasks As clsMaskFieldDictionary,
                                        ByVal objFileInfo As FileInfo)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cFILENAME, Path.GetFileName(objFileInfo.FullName))
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCONTENTTYPE, GetMIMEType(objFileInfo.FullName))
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSIZE, objFileInfo.Length)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSUFFIX, Path.GetExtension(objFileInfo.FullName).Replace(".", String.Empty))

        Dim objFileSummary As New clsFileSummary(objFileInfo.FullName)
        If objFileSummary IsNot Nothing Then
            If Not String.IsNullOrEmpty(objFileSummary.Author) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cAUTHOR, objFileSummary.Author)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Category) Then
                colMasks.UpdateMaskObj("Category", objFileSummary.Category)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Comments) Then
                colMasks.UpdateMaskObj("Comments", objFileSummary.Comments)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.KeyWords) Then
                colMasks.UpdateMaskObj("KeyWords", objFileSummary.KeyWords)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Subject) Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cSUBJECT, objFileSummary.Subject)
            End If

            If Not String.IsNullOrEmpty(objFileSummary.Title) Then
                colMasks.UpdateMaskObj("Title", objFileSummary.Title)
            End If
        End If

    End Sub

    Protected Function GetFullQualifiedTempFileLocation(objEdoc As clsEDOC) As String
        Return Path.Combine(m_tempFileLocation, Path.GetFileNameWithoutExtension(objEdoc.FileName) & objEdoc.FullFileExtension)
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose

        Disposing(True)
        GC.SuppressFinalize(Me)

    End Sub

    Protected Overridable Sub Disposing(ByVal blnDisposing As Boolean)

        If (blnDisposing) Then
            m_objDbContext = Nothing
            m_tempFileLocation = Nothing
            m_colRecordLocks = Nothing
            m_objTable.Dispose()
            m_objTable = Nothing
        End If

    End Sub

End Class