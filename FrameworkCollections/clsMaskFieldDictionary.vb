Imports K1Library.DBObject
Imports K1Library.clsDBConstants

Public Class clsMaskFieldDictionary
    Inherits K1Library.FrameworkCollections.K1Dictionary(Of clsMaskField)

#Region " Members "

    Private m_objTable As clsTable
    Private m_objTableMask As clsTableMask
    Private m_dateCurrentServerDate As Date = Date.MinValue
#End Region

#Region " Enumerations "

    Public Enum enumMSSearchType
        ALL = 0
        EXTERNALID = 1
        MULTIPLESEQUENCEFIELDS = 2
        EXTERNALID_IF_NOT_MS_FIELDS = 3
    End Enum
#End Region

#Region " Properties "

    Public Property Table() As clsTable
        Get
            Return m_objTable
        End Get
        Set(ByVal Value As clsTable)
            m_objTable = Value
        End Set
    End Property

    ''' <summary>
    ''' This is the parent table mask if one exists
    ''' </summary>
    Public Property TableMask() As clsTableMask
        Get
            Return m_objTableMask
        End Get
        Set(ByVal value As clsTableMask)
            m_objTableMask = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the ID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property ID() As Integer
        Get
            If Me(clsDBConstants.Fields.cID) Is Nothing OrElse _
            Me(clsDBConstants.Fields.cID).Value1.Value Is Nothing Then
                Return clsDBConstants.cintNULL
            Else
                Return CInt(Me(clsDBConstants.Fields.cID).Value1.Value)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the ExternalID value in the Mask Field Collection (if one exists)
    ''' </summary>
    Public ReadOnly Property ExternalID() As String
        Get
            If Me(clsDBConstants.Fields.cEXTERNALID) Is Nothing OrElse _
            Me(clsDBConstants.Fields.cEXTERNALID).Value1.Value Is Nothing Then
                Return clsDBConstants.cstrNULL
            Else
                Return CStr(Me(clsDBConstants.Fields.cEXTERNALID).Value1.Value)
            End If
        End Get
    End Property

    Public ReadOnly Property CurrentServerDate() As Date
        Get
            If m_dateCurrentServerDate = Date.MinValue Then
                If m_objTable IsNot Nothing Then
                    m_dateCurrentServerDate = m_objTable.Database.GetCurrentTime
                Else
                    Return Now
                End If
            End If

            Return m_dateCurrentServerDate
        End Get
    End Property
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeObj()
        m_objTable = Nothing
        m_objTableMask = Nothing
    End Sub
#End Region

#Region " Overrides "

    'Protected Overrides Sub OnInsert(ByVal key As [Object], ByVal value As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    '    If Not TypeOf value Is clsMaskField Then
    '        Throw New ArgumentException("value must be of type clsMask.", "value")
    '    End If
    'End Sub 'OnInsert

    'Protected Overrides Sub OnRemove(ByVal key As [Object], ByVal value As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    'End Sub 'OnRemove

    'Protected Overrides Sub OnSet(ByVal key As [Object], ByVal oldValue As [Object], ByVal newValue As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    '    If Not TypeOf newValue Is clsMaskField Then
    '        Throw New ArgumentException("newValue must be of type clsMask.", "value")
    '    End If
    'End Sub 'OnSet

    'Protected Overrides Sub OnValidate(ByVal key As [Object], ByVal value As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    '    If Not TypeOf value Is clsMaskField Then
    '        Throw New ArgumentException("value must be of type clsMask.", "value")
    '    End If
    'End Sub 'OnValidate
#End Region

    Public Overloads Sub Add(ByVal value As clsMaskField)
        Dim strKey As String

        strKey = value.Field.DatabaseName.ToUpper
        Me.Add(strKey, value)
    End Sub 'Add

    Public Sub AddRange(ByVal value As IEnumerable(Of clsMaskField))
        Dim strKey As String

        For Each objMask As clsMaskField In value
            strKey = objMask.Field.DatabaseName.ToUpper
            Me.Add(strKey, objMask)
        Next
    End Sub 'Add

#Region " Insert "

    Public Function Insert(ByVal objDb As clsDB,
                           Optional ByVal blnUseAutoNumbering As Boolean = True,
                           Optional ByVal blnCreateAuditTrail As Boolean = True) As Integer

        Dim objParameter As clsDBParameter
        Dim intId As Integer = clsDBConstants.cintNULL
        Dim colParams As clsDBParameterDictionary
        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not objDb.HasTransaction Then
                objDb.BeginTransaction()
                blnCreatedTransaction = True
            End If

            CheckTrainingVersion(objDb)

            Try
                If blnUseAutoNumbering Then
                    For Each objMask As clsMaskField In Me.Values
                        If Not objMask.Value1.AutoNumber Is Nothing Then
                            objMask.Value1.AutoNumber.Generate(objMask.Value1)
                        End If
                    Next
                End If

                colParams = ExecuteStoredProcedure(objDb, clsDB_Direct.enumSPType.Insert)
            Catch ex As Exception
                If blnUseAutoNumbering Then
                    For Each objMask As clsMaskField In Me.Values
                        If objMask.Value1.AutoNumber IsNot Nothing Then
                            objMask.Value1.AutoNumber.RollBack(objMask.Value1)
                        End If
                    Next
                End If

                Throw
            End Try

            If colParams IsNot Nothing Then
                objParameter = colParams(clsDB_Direct.ParamName(clsDBConstants.Fields.cID))
                If Not objParameter Is Nothing Then
                    intId = CType(objParameter.Value, Integer)
                    Me(clsDBConstants.Fields.cID).Value1.InitializeValue(intId)
                End If
            End If

            Dim colFileNames As New Hashtable

            For Each objMask As clsMaskField In From objMask1 In Values
                                                Where objMask1.Field.IsBinaryType AndAlso Not objMask1.Value1.FileName Is Nothing

                If objDb.ThreadedOperationCancelled Then
                    Throw New clsK1Exception("User cancelled save operation.")
                End If

                objDb.WriteBLOB(m_objTable, objMask.Field, intId, objMask.Value1.FileName)

                If objDb.ThreadedOperationCancelled Then
                    Throw New clsK1Exception("User cancelled save operation.")
                End If

                If objMask.Value1.UseOCR Then
                    Dim objWorkItem As New clsOcrWorkItem With
                            {
                            .EdocId = intId,
                            .SubmittedDate = Date.Now,
                            .FileFormat = [Enum].GetName(GetType(enumImageTypes), objMask.Value1.ImageType),
                            .LanguageId = objMask.Value1.ModiArguments.Language,
                            .TextOption = [Enum].GetName(GetType(TextMappedOptions), objMask.Value1.ModiArguments.TextMappedPDF),
                            .AutoRotate = objMask.Value1.ModiArguments.WithAutoRotation
                            }
                    Dim owiRepo As clsOcrWorkItemRepository = New clsOcrWorkItemRepository(objDb, Path.GetDirectoryName(objMask.Value1.FileName))
                    owiRepo.InsertWorkItem(objWorkItem)
                End If

                colFileNames.Add(objMask.Field.DatabaseName, objMask.Value1.FileName)
            Next

            If blnCreateAuditTrail Then

                Dim objAtd As clsAuditTrailRecordData = Nothing

                If clsAuditTrail.AuditTableMethodData(objDb, clsMethod.enumMethods.cADD, m_objTable) Then

                    Dim objTableMask As New clsTableMask(m_objTable, clsTableMask.enumMaskType.VIEW, intId)

                    If Me.TableMask IsNot Nothing Then
                        For Each objMaskFieldLink As clsMaskFieldLink In objTableMask.MaskManyToManyCollection.Values
                            Dim objParentFieldLink As clsMaskFieldLink = m_objTableMask.MaskManyToManyCollection(CStr(objMaskFieldLink.FieldLink.ID))

                            objMaskFieldLink.LoadValues(objTableMask.ID)

                            If objMaskFieldLink.IDCollection.Count = 0 Then
                                objMaskFieldLink.IDCollection = CType(objParentFieldLink.NewIDCollection.Clone, Hashtable)
                            Else
                                For Each intLinkedID As Integer In objParentFieldLink.NewIDCollection.Values
                                    If objMaskFieldLink.IDCollection(CStr(intLinkedID)) Is Nothing Then
                                        objMaskFieldLink.IDCollection.Add(CStr(intLinkedID), intLinkedID)
                                    End If
                                Next
                            End If
                        Next
                    End If

                    If colFileNames.Count > 0 Then
                        For Each strKey As String In colFileNames.Keys
                            Try
                                objTableMask.MaskFieldCollection(strKey).Value1.FileName = CType(colFileNames(strKey), String)
                            Catch ex As Exception

                            End Try
                        Next
                    End If

                    objAtd = New clsAuditTrailRecordData(clsMethod.enumMethods.cADD, objTableMask, Nothing)
                End If

                clsAuditTrail.CreateTableMethodRecord(objDb, clsMethod.enumMethods.cADD, m_objTable, intId, ExternalID, objAtd)

            End If

            For Each objMask As clsMaskField In Me.Values
                objMask.Value1.SequentialNumber = clsDBConstants.cintNULL
            Next

            If blnCreatedTransaction Then
                objDb.EndTransaction(True)
            End If

        Catch ex As Exception

            If blnCreatedTransaction Then
                objDb.EndTransaction(False)
            End If

            Throw

        End Try

        Return intId
    End Function

    Private Sub CheckTrainingVersion(ByVal objDb As clsDB)
        If objDb.IsTrainingVersion AndAlso
            Not (m_objTable.ClassType = clsDBConstants.enumTableClass.SYSTEM_LOCKED OrElse
                 m_objTable.DatabaseName = clsDBConstants.Tables.cCAPTION OrElse
                 m_objTable.DatabaseName = clsDBConstants.Tables.cSTRING OrElse
                 m_objTable.DatabaseName = clsDBConstants.Tables.cLANGUAGESTRING OrElse
                 m_objTable.DatabaseName = clsDBConstants.Tables.cTABLEMETHOD) Then
            If objDb.RecordCountExceeded(m_objTable.DatabaseName) Then
                Throw New Exception("You cannot insert any more records into this training database. " &
                                    "The maximum number of rows (" & objDb.RecordLimit & ") has been reached or exceeded for table '" &
                                    m_objTable.DatabaseName & "'.")
            End If
        End If
    End Sub

#End Region

#Region " Update "

    Public Sub Update(ByVal objDb As clsDB, Optional ByVal blnCreateAuditTrail As Boolean = True)

        Dim blnCreatedTransaction As Boolean = False

        Dim objOriginalRecord As clsTableMask = Nothing

        Try
            If Not objDb.HasTransaction Then
                objDb.BeginTransaction()
                blnCreatedTransaction = True
            End If

            Try
                For Each objMask As clsMaskField In Me.Values
                    If Not objMask.Value1.AutoNumber Is Nothing AndAlso objMask.AutonumberUpdated Then
                        objMask.Value1.AutoNumber.Generate(objMask.Value1)
                    End If
                Next

                If blnCreateAuditTrail AndAlso clsAuditTrail.AuditTableMethodData(m_objTable.Database, clsMethod.enumMethods.cMODIFY, m_objTable) Then

                    'Load the original values
                    objOriginalRecord = New clsTableMask(m_objTable, clsTableMask.enumMaskType.VIEW, ID)

                    For Each objMaskFieldLink As clsMaskFieldLink In objOriginalRecord.MaskManyToManyCollection.Values
                        If TableMask IsNot Nothing Then
                            Dim objParentFieldLink As clsMaskFieldLink = m_objTableMask.MaskManyToManyCollection(CStr(objMaskFieldLink.FieldLink.ID))

                            objParentFieldLink.LoadValues(m_objTableMask.ID)

                            objMaskFieldLink.ValuesLoaded = True
                            objMaskFieldLink.IDCollection = CType(objParentFieldLink.IDCollection.Clone, Hashtable)
                        End If
                    Next

                    Dim originalMaskImageFields = objOriginalRecord.MaskFieldCollection.
                        Where(Function(mf) mf.Value.Field.IsBinaryType).
                        Select(Function(mf) mf.Value)

                    For Each oMaskImgField As clsMaskField In originalMaskImageFields

                        Dim strTempFolder = Path.GetTempPath()
                        '[Naing] Lets check if the Blobs needs to be included in the AuditTrail log
                        Dim currentBlobMaskField = Me(oMaskImgField.Field.DatabaseName)

                        If (currentBlobMaskField.Value1 IsNot Nothing AndAlso
                            currentBlobMaskField.Value1.IsDirty) Then

                            '[Naing] Must load the original files from the database to disk for further processing.
                            Dim strSubPath = String.Format("Knowledgeone\Client\{0}", Path.GetRandomFileName())
                            oMaskImgField.Value1.FileName = Path.Combine(strTempFolder, strSubPath)
                            objDb.ReadBLOB(objOriginalRecord.Table.DatabaseName, oMaskImgField.Field.DatabaseName, ID, oMaskImgField.Value1.FileName, False)

                        End If

                    Next

                End If

                'Execute the update command
                ExecuteStoredProcedure(objDb, clsDB_Direct.enumSPType.Update)

            Catch ex As Exception
                For Each objMask As clsMaskField In Me.Values
                    If Not objMask.Value1.AutoNumber Is Nothing AndAlso objMask.AutonumberUpdated Then
                        If Not objMask.Value1.RollbackValue Is Nothing Then
                            objMask.Value1.InitializeValue(objMask.Value1.RollbackValue)
                        End If
                        objMask.Value1.RollbackValue = Nothing
                    End If
                Next

                Throw
            End Try

            If blnCreateAuditTrail Then

                Dim objAtd As clsAuditTrailRecordData = Nothing

                If objOriginalRecord IsNot Nothing Then

                    Dim objNewRecord As New clsTableMask(m_objTable, clsTableMask.enumMaskType.VIEW, ID)

                    objAtd = New clsAuditTrailRecordData(clsMethod.enumMethods.cMODIFY, objOriginalRecord, Nothing)
                    objAtd.NewRecord = objNewRecord

                    If Me.TableMask IsNot Nothing Then
                        For Each objMaskFieldLink As clsMaskFieldLink In objNewRecord.MaskManyToManyCollection.Values
                            Dim objParentFieldLink As clsMaskFieldLink = m_objTableMask.MaskManyToManyCollection(CStr(objMaskFieldLink.FieldLink.ID))
                            Dim objOriginalFieldLink As clsMaskFieldLink = objOriginalRecord.MaskManyToManyCollection(CStr(objMaskFieldLink.FieldLink.ID))

                            objMaskFieldLink.LoadValues(objNewRecord.ID)

                            'just in case record were created in a trigger (as in the case of edoc workflow)
                            If objParentFieldLink.NewIDCollection.Count > 0 Then
                                If objMaskFieldLink.IDCollection.Count = 0 Then
                                    objMaskFieldLink.IDCollection = CType(objParentFieldLink.NewIDCollection.Clone, Hashtable)
                                Else
                                    For Each intId As Integer In objParentFieldLink.IDCollection.Values
                                        If objMaskFieldLink.IDCollection(CStr(intId)) IsNot Nothing Then
                                            objMaskFieldLink.IDCollection.Remove(CStr(intId))
                                        End If
                                    Next

                                    For Each intID As Integer In objParentFieldLink.NewIDCollection.Values
                                        If objMaskFieldLink.IDCollection(CStr(intID)) Is Nothing Then
                                            objMaskFieldLink.IDCollection.Add(CStr(intID), intID)
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If

                clsAuditTrail.CreateTableMethodRecord(m_objTable.Database, clsMethod.enumMethods.cMODIFY,
                                                      m_objTable, ID, ExternalID,
                                                      objAtd)
            End If

            For Each objMask As clsMaskField In Me.Values
                objMask.Value1.SequentialNumber = clsDBConstants.cintNULL
            Next

            If blnCreatedTransaction Then
                objDb.EndTransaction(True)
            End If

        Catch ex As Exception

            If blnCreatedTransaction Then
                objDb.EndTransaction(False)
            End If

            Throw

        End Try
    End Sub
#End Region

#Region " Delete "

    Public Sub Delete(ByVal objDB As clsDB, Optional ByVal blnCreateAuditTrail As Boolean = True)

        Dim blnCreatedTransaction As Boolean = False
        Dim objOriginalRecord As clsTableMask = Nothing

        Try
            If Not objDB.HasTransaction Then
                objDB.BeginTransaction()
                blnCreatedTransaction = True
            End If

            If blnCreateAuditTrail Then
                If clsAuditTrail.AuditTableMethodData(m_objTable.Database, clsMethod.enumMethods.cDELETE, m_objTable) Then

                    objOriginalRecord = New clsTableMask(m_objTable, clsTableMask.enumMaskType.VIEW, ID)

                    For Each objMaskField As clsMaskField In objOriginalRecord.MaskFieldCollection.Values
                        If objMaskField.Field.IsBinaryType Then

                            Dim strFile As String = Nothing

                            If Me.TableMask IsNot Nothing Then
                                strFile = Me.TableMask.MaskFieldCollection(objMaskField.Field.DatabaseName).Value1.FileName
                            End If

                            If String.IsNullOrEmpty(strFile) Then

                                Dim strSubPath = String.Format("\Knowledgeone\Client\{0}", Path.GetRandomFileName())
                                strFile = Path.Combine(Path.GetTempPath(), strSubPath)

                                objMaskField.Database.ReadBLOB(objMaskField.Field.Table.DatabaseName, _
                                    objMaskField.Field.DatabaseName, objMaskField.MaskFieldCollection.ID, strFile)
                            End If

                            objMaskField.Value1.FileName = strFile

                        End If
                    Next
                End If
            End If

            ExecuteStoredProcedure(objDB, clsDB_Direct.enumSPType.Delete)

            If blnCreateAuditTrail Then
                Dim objAtd As clsAuditTrailRecordData = Nothing

                If objOriginalRecord IsNot Nothing Then
                    objAtd = New clsAuditTrailRecordData(clsMethod.enumMethods.cDELETE, objOriginalRecord, Nothing)
                End If

                clsAuditTrail.CreateTableMethodRecord(m_objTable.Database, _
                    clsMethod.enumMethods.cDELETE, m_objTable, ID, ExternalID, objAtd)
            End If

            If blnCreatedTransaction Then
                objDB.EndTransaction(True)
            End If
        Catch ex As Exception
            If blnCreatedTransaction Then
                objDB.EndTransaction(False)
            End If

            Throw
        End Try
    End Sub
#End Region

    Private Function ExecuteStoredProcedure(ByVal objDb As clsDB,
                                            ByVal eType As clsDB_Direct.enumSPType) As clsDBParameterDictionary
        Dim strStoredProcedure As String
        Dim colParams As clsDBParameterDictionary

        Try
            strStoredProcedure = m_objTable.DatabaseName

            Select Case eType
                Case clsDB_Direct.enumSPType.Insert
                    strStoredProcedure &= clsDBConstants.StoredProcedures.cINSERT
                Case clsDB_Direct.enumSPType.Update
                    strStoredProcedure &= clsDBConstants.StoredProcedures.cUPDATE
                Case clsDB_Direct.enumSPType.Delete
                    strStoredProcedure &= clsDBConstants.StoredProcedures.cDELETE
            End Select

            colParams = clsDB_Direct.GetParamCollection(eType, Me)
            objDb.ExecuteProcedure(strStoredProcedure, colParams)
        Catch ex As Exception
            Throw
        End Try

        Return colParams
    End Function

#Region " Get/Update Mask Object Values "

    Private Function CheckValueIsNothing(ByVal objMask As clsMaskField, ByVal objValue As Object) As Object
        Try
            If Not objValue Is Nothing Then
                Select Case objMask.Field.DataType
                    Case SqlDbType.BigInt, SqlDbType.Bit, _
                    SqlDbType.Int, SqlDbType.Real, _
                     SqlDbType.SmallInt, SqlDbType.TinyInt
                        If CInt(objValue) = clsDBConstants.cintNULL Then
                            Return Nothing
                        End If

                    Case SqlDbType.Decimal, SqlDbType.SmallMoney, _
                    SqlDbType.Float, SqlDbType.Money
                        If CDbl(objValue) = clsDBConstants.cintNULL Then
                            Return Nothing
                        End If

                    Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NText, SqlDbType.NVarChar, _
                    SqlDbType.Text, SqlDbType.VarChar
                        If CStr(objValue) = clsDBConstants.cstrNULL Then
                            Return Nothing
                        End If

                    Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                        If CDate(objValue) = Date.MinValue Then
                            Return Nothing
                        End If
                End Select
            End If
        Catch ex As Exception
            Return objValue
        End Try

        Return objValue
    End Function

    ''' <summary>
    ''' Set the Mask Value
    ''' </summary>
    ''' <param name="strField"></param>
    ''' <param name="objValue"></param>
    ''' <param name="eCheckState"></param>
    ''' <remarks>Should be named SetMaskValue. Grrr ... who ever added this dependency to [Windows.Forms.CheckState] must be shot !!!</remarks>
    Public Sub UpdateMaskObj(ByVal strField As String,
                             ByVal objValue As Object,
                             Optional ByVal eCheckState As Windows.Forms.CheckState = Windows.Forms.CheckState.Unchecked)
        Dim objMask As clsMaskField = Me(strField)
        If Not objMask Is Nothing Then
            objMask.Value1.Value = CheckValueIsNothing(objMask, objValue)
            objMask.CheckState = eCheckState
        End If
    End Sub

    Public Function GetMaskValue(ByVal strField As String,
                                 Optional ByVal objDefault As Object = Nothing) As Object
        Dim objMask As clsMaskField = Me(strField)
        If objMask Is Nothing Then
            Return objDefault
        Else
            If (objMask.Value1.Value Is Nothing) Then
                Return objDefault
            Else
                Return objMask.Value1.Value
            End If
        End If
    End Function

    'use this when we know its a string
    Public Function GetStringValue(ByVal strField As String) As String
        Dim objMask As clsMaskField = Me(strField)
        If objMask Is Nothing Then
            Return String.Empty
        Else
            If (objMask.Value1.Value Is Nothing) Then
                Return String.Empty
            Else
                Return objMask.Value1.Value.ToString()
            End If
        End If
    End Function

#End Region

#Region " Auto number Functions "

    Public Sub FillAutoNumberFormats(ByVal blnPopulateFields As Boolean)
        Dim blnLoadMS As Boolean = True

        Dim objMultipleSequenceMask As clsMaskField = Nothing

        For Each objMask As clsMaskField In Me.Values
            If Not objMask.Value1.AutoNumber Is Nothing AndAlso _
            objMask.Value1.AutoNumber.IsMultipleSequenceAutoNumber Then
                objMultipleSequenceMask = objMask
                Exit For
            End If
        Next

        If objMultipleSequenceMask Is Nothing Then
            Return
        End If

        Dim objDB As clsDB = objMultipleSequenceMask.Database

        For Each objMask As clsMaskField In Me.Values
            If objMask.IsMultipleSequenceField AndAlso _
            objMask.DeterminesMultipleSequence AndAlso _
            Not CompareMSField(objDB, objMask.Field) Then
                blnLoadMS = False
            End If
        Next

        If blnLoadMS Then
            Dim intID As Integer = clsDBConstants.cintNULL
            Dim intNewID As Integer = clsDBConstants.cintNULL

            If objMultipleSequenceMask.Value1.MultipleSequence IsNot Nothing Then
                intID = CInt(objMultipleSequenceMask.Value1.MultipleSequence.GetMaskValue( _
                    clsDBConstants.Fields.cID))
            End If

            Dim objMS As clsMaskFieldDictionary = GetMultipleSequenceRecord(objDB, _
                objMultipleSequenceMask.Value1, enumMSSearchType.MULTIPLESEQUENCEFIELDS)

            If objMS IsNot Nothing Then
                intNewID = CInt(objMS.GetMaskValue(clsDBConstants.Fields.cID))
            End If

            If Not intID = intNewID Then
                If intNewID = clsDBConstants.cintNULL Then
                    objMultipleSequenceMask.Value1.Value = Nothing
                    objMultipleSequenceMask.Value1.MultipleSequence = Nothing
                Else
                    Dim strCurrentValue As String = Nothing
                    If objMultipleSequenceMask.Value1.Value IsNot Nothing Then
                        strCurrentValue = CStr(objMultipleSequenceMask.Value1.Value)
                    End If

                    objMultipleSequenceMask.Value1.Value = objMultipleSequenceMask.Value1.AutoNumber.GetMaskValue(CType(objMS.GetMaskValue(clsDBConstants.Fields.cEXTERNALID), String),
                                                                                                                  strCurrentValue)
                    objMultipleSequenceMask.Value1.MultipleSequence = objMS

                    If blnPopulateFields Then
                        'we need to remove or assign values
                        For Each objMask As clsMaskField In Me.Values
                            If objMask.IsMultipleSequenceField AndAlso _
                            CompareMSField(objDB, objMask.Field) Then

                                Dim objNewValue As Object = objMS.GetMaskValue(objMask.Field.DatabaseName)
                                Dim blnNewValue As Boolean = False

                                If objMask.Value1.Value Is Nothing Then
                                    objMask.Value1.Value = objNewValue
                                    blnNewValue = True
                                End If

                                If blnNewValue AndAlso objMask.Field.IsForeignKey Then
                                    clsMaskField.LoadLinkedData(objMask, objMask.Field.FieldLink.IdentityTable)
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Else
            objMultipleSequenceMask.Value1.Value = Nothing
            objMultipleSequenceMask.Value1.MultipleSequence = Nothing
        End If
    End Sub

    Public Shared Function GetMultipleSequenceRecord(ByVal objDb As clsDB,
                                                     ByVal objMaskValue As clsMaskFieldValue,
                                                     ByVal eType As enumMSSearchType) As clsMaskFieldDictionary

        Dim colMasks As clsMaskFieldDictionary = Nothing
        Dim objTable As clsTable = objDb.SysInfo.Tables(Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE)

        Dim objFilter = CreateSearchFilter(objDb, objMaskValue, eType)

        Dim objSelectInfo As New clsSelectInfo(objTable,
                                               Nothing,
                                               objFilter,
                                               colrecordstart:=Nothing, colRecordEnd:=Nothing)

        If objSelectInfo.DataTable IsNot Nothing AndAlso objSelectInfo.DataTable.Rows.Count = 1 Then

            Dim intID As Integer = CInt(objSelectInfo.DataTable.Rows(0)(0))
            colMasks = clsMaskField.CreateMaskCollection(objTable, intID)

        End If

        Return colMasks

    End Function

    Public Shared Function CreateSearchFilter(ByVal objDB As clsDB,
                                              ByVal objMaskValue As clsMaskFieldValue,
                                              ByVal eType As enumMSSearchType) As clsSearchFilter
        Dim objSF As clsSearchFilter
        Dim colSOs As New List(Of clsSearchObjBase)

        Dim objTable As clsTable = objDB.SysInfo.Tables(Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE)

        Dim objSE As New clsSearchElement(clsSearchFilter.enumOperatorType.NONE,
                                          objTable.DatabaseName & "." & Fields.AutoNumberFormatMultipleSequence.cAUTONUMBERFORMATID,
                                          clsSearchFilter.enumComparisonType.EQUAL,
                                          objMaskValue.AutoNumber.ID)

        colSOs.Add(objSE)

        Dim blnMSField As Boolean = False
        If Not eType = enumMSSearchType.EXTERNALID Then
            For Each objMask As clsMaskField In objMaskValue.MaskField.MaskFieldCollection.Values
                If objMask.IsMultipleSequenceField AndAlso _
                objMask.DeterminesMultipleSequence AndAlso _
                Not objMask.Value1.Value Is Nothing AndAlso _
                CompareMSField(objDB, objMask.Field) Then
                    Dim objValue As Object = objMask.Value1.Value

                    If TypeOf objMask.Value1.Value Is String Then
                        objValue = CStr(objValue) & "*"
                    End If

                    objSE = New clsSearchElement(clsSearchFilter.enumOperatorType.AND, _
                        objTable.DatabaseName & "." & objMask.Field.DatabaseName, _
                        clsSearchFilter.enumComparisonType.EQUAL, objValue)

                    colSOs.Add(objSE)

                    blnMSField = True
                End If
            Next
        End If

        If objMaskValue.Value IsNot Nothing AndAlso
            eType = enumMSSearchType.ALL OrElse
            eType = enumMSSearchType.EXTERNALID OrElse
            (eType = enumMSSearchType.EXTERNALID_IF_NOT_MS_FIELDS AndAlso Not blnMSField) Then

            Dim strValue As String = objMaskValue.AutoNumber.GetMultipleSequenceValue(CType(objMaskValue.Value, String))

            objSE = New clsSearchElement(clsSearchFilter.enumOperatorType.AND,
                                         objTable.DatabaseName & "." & Fields.cEXTERNALID,
                                         clsSearchFilter.enumComparisonType.EQUAL, strValue & "*")

            colSOs.Add(objSE)

        End If

        Dim objSG As New clsSearchGroup(clsSearchFilter.enumOperatorType.NONE, colSOs)
        objSF = New clsSearchFilter(objDB, objSG, objTable.DatabaseName)

        Return objSF

    End Function

    Public Shared Function CompareMSField(ByVal objDB As clsDB, ByVal objField As clsField) As Boolean
        Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE)

        Dim objField2 As clsField = objDB.SysInfo.Fields( _
            objTable.ID & "_" & objField.DatabaseName)

        If objField2 Is Nothing Then
            Return False
        End If

        If Not objField.DataType = objField2.DataType Then
            Return False
        End If

        If objField.IsForeignKey AndAlso _
        (Not objField2.IsForeignKey OrElse _
        Not objField.FieldLink.IdentityTable.ID = objField.FieldLink.IdentityTable.ID) Then
            Return False
        End If

        Return True
    End Function
#End Region

End Class
