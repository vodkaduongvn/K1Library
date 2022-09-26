Imports System.Xml
Imports System.Text

Public Class clsAuditTrailRecordData

    Private Const cVERSION As Single = 1.0

    Private m_objFilter As clsSearchFilter
    Private m_objOriginalRecord As clsTableMask
    Private m_objNewRecord As clsTableMask
    Private ReadOnly m_eMethod As clsMethod.enumMethods
    Private m_includeBinaryData As Boolean = True
    
    Public Sub New(Optional ByVal eMethod As clsMethod.enumMethods = clsMethod.enumMethods.cSEARCH)
        'm_blnDeleteFile = False
        m_eMethod = eMethod
    End Sub

    Public Sub New(ByVal eMethod As clsMethod.enumMethods, ByVal objRecordData As clsTableMask, ByVal objFilter As clsSearchFilter)
        'm_blnDeleteFile = False
        m_objOriginalRecord = objRecordData
        m_objFilter = objFilter
        m_eMethod = eMethod
    End Sub

    ''' <summary>
    ''' Search filter used by the user.
    ''' </summary>
    Public Property Filter() As clsSearchFilter
        Get
            Return m_objFilter
        End Get
        Set(ByVal value As clsSearchFilter)
            m_objFilter = value
        End Set
    End Property

    ''' <summary>
    ''' Original values of the record.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OriginalRecord() As clsTableMask
        Get
            Return m_objOriginalRecord
        End Get
        Set(ByVal value As clsTableMask)
            m_objOriginalRecord = value
        End Set
    End Property

    ''' <summary>
    ''' The new values for the record (update only)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NewRecord() As clsTableMask
        Get
            Return m_objNewRecord
        End Get
        Set(ByVal value As clsTableMask)
            m_objNewRecord = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the XML representation of this AuditTrailRecordData
    ''' </summary>
    ''' <returns>A System.String of XML</returns>
    Public Overrides Function ToString() As String
        Try

            Dim objSB As New StringBuilder

            Using objSW As New StringWriter(objSB)
                Using objXW As New XmlTextWriter(objSW)
                    SerializeToXml(objXW)
                End Using
            End Using

            Return objSB.ToString

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Serializes the record data to XML and saves to a file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SerializeToXml(Optional blnIncludeBinaryData As Boolean = True) As String
        Try

            m_IncludeBinaryData = blnIncludeBinaryData

            Dim strTempFile As String = IO.Path.GetTempFileName

            Using objXw As New XmlTextWriter(strTempFile, Nothing)
                SerializeToXml(objXw)
                objXw.Close()
            End Using

            Return strTempFile
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub SerializeToXml(ByVal objXmlWriter As XmlTextWriter)

        objXmlWriter.WriteStartElement("RecordData")
        objXmlWriter.WriteAttributeString("version", CType(cVERSION, String))

        If Me.Filter IsNot Nothing Then
            objXmlWriter.WriteStartElement("Filter")
            objXmlWriter.WriteValue(Me.Filter.UserSyntax)
            objXmlWriter.WriteEndElement()
        End If

        If Me.OriginalRecord IsNot Nothing Then
            Dim colMaskBase As FrameworkCollections.K1Collection(Of clsMaskBase) = GetMaskCollection(m_objOriginalRecord)

            If m_objNewRecord Is Nothing Then
                objXmlWriter.WriteStartElement("Record")

                For intLoop As Integer = 0 To colMaskBase.Count - 1
                    Dim objMask As clsMaskBase = colMaskBase(intLoop)

                    If TypeOf objMask Is clsMaskField Then
                        Dim objMaskField As clsMaskField = CType(objMask, clsMaskField)

                        objXmlWriter.WriteStartElement("Field")
                        objXmlWriter.WriteAttributeString("caption", objMaskField.Caption)
                        objXmlWriter.WriteAttributeString("dbname", objMaskField.Field.DatabaseName)
                        objXmlWriter.WriteStartElement("Value")

                        SerializeValue(objXmlWriter, objMaskField)

                        objXmlWriter.WriteEndElement()
                        objXmlWriter.WriteEndElement()
                    Else
                        If objMask.ObjectType = clsMaskBase.enumMaskObjectType.MANYTOMANY Then
                            Dim objMaskFieldLink As clsMaskFieldLink = CType(objMask, clsMaskFieldLink)

                            objXmlWriter.WriteStartElement("FieldLink")
                            objXmlWriter.WriteAttributeString("caption", objMaskFieldLink.Caption)
                            objXmlWriter.WriteAttributeString("dbname", objMaskFieldLink.FieldLink.ForeignKeyTable.DatabaseName & _
                                "." & objMaskFieldLink.FieldLink.ForeignKeyField.DatabaseName)

                            objMaskFieldLink.LoadValues(m_objOriginalRecord.ID)

                            If objMaskFieldLink.IDCollection IsNot Nothing AndAlso objMaskFieldLink.IDCollection.Count > 0 Then
                                Dim objDT As DataTable = objMaskFieldLink.Database.GetDataTableBySQL( _
                                    "SELECT [" & clsDBConstants.Fields.cID & "]," & _
                                    "[" & clsDBConstants.Fields.cEXTERNALID & "] " & _
                                    "FROM [" & objMaskFieldLink.FieldLink.LinkedTable.DatabaseName & "] " & _
                                    "WHERE [" & clsDBConstants.Fields.cID & "] IN (" & _
                                    CreateIDStringFromCollection(objMaskFieldLink.IDCollection.Values) & ")")

                                objDT.DefaultView.Sort = clsDBConstants.Fields.cEXTERNALID

                                For intIndex As Integer = 0 To objDT.DefaultView.Count - 1
                                    objXmlWriter.WriteStartElement("Value")
                                    objXmlWriter.WriteAttributeString("id", CType(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cID), String))
                                    objXmlWriter.WriteValue(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cEXTERNALID))
                                    objXmlWriter.WriteEndElement()
                                Next

                                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                                If objDT IsNot Nothing Then
                                    objDT.Dispose()
                                    objDT = Nothing
                                End If
                            End If

                            objXmlWriter.WriteEndElement()
                        End If
                    End If
                Next

                objXmlWriter.WriteEndElement()
            Else
                objXmlWriter.WriteStartElement("RecordChanges")
                objXmlWriter.WriteAttributeString("id", CType(m_objOriginalRecord.ID, String))
                objXmlWriter.WriteAttributeString("externalid", m_objOriginalRecord.ExternalID)

                For intLoop As Integer = 0 To colMaskBase.Count - 1
                    Dim objMask As clsMaskBase = colMaskBase(intLoop)

                    If TypeOf objMask Is clsMaskField Then
                        Dim objMaskField1 As clsMaskField = CType(objMask, clsMaskField)
                        Dim objMaskField2 As clsMaskField = m_objNewRecord.MaskFieldCollection(objMaskField1.Field.DatabaseName)

                        CompareAndCreateNode(objXmlWriter, objMaskField1, objMaskField2)
                    Else
                        If objMask.ObjectType = clsMaskBase.enumMaskObjectType.MANYTOMANY Then
                            Dim objMaskFieldLink1 As clsMaskFieldLink = CType(objMask, clsMaskFieldLink)
                            Dim objMaskFieldLink2 As clsMaskFieldLink = m_objNewRecord.MaskManyToManyCollection(CStr(objMaskFieldLink1.FieldLink.ID))

                            objMaskFieldLink1.LoadValues(m_objOriginalRecord.ID)
                            objMaskFieldLink2.LoadValues(m_objNewRecord.ID)

                            CompareAndCreateNode(objXmlWriter, objMaskFieldLink1, objMaskFieldLink2)
                        End If
                    End If
                Next

                objXmlWriter.WriteEndElement()
            End If
        End If

        objXmlWriter.WriteEndElement()

    End Sub

    Private Function GetMaskCollection(ByVal objTM As clsTableMask) As FrameworkCollections.K1Collection(Of clsMaskBase)
        Dim colMaskBase As New FrameworkCollections.K1Collection(Of clsMaskBase)

        Dim objDV As DataView = clsTableMask.GetMaskOrderDataView(objTM.Table, True, False, objTM.TypeID, True)

        For intLoop As Integer = 0 To objDV.Count - 1
            Dim objRow As DataRow = objDV(intLoop).Row
            Dim eObjType As clsMaskBase.enumMaskObjectType = _
                CType(objRow(clsTableMask.Columns.cCOL_OBJTYPE),  _
                clsMaskBase.enumMaskObjectType)
            Dim strKey As String = CType(objRow(clsTableMask.Columns.cCOL_KEY), String)

            Select Case eObjType
                Case clsMaskBase.enumMaskObjectType.FIELD
                    Dim objField As clsField = objTM.Table.Fields(strKey)
                    colMaskBase.Add(objTM.MaskFieldCollection(objField.DatabaseName))

                Case clsMaskBase.enumMaskObjectType.ONETOMANY
                    Dim objFieldLink As clsFieldLink = objTM.Table.OneToManyLinks(strKey)
                    colMaskBase.Add(objTM.MaskOneToManyCollection(objFieldLink.KeyID))

                Case clsMaskBase.enumMaskObjectType.MANYTOMANY
                    Dim objFieldLink As clsFieldLink = objTM.Table.ManyToManyLinks(strKey)
                    colMaskBase.Add(objTM.MaskManyToManyCollection(objFieldLink.KeyID))

            End Select
        Next

        Return colMaskBase
    End Function

    Private Sub SerializeValue(ByVal objXW As XmlTextWriter, ByVal objMaskField As clsMaskField)

        If objMaskField.Field.IsBinaryType AndAlso m_IncludeBinaryData Then
            Dim strFile As String = String.Empty

            If objMaskField.Value1.FileName IsNot Nothing Then
                strFile = objMaskField.Value1.FileName
            End If

            If Not (String.IsNullOrEmpty(strFile)) AndAlso System.IO.File.Exists(strFile) Then
                Dim objTempFS As New IO.FileStream(strFile, _
                    IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)

                Dim intBuffer As Integer = 4096
                Dim arrBuffer(intBuffer - 1) As Byte
                Dim intReadByte As Integer = 0

                Dim objBR As New IO.BinaryReader(objTempFS)

                Do
                    intReadByte = objBR.Read(arrBuffer, 0, intBuffer)
                    objXW.WriteBase64(arrBuffer, 0, intReadByte)
                Loop While (intBuffer <= intReadByte)

                objTempFS.Close()
                objTempFS.Dispose()
                objTempFS = Nothing

                If m_eMethod = clsMethod.enumMethods.cDELETE Then
                    Try
                        IO.File.Delete(strFile)
                    Catch ex As Exception
                    End Try
                End If
            End If

            Return
        End If

        If objMaskField.Value1.Value Is Nothing Then
            Return
        End If

        If objMaskField.Field.IsForeignKey Then
            objXW.WriteAttributeString("id", CType(objMaskField.Value1.Value, String))
            objXW.WriteValue(objMaskField.Value1.Display)
            Return
        End If

        If objMaskField.Field.IsDateType Then
            objXW.WriteValue(CDate(objMaskField.Value1.Value))
        Else
            objXW.WriteValue(objMaskField.Value1.Value)
        End If
    End Sub

    Private Sub CompareAndCreateNode(ByVal objXW As XmlTextWriter, _
    ByVal objMaskField1 As clsMaskField, ByVal objMaskField2 As clsMaskField)
        If objMaskField1.Value1.Value Is Nothing AndAlso _
        objMaskField2.Value1.Value Is Nothing Then
            Return
        End If

        If objMaskField1.Value1.Value Is Nothing Then
            SerializeValues(objXW, objMaskField1, objMaskField2)
            Return
        End If

        If objMaskField2.Value1.Value Is Nothing Then
            SerializeValues(objXW, objMaskField1, objMaskField2)
            Return
        End If

        Dim blnDifferent As Boolean = False

        Select Case objMaskField1.Field.DataType
            Case SqlDbType.BigInt, SqlDbType.Int, SqlDbType.SmallInt, SqlDbType.TinyInt
                If Not (CInt(objMaskField1.Value1.Value) = CInt(objMaskField2.Value1.Value)) Then
                    blnDifferent = True
                End If
            Case SqlDbType.Bit
                If Not (CBool(objMaskField1.Value1.Value) = CBool(objMaskField2.Value1.Value)) Then
                    blnDifferent = True
                End If
            Case SqlDbType.Decimal, SqlDbType.Float, SqlDbType.Real
                If Not (CDbl(objMaskField1.Value1.Value) = CDbl(objMaskField2.Value1.Value)) Then
                    blnDifferent = True
                End If
            Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                If Not (CDate(objMaskField1.Value1.Value) = CDate(objMaskField2.Value1.Value)) Then
                    blnDifferent = True
                End If
            Case SqlDbType.Binary, SqlDbType.Image, SqlDbType.VarBinary
                'Can't update image fields, only create new versions
                Return
            Case Else
                If Not (CStr(objMaskField1.Value1.Value) = CStr(objMaskField2.Value1.Value)) Then
                    blnDifferent = True
                End If
        End Select

        If blnDifferent Then
            SerializeValues(objXW, objMaskField1, objMaskField2)
        End If
    End Sub

    Private Sub SerializeValues(ByVal objXW As XmlTextWriter, ByVal objMaskField1 As clsMaskField, ByVal objMaskField2 As clsMaskField)
        objXW.WriteStartElement("Field")
        objXW.WriteAttributeString("caption", objMaskField1.Caption)
        objXW.WriteAttributeString("dbname", objMaskField1.Field.DatabaseName)
        objXW.WriteStartElement("OldValue")

        SerializeValue(objXW, objMaskField1)

        objXW.WriteEndElement()
        objXW.WriteStartElement("NewValue")

        SerializeValue(objXW, objMaskField2)

        objXW.WriteEndElement()
        objXW.WriteEndElement()
    End Sub

    Private Sub CompareAndCreateNode(ByVal objXW As XmlTextWriter, _
    ByVal objMaskFieldLink1 As clsMaskFieldLink, ByVal objMaskFieldLink2 As clsMaskFieldLink)
        Dim colDeletedIDs As New ArrayList
        Dim colInsertedIDs As New ArrayList

        For Each intID As Integer In objMaskFieldLink1.IDCollection.Values
            If objMaskFieldLink2.IDCollection(CStr(intID)) Is Nothing Then
                colDeletedIDs.Add(intID)
            End If
        Next

        For Each intID As Integer In objMaskFieldLink2.IDCollection.Values
            If objMaskFieldLink1.IDCollection(CStr(intID)) Is Nothing Then
                colInsertedIDs.Add(intID)
            End If
        Next

        If colDeletedIDs.Count > 0 OrElse colInsertedIDs.Count > 0 Then
            objXW.WriteStartElement("FieldLink")
            objXW.WriteAttributeString("caption", objMaskFieldLink1.Caption)
            objXW.WriteAttributeString("id", objMaskFieldLink1.FieldLink.ForeignKeyTable.DatabaseName & "." & objMaskFieldLink1.FieldLink.ForeignKeyField.DatabaseName)

            If colDeletedIDs.Count > 0 Then
                Dim objDT As DataTable = objMaskFieldLink1.Database.GetDataTableBySQL( _
                    "SELECT [" & clsDBConstants.Fields.cID & "]," & _
                    "[" & clsDBConstants.Fields.cEXTERNALID & "] " & _
                    "FROM [" & objMaskFieldLink1.FieldLink.LinkedTable.DatabaseName & "] " & _
                    "WHERE [" & clsDBConstants.Fields.cID & "] IN (" & _
                    CreateIDStringFromCollection(colDeletedIDs) & ")")

                objDT.DefaultView.Sort = clsDBConstants.Fields.cEXTERNALID

                For intIndex As Integer = 0 To objDT.DefaultView.Count - 1
                    objXW.WriteStartElement("DeletedValue")
                    objXW.WriteAttributeString("id", CType(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cID), String))
                    objXW.WriteValue(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cEXTERNALID))
                    objXW.WriteEndElement()
                Next

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If
            End If

            If colInsertedIDs.Count > 0 Then
                Dim objDT As DataTable = objMaskFieldLink1.Database.GetDataTableBySQL( _
                    "SELECT [" & clsDBConstants.Fields.cID & "]," & _
                    "[" & clsDBConstants.Fields.cEXTERNALID & "] " & _
                    "FROM [" & objMaskFieldLink1.FieldLink.LinkedTable.DatabaseName & "] " & _
                    "WHERE [" & clsDBConstants.Fields.cID & "] IN (" & _
                    CreateIDStringFromCollection(colInsertedIDs) & ")")

                objDT.DefaultView.Sort = clsDBConstants.Fields.cEXTERNALID

                For intIndex As Integer = 0 To objDT.DefaultView.Count - 1
                    objXW.WriteStartElement("InsertedValue")
                    objXW.WriteAttributeString("id", CType(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cID), String))
                    objXW.WriteValue(objDT.DefaultView(intIndex)(clsDBConstants.Fields.cEXTERNALID))
                    objXW.WriteEndElement()
                Next

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If
            End If

            objXW.WriteEndElement()
        End If
    End Sub

End Class
