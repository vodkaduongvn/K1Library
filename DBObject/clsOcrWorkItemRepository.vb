Imports System.Reflection
Imports System.Globalization
Imports K1Library.FrameworkCollections

Namespace DBObject

    Public Class clsOcrWorkItemRepository

        Private m_objDbContext As clsDB
        Private m_objTable As clsTable
        Private m_TempFileLocation As String
        Private m_colRecordLocks As K1Dictionary(Of clsRecordLock)

        Sub New(objDbContext As clsDB, tempFileLocation As String)

            m_objDbContext = objDbContext

            Dim tbl = m_objDbContext.SysInfo.Tables.Where(Function(t) t.Value.DatabaseName = "OcrServiceWorkItems").Select(Function(t) t.Value).FirstOrDefault()

            If (tbl Is Nothing) Then
                Throw New clsK1Exception("Could not find table OcrServiceworkItems. Please ensure that you are using the correct version of Recfind 6.")
            Else
                m_objTable = tbl
            End If

            m_TempFileLocation = tempFileLocation

            'If (m_objDbContext.Profile IsNot Nothing) Then
            '    m_TempFileLocation = Path.Combine(Path.GetTempPath(), String.Format("\Knowledgeone Corporation\{0}\OCR Work Item Repository\TempFiles\", m_objDbContext.Profile.ID))
            'Else
            '    m_TempFileLocation = Path.Combine(Path.GetTempPath(), String.Format("\Knowledgeone Corporation\{0}\OCR Work Item Repository\TempFiles\", Guid.NewGuid()))
            'End If

            m_colRecordLocks = New K1Dictionary(Of clsRecordLock)()

        End Sub

        Public Function GetOcrWorkItems(ByVal sortExp As System.Linq.Expressions.Expression(Of Func(Of clsOcrWorkItem, Object)),
                                        Optional sortOrder As String = "ASC", Optional take As Integer = 5) As IEnumerable(Of clsOcrWorkItem)

            Dim orderByField = clsObjectMapper.GetMemeberName(Of clsOcrWorkItem)(sortExp)

            If (Not {"ASC", "DESC"}.Any(Function(so) so = sortOrder.ToUpper())) Then
                Throw New ArgumentException(String.Format("sortOrder with value ({0}) is not allowed.", sortOrder), "sortOrder")
            End If

            Dim strQuery As String = "SELECT TOP " & take.ToString() & " " &
                clsObjectMapper.GetObjectPropertiesToCsv(Of clsOcrWorkItem)() &
                " FROM " & m_objTable.DatabaseName &
                " WHERE IsDone = @IsDone" &
                " ORDER BY " & orderByField & " " & sortOrder
            Dim params = New clsDBParameterDictionary()
            params.Add(New clsDBParameter("@IsDone", 0))

            Dim objDt As DataTable = m_objDbContext.GetDataTableBySQL(strQuery, params)
            Dim domainObjects As List(Of clsOcrWorkItem) = New List(Of clsOcrWorkItem)(objDt.Rows.Count)
            domainObjects.AddRange(From dr As DataRow In objDt.Rows.OfType(Of DataRow)() Select clsObjectMapper.GetObject(Of clsOcrWorkItem)(dr, m_objTable))

            Return domainObjects

        End Function

        Public Function GetPendingOcrWorkItemsByEdocIds(colEdocIds As IEnumerable(Of Integer),
                                                        Optional queryUtil As clsSelectInfo = Nothing) As IEnumerable(Of clsOcrWorkItem)

            If (colEdocIds.Count() = 0) Then
                Return New List(Of clsOcrWorkItem)()
            End If

            Dim objSearchElement = New clsSearchElement(clsSearchFilter.enumOperatorType.NONE, "EdocId",
                                                        clsSearchFilter.enumComparisonType.IN,
                                                        colEdocIds.ToDelimitedString(",".ToCharArray()(0)))

            Dim objEdocIdField = m_objTable.GetField("EdocId")
            If (queryUtil Is Nothing) Then
                queryUtil = New clsSelectInfo(m_objTable, Nothing, False)
            End If

            Dim strFilter As String = queryUtil.GetFilter(objSearchElement, objEdocIdField, objEdocIdField.DatabaseName, False)

            Dim strQuery As String = "SELECT " &
                " " & clsObjectMapper.GetMemeberName(Of clsOcrWorkItem)(Function(o) o.Id) &
                ", " & clsObjectMapper.GetMemeberName(Of clsOcrWorkItem)(Function(o) o.EdocId) &
                ", " & clsObjectMapper.GetMemeberName(Of clsOcrWorkItem)(Function(o) o.ProcessedDate) &
                " FROM " & m_objTable.DatabaseName &
                " WHERE IsDone = @IsDone" &
                " AND (" & strFilter & ")"

            Dim params = New clsDBParameterDictionary()
            params.Add(New clsDBParameter("@IsDone", 0))

            Dim dt = m_objDbContext.GetDataTableBySQL(strQuery, params)

            Dim resultSet = New List(Of clsOcrWorkItem)(dt.Rows.Count)
            If (resultSet.Capacity > 0) Then
                resultSet.AddRange(From dr As DataRow In dt.Rows.OfType(Of DataRow)()
                                   Select clsObjectMapper.GetObject(Of clsOcrWorkItem)(dr, m_objTable))
            End If

            Return resultSet

        End Function

        Public Function GetOcrWorkItem(id As Integer, Optional edocId As Nullable(Of Integer) = Nothing) As clsOcrWorkItem
            Dim strQuery As String = "SELECT Top 1 " & clsObjectMapper.GetObjectPropertiesToCsv(Of clsOcrWorkItem)() & " FROM " & m_objTable.DatabaseName & " WHERE ISDONE = @IsDone And ID = @Id AND EDOCID = @EdocId"
            Dim params = New clsDBParameterDictionary()
            params.Add(New clsDBParameter("@IsDone", 0))
            params.Add(New clsDBParameter("@Id", id))
            If (edocId.HasValue) Then
                params.Add(New clsDBParameter("@EdocId", edocId.Value))
            End If
            Dim objDt As DataTable = m_objDbContext.GetDataTableBySQL(strQuery, params)
            If (objDt.Rows.Count > 0) Then
                Return Nothing
            End If
            Dim obj = clsObjectMapper.GetObject(Of clsOcrWorkItem)(objDt.Rows(0), m_objTable)
            Return obj
        End Function

        Public Function UpdateWorkItem(ByVal objOcrWorkItem As clsOcrWorkItem, Optional fullQualifiedFileName As String = "") As Boolean
            'Note this bypasses authorization security checks!
            Dim colMaskCollection = clsMaskField.CreateMaskCollection(m_objTable, objOcrWorkItem.Id)
            colMaskCollection.UpdateMaskObj("IsDone", objOcrWorkItem.IsDone)
            colMaskCollection.UpdateMaskObj("ProcessedDate", objOcrWorkItem.ProcessedDate)

            If (Not String.IsNullOrEmpty(fullQualifiedFileName)) Then
                Dim colParams As New clsDBParameterDictionary
                colParams.Add(New clsDBParameter("@BinaryData", New Byte() {0}))
                colParams.Add(New clsDBParameter("@Id", objOcrWorkItem.Id))
                m_objDbContext.ExecuteSQL("UPDATE " & m_objTable.DatabaseName & " SET " & "[File]" & " = @BinaryData WHERE ID = @Id",
                                          colParams)
                Dim objField = m_objTable.GetField("File")
                m_objDbContext.WriteBLOB(m_objTable.DatabaseName, objField.DatabaseName, objField.DataType, objField.Length, objOcrWorkItem.Id, fullQualifiedFileName, False)
            End If

            colMaskCollection.Update(m_objDbContext, False)
            Return True
        End Function

        Public Sub InsertWorkItem(ByVal objWorkItem As clsOcrWorkItem)
            'create a new empty MaskFieldDictionary
            Dim record As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(m_objTable, clsDBConstants.cintNULL)
            'set the values of the MaskFieldDictionary
            For Each propInfo As PropertyInfo In objWorkItem.GetType.GetProperties().Where(Function(p) p.Name.ToUpper() <> clsDBConstants.Fields.cID.ToUpper())
                Dim value As Object
                Try
                    value = propInfo.GetValue(objWorkItem, BindingFlags.Public, Nothing, Nothing, CultureInfo.CurrentCulture)
                Catch ex As Exception
                    value = Nothing
                End Try
                record.UpdateMaskObj(propInfo.Name, value)
            Next
            'system columns which are used for this table are defaulted to null values
            record.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_objDbContext.SysInfo.K1Configuration.DRMDefaultSecurityID)
            record.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, clsDBConstants.cintNULL)
            record.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, objWorkItem.EdocId)
            'use the MaskFieldDictionary class to persist the object
            record.Insert(m_objDbContext)
        End Sub

        Public Function GetWorkItemFile(owi As clsOcrWorkItem) As String
            Try
                Dim tempFullQualifiedFileName As String = Path.Combine(m_TempFileLocation, Guid.NewGuid().ToString() & ".pdf")
                SaveBlobToFile(owi.Id, tempFullQualifiedFileName)
                Return tempFullQualifiedFileName
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        Private Sub SaveBlobToFile(id As Int32, ByVal strFile As String)
            m_objDbContext.ReadBLOB(m_objTable.DatabaseName, "File", id, strFile, False)
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

        Sub DeleteProcessedWorkItems()

            Dim strQuery As String = "DELETE " & m_objTable.DatabaseName & " WHERE ISDONE = 1 AND PROCESSEDDATE IS NOT NULL"
            m_objDbContext.ExecuteSQL(strQuery)

        End Sub

    End Class

End Namespace