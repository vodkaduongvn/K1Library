Public MustInherit Class clsDRMBase
    Implements IDisposable

#Region " Members "

    Protected m_objDB As clsDB_System
    Protected m_intID As Integer = clsDBConstants.cintNULL
    Protected m_strExternalID As String = clsDBConstants.cstrNULL
    Protected m_intSecurityID As Integer = clsDBConstants.cintNULL
    Protected m_intTypeID As Integer = clsDBConstants.cintNULL
    Protected m_objDBObj As clsDBObjBase
    Protected m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " constructors "

    Protected Sub New(ByVal objDB As clsDB, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer)
        m_objDB = CType(objDB, clsDB_System)
        m_strExternalID = strExternalID
        m_intSecurityID = intSecurityID
        m_intTypeID = intTypeID
    End Sub

    Protected Sub New(ByVal objDB As clsDB, _
    ByVal objDatabaseObj As clsDBObjBase)
        m_objDB = CType(objDB, clsDB_System)
        m_objDBObj = objDatabaseObj
        m_intID = objDatabaseObj.ID
        m_strExternalID = objDatabaseObj.ExternalID
        m_intSecurityID = objDatabaseObj.SecurityID
        m_intTypeID = objDatabaseObj.TypeID
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property ID() As Integer
        Get
            Return m_intID
        End Get
    End Property

    Public Property ExternalID() As String
        Get
            Return m_strExternalID
        End Get
        Set(ByVal value As String)
            m_strExternalID = value
        End Set
    End Property

    Public Property SecurityID() As Integer
        Get
            Return m_intSecurityID
        End Get
        Set(ByVal value As Integer)
            m_intSecurityID = value
        End Set
    End Property

    Public Property TypeID() As Integer
        Get
            Return m_intTypeID
        End Get
        Set(ByVal value As Integer)
            m_intTypeID = value
        End Set
    End Property

    Public ReadOnly Property DatabaseObject() As clsDBObjBase
        Get
            Return m_objDBObj
        End Get
    End Property

    Public ReadOnly Property SystemDB() As clsDB_System
        Get
            Return m_objDB
        End Get
    End Property
#End Region

#Region " Methods "

    Public Shared Function ValidNameCharacters(ByVal strName As String) As Boolean
        Dim arrChars() As Char = strName.ToCharArray
        For Each strChar As Char In arrChars
            If Not Char.IsLetter(strChar) AndAlso Not strChar = "_"c AndAlso Not Char.IsDigit(strChar) Then
                Return False
            End If
        Next

        Return True
    End Function

    Protected Function CreateCaption(ByVal objCaption As clsCaption, ByVal strText As String, Optional ByVal blnCreateAuditTrailRecord As Boolean = True) As Integer
        Dim objDRMCaption As clsDRMCaption

        objDRMCaption = New clsDRMCaption(m_objDB, objCaption, strText)

        objDRMCaption.InsertUpdate(blnCreateAuditTrailRecord)

        Return objDRMCaption.Caption.ID
    End Function

    Protected Function CreateCaption(ByVal strExternalID As String, ByVal strText As String, Optional ByVal blnCreateAuditTrailRecord As Boolean = True) As Integer
        Dim objDRMCaption As clsDRMCaption

        objDRMCaption = New clsDRMCaption(m_objDB, strExternalID, _
                m_intSecurityID, m_intTypeID, clsDBConstants.cintNULL, strText)

        objDRMCaption.InsertUpdate(blnCreateAuditTrailRecord)

        Return objDRMCaption.Caption.ID
    End Function

    Protected Sub DeleteCaption(ByVal objCaption As clsCaption, Optional ByVal blnCreateAuditTrailRecord As Boolean = True)
        If Not objCaption Is Nothing Then
            Dim objDRMCaption As New clsDRMCaption(m_objDB, objCaption, "")
            objDRMCaption.Delete(blnCreateAuditTrailRecord)
        End If
    End Sub

#Region " Recursive Record Deletion "

    Public Shared Sub RecurseDeleteRelatedRecords(ByVal objDB As clsDB_System,
                                                  ByVal strTable As String,
                                                  ByVal intID As Integer,
                                                  Optional ByVal blnCreateAuditTrailRecord As Boolean = True)

        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not objDB.HasTransaction Then
                objDB.BeginTransaction()
                blnCreatedTransaction = True
            End If

            Dim objDT As New DataTable("DT")

            objDT.Columns.Add(clsDBConstants.Fields.cID, GetType(Integer))

            Dim objRow As DataRow = objDT.NewRow
            objRow(clsDBConstants.Fields.cID) = intID
            objDT.Rows.Add(objRow)

            RecurseDeleteRelatedRecords(objDB, strTable, objDT, Nothing, blnCreateAuditTrailRecord)

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

    Private Shared Sub RecurseDeleteRelatedRecords(ByVal objDB As clsDB_System,
                                                   ByVal strTable As String,
                                                   ByVal objDT As DataTable,
                                                   ByVal objField As clsField,
                                                   Optional ByVal blnCreateAuditTrailRecord As Boolean = True)

        If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
            Return
        End If

        Dim blnDeleteData As Boolean = (objField Is Nothing OrElse Not objField.IsNullable)
        Dim strIDs As String = CreateIDStringFromDataTable(objDT)

        Dim objTable As clsTable = objDB.SysInfo.Tables(strTable)

        If Not blnDeleteData Then
            Try
                objDB.ClearRecordKey(objTable.DatabaseName,
                                     objField.DatabaseName,
                                     clsDBConstants.Fields.cID,
                                     strIDs,
                                     blnCreateAuditTrailRecord)
            Catch ex As Exception
                blnDeleteData = True
            End Try
        End If

        If blnDeleteData Then

            'Clear all the field links
            For Each objFieldLink As clsFieldLink In objTable.ManyToManyLinks.Values

                Dim objLinkDT As DataTable = GetRelatedRecords(objDB,
                                                               objFieldLink,
                                                               strIDs)
                Dim strLinkIDs As String = CreateIDStringFromDataTable(objLinkDT)


                If (blnCreateAuditTrailRecord) Then
                    clsAuditTrail.CreateAuditTrailRecords(objTable.Database,
                                                      clsMethod.enumMethods.cDELETE,
                                                      objFieldLink.ForeignKeyTable,
                                                      objLinkDT)
                End If

                objDB.DeleteRecordRange(objFieldLink.ForeignKeyTable.DatabaseName,
                                        clsDBConstants.Fields.cID,
                                        strLinkIDs)
            Next

            'Go through the foreign tables and either delete if not allow nulls, or clear fkeys
            For Each objFieldLink As clsFieldLink In objTable.OneToManyLinks.Values
                RecurseDeleteRelatedRecords(objDB, objFieldLink.ForeignKeyTable.DatabaseName,
                                            GetRelatedRecords(objDB, objFieldLink, strIDs),
                                            objFieldLink.ForeignKeyField,
                                            blnCreateAuditTrailRecord)
            Next

            objDB.DeleteRecordRange(objTable.DatabaseName,
                                    clsDBConstants.Fields.cID,
                                    strIDs,
                                    blnCreateAuditTrailRecord)

        End If

    End Sub

    Private Shared Function GetRelatedRecords(ByVal objDB As clsDB_System, _
    ByVal objFieldLink As clsFieldLink, ByVal strIDs As String) As DataTable
        Try
            Dim strSQL As String = "SELECT [" & clsDBConstants.Fields.cID & "] " & _
                "FROM [" & objFieldLink.ForeignKeyTable.DatabaseName & "] " & _
                "WHERE [" & objFieldLink.ForeignKeyField.DatabaseName & "] IN (" & strIDs & ")"

            Return objDB.GetDataTableBySQL(strSQL)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#End Region

#Region " IDisposable Support "

    Protected MustOverride Sub DisposeDRMObject()

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
                m_objDBObj = Nothing

                DisposeDRMObject()
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
