Public Class clsMaskFieldLinkDictionary
    Inherits K1Library.FrameworkCollections.K1Dictionary(Of clsMaskFieldLink)

#Region " Properties "

    'Default Public Property Item(ByVal key As [String]) As clsMaskFieldLink
    '    Get
    '        Return CType(Dictionary(key), clsMaskFieldLink)
    '    End Get
    '    Set(ByVal Value As clsMaskFieldLink)
    '        Dictionary(key) = Value
    '    End Set
    'End Property
#End Region

#Region " Methods "

#Region " Overrides "

    'Protected Overrides Sub OnInsert(ByVal key As [Object], ByVal value As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    '    If Not TypeOf value Is clsMaskFieldLink Then
    '        Throw New ArgumentException("value must be of type clsMaskFieldLink.", "value")
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
    '    If Not TypeOf newValue Is clsMaskFieldLink Then
    '        Throw New ArgumentException("newValue must be of type clsMaskFieldLink.", "value")
    '    End If
    'End Sub 'OnSet

    'Protected Overrides Sub OnValidate(ByVal key As [Object], ByVal value As [Object])
    '    If Not TypeOf key Is String Then
    '        Throw New ArgumentException("key must be of type String.", "key")
    '    End If
    '    If Not TypeOf value Is clsMaskFieldLink Then
    '        Throw New ArgumentException("value must be of type clsMaskFieldLink.", "value")
    '    End If
    'End Sub 'OnValidate
#End Region

#Region " Public "

    Public Overloads Sub Add(ByVal value As clsMaskFieldLink)
        Dim strKey As String

        strKey = value.FieldLink.KeyID
        Me.Add(strKey, value)
    End Sub 'Add

    Public Sub InsertUpdate(ByVal objDB As clsDB, ByVal intRecordID As Integer)
        For Each objMaskFieldLink As clsMaskFieldLink In Me.Values
            If objMaskFieldLink.Updated Then
                Dim arrIDs As New ArrayList
                For Each intID As Integer In objMaskFieldLink.IDCollection.Values
                    arrIDs.Add(intID)
                Next

                For Each intID As Integer In arrIDs
                    If Not objMaskFieldLink.NewIDCollection(CStr(intID)) Is Nothing Then
                        objMaskFieldLink.IDCollection.Remove(CStr(intID))
                        objMaskFieldLink.NewIDCollection.Remove(CStr(intID))
                    End If
                Next

                If objMaskFieldLink.IDCollection.Count > 0 Then
                    objDB.ExecuteSQL("CREATE TABLE #tmpTable (ID INT)")

                    objDB.BulkInsert(ConvertIDsToDataTable(objMaskFieldLink.IDCollection), "#tmpTable")

                    Try
                        objDB.ExecuteSQL("DELETE FROM [" & objMaskFieldLink.FieldLink.ForeignKeyTable.DatabaseName & "] " & _
                            "WHERE [" & objMaskFieldLink.FieldLink.LinkTableOppositeFieldLink.ForeignKeyField.DatabaseName & "] " & _
                            "IN (SELECT ID FROM #tmpTable) " & _
                            "AND [" & objMaskFieldLink.FieldLink.ForeignKeyField.DatabaseName & "] = " & intRecordID)
                    Catch ex As clsK1Exception
                        If Not ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE Then
                            Throw
                        End If
                    End Try

                    objDB.ExecuteSQL("DROP TABLE #tmpTable")
                End If

                If objMaskFieldLink.NewIDCollection.Count > 0 Then
                    objDB.ExecuteSQL("CREATE TABLE #tmpTable (ID INT)")

                    objDB.BulkInsert(ConvertIDsToDataTable(objMaskFieldLink.NewIDCollection), "#tmpTable")

                    Try
                        objDB.ExecuteSQL("INSERT INTO [" & objMaskFieldLink.FieldLink.ForeignKeyTable.DatabaseName & "] " & _
                            "([" & objMaskFieldLink.FieldLink.ForeignKeyField.DatabaseName & "],[" & _
                            objMaskFieldLink.FieldLink.LinkTableOppositeFieldLink.ForeignKeyField.DatabaseName & "]) " & _
                            "SELECT " & intRecordID & ", [ID] FROM #tmpTable")
                    Catch ex As clsK1Exception
                        If Not ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.NO_SUCH_TABLE Then
                            Throw
                        End If
                    End Try

                    objDB.ExecuteSQL("DROP TABLE #tmpTable")
                End If

                objMaskFieldLink.ValuesLoaded = False
                objMaskFieldLink.Updated = False
            End If
        Next
    End Sub
#End Region

#End Region

End Class
