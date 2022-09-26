Namespace FilterClasses

    Public Class K1QueryFilterFactory

        Private m_objDb As clsDB

        Public Sub New(db As clsDB)
            m_objDb = db
        End Sub

        Public Function CreateFilter(Of T)(comparisonType As clsSearchFilter.enumComparisonType,
                                           strFieldName As String,
                                           strTableName As String,
                                           value As T) As clsSearchFilter

            Return New clsSearchFilter(m_objDb, String.Format("{0}.{1}", strTableName, strFieldName), comparisonType, value)

        End Function

        Public Function CreateFilter(strFilterString As String, strTableName As String) As clsSearchFilter

            Return New clsSearchFilter(m_objDb, strFilterString, strTableName)

        End Function

        Public Function CreateFilter(ByVal strTableName As String, ByVal searchCriteria As clsSearchGroup) As clsSearchFilter

            Return New clsSearchFilter(m_objDb, searchCriteria, strTableName)

        End Function

        Public Function CreateFilter(strTableName As String, ParamArray searchCriterias() As clsSearchGroup) As clsSearchFilter

            Dim sg = New clsSearchGroup(clsSearchFilter.enumOperatorType.NONE, searchCriterias.Cast(Of clsSearchObjBase)().ToList())

            Return New clsSearchFilter(m_objDb, sg, strTableName)

        End Function

        ''' <summary>
        ''' Create a single search condition i.e table.Column = var etc...
        ''' </summary>
        ''' <param name="comparitor"></param>
        ''' <param name="tableName"></param>
        ''' <param name="condition"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateSearchCondition(Of T)(comparitor As clsSearchFilter.enumComparisonType,
                                                    tableName As String,
                                                    condition As KeyValuePair(Of String, T)) As clsSearchElement

            Return New clsSearchElement(clsSearchFilter.enumOperatorType.NONE,
                                        String.Format("{0}.{1}", tableName, condition.Key),
                                        comparitor,
                                        condition.Value)
        End Function

        Public Function CreateSearchCriteriaFromConditions([operator] As clsSearchFilter.enumOperatorType,
                                                           tableName As String,
                                                           ParamArray conditions() As clsSearchElement) As clsSearchGroup

            Dim firstCondition = conditions.FirstOrDefault()
            If (firstCondition IsNot Nothing) Then
                firstCondition.OperatorType = clsSearchFilter.enumOperatorType.NONE
            End If

            For Each condition As clsSearchElement In conditions.Skip(1)
                condition.OperatorType = [operator]
            Next

            Return New clsSearchGroup([operator], conditions.Cast(Of clsSearchObjBase)().ToList())

        End Function

        Public Function CreateSearchCriteriaFromCriteria([operator] As clsSearchFilter.enumOperatorType,
                                                         tableName As String,
                                                         ParamArray criterias() As clsSearchGroup) As clsSearchGroup

            Dim firstCondition = criterias.FirstOrDefault()
            If (firstCondition IsNot Nothing) Then
                firstCondition.OperatorType = clsSearchFilter.enumOperatorType.NONE
            End If

            For Each condition As clsSearchGroup In criterias.Skip(1)
                condition.OperatorType = [operator]
            Next

            Return New clsSearchGroup(clsSearchFilter.enumOperatorType.NONE, criterias.Select(Function(c) c).Cast(Of clsSearchObjBase)().ToList())

        End Function

    End Class
End Namespace