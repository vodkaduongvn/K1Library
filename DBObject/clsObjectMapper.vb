Imports System.Reflection
Imports System.Globalization
Imports System.Text
Imports System.Linq.Expressions

Namespace DBObject
    Friend Class clsObjectMapper

        Public Shared Function [GetObject](Of T As Class)(dr As DataRow, objTable As clsTable) As T

            Dim domainObj = Activator.CreateInstance(Of T)()

            For Each pInfo As PropertyInfo In domainObj.GetType().GetProperties()

                Dim pInfoCopy As PropertyInfo = pInfo

                Dim field = objTable.Fields.Values.FirstOrDefault(Function(f) String.Equals(f.DatabaseName, pInfoCopy.Name, StringComparison.CurrentCultureIgnoreCase))

                If (pInfo.CanWrite AndAlso field IsNot Nothing) Then

                    '[Naing] Could be a projection so ignore when column is not found
                    If (dr.Table.Columns.Contains(field.DatabaseName)) Then
                        Dim value = dr(field.DatabaseName)
                        If (IsDBNull(value)) Then
                            value = Nothing
                        End If
                        pInfo.SetValue(domainObj, value, BindingFlags.Default, Nothing, Nothing, CultureInfo.CurrentCulture)
                    End If
                    
                End If

            Next pInfo

            Return domainObj

        End Function

        Public Shared Function GetObjectPropertiesToCsv(Of T As Class)() As String

            Dim fields As StringBuilder = New StringBuilder()
            Const strSeperator As String = ", "
            For Each pInfo As PropertyInfo In GetType(T).GetProperties()
                fields.Append(pInfo.Name)
                fields.Append(strSeperator)
            Next

            Return fields.ToString().TrimEnd(CType(",", Char), CType(" ", Char))

        End Function

        Public Shared Function GetMemeberName(Of T As Class)(ByVal exp As Expression(Of Func(Of T, Object))) As String

            Dim member = TryCast(exp.Body, MemberExpression)
            If (member Is Nothing) Then
                Dim convert = TryCast(exp.Body, UnaryExpression)
                If (convert Is Nothing) Then
                    Return String.Empty
                End If
                member = TryCast(convert.Operand, MemberExpression)
                If (member Is Nothing) Then
                    Return String.Empty
                End If
            End If
            Return member.Member.Name

        End Function

    End Class
End Namespace