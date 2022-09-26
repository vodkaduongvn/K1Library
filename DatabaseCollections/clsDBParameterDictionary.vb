Imports System.Xml.Serialization

<Serializable()> _
Public Class clsDBParameterDictionary
    Inherits FrameworkCollections.K1Dictionary(Of clsDBParameter)

#Region " Public "

    Public Overloads Sub Add(ByVal value As clsDBParameter)
        Me.Add(value.Name, value)
    End Sub 'Add

    Public Function Serialize() As Byte()()
        Try
            Dim arrParams As Byte()() = Nothing
            Dim intParam As Integer

            If Me.Count > 0 Then
                arrParams = CType(Array.CreateInstance(GetType(Byte()), Me.Count), Byte()())

                For Each objParam As clsDBParameter In Me.Values
                    arrParams(intParam) = objParam.Serialize()
                    intParam += 1
                Next
            End If

            Return arrParams
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function Deserialize(ByVal arrParams As Byte()()) As clsDBParameterDictionary
        Try
            Dim colParameterCollection As clsDBParameterDictionary = Nothing

            If Not arrParams Is Nothing AndAlso arrParams.Length > 0 Then
                colParameterCollection = New clsDBParameterDictionary

                For Each arrParam As Byte() In arrParams
                    colParameterCollection.Add(clsDBParameter.Deserialize(arrParam))
                Next
            End If

            Return colParameterCollection
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class
