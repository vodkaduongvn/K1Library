Public Class clsSortCollection
    Inherits System.Collections.Generic.List(Of clsSort)
    Implements IDisposable

#Region " Member Variables "

    Private m_colSorts As New Hashtable
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls

#End Region

#Region " Properties "

    'Default Property Item(ByVal index As Integer) As clsSort
    '    Get
    '        Return CType(MyBase.List.Item(index), clsSort)
    '    End Get
    '    Set(ByVal Value As clsSort)
    '        MyBase.List.Item(index) = Value
    '    End Set
    'End Property

    Default Overloads ReadOnly Property Item(ByVal strKey As String) As clsSort
        Get
            Return CType(m_colSorts(strKey), clsSort)
        End Get
    End Property

#End Region

#Region " Methods "

#Region " Overrides "

    'Protected Overrides Sub OnInsert(ByVal index As Integer, ByVal value As Object)
    '    If Not TypeOf (value) Is clsSort Then
    '        Throw New ArgumentException("Invalid type.")
    '    End If
    'End Sub

    'Protected Overrides Sub OnSet(ByVal index As Integer, _
    'ByVal oldValue As Object, ByVal newValue As Object)
    '    If Not TypeOf (newValue) Is clsSort Then
    '        Throw New ArgumentException("Invalid type.")
    '    End If
    'End Sub

    'Protected Overrides Sub OnValidate(ByVal value As Object)
    '    If Not TypeOf (value) Is clsSort Then
    '        Throw New ArgumentException("Invalid type.")
    '    End If
    'End Sub
#End Region

#Region " Public "

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                For Each objItem As clsSort In Me
                    If objItem IsNot Nothing Then
                        objItem.Dispose()
                        objItem = Nothing
                    End If
                Next

                If m_colSorts IsNot Nothing Then
                    m_colSorts.Clear()
                    m_colSorts = Nothing
                End If

                Me.Clear()
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Public Overloads Function Contains(ByVal value As clsSort) As Boolean
        For Each objSort As clsSort In Me
            If value.Field.ID = objSort.Field.ID Then
                Return True
            End If
        Next

        Return False
    End Function

    '2017-09-11 -- Peter Melisi -- Bug fix for #1700003358
    Public Overloads Function Contains(ByVal value As String) As Boolean
        For Each objSort As clsSort In Me
            If value = objSort.Field.DatabaseName Then
                Return True
            End If
        Next

        Return False
    End Function

    Public Overloads Function Add(ByVal value As clsSort) As Integer
        m_colSorts.Add(value.Field.DatabaseName, value)
        MyBase.Add(value)
        value.SortOrder = Me.Count
        Dim intIndex As Integer = Me.Count - 1
        Return intIndex
    End Function

    Public Overloads Sub Insert(ByVal index As Integer, ByVal value As clsSort)
        value.SortOrder = index + 1
        For Each objSort As clsSort In Me
            If objSort.SortOrder >= value.SortOrder Then
                value.SortOrder += 1
            End If
        Next
        m_colSorts.Add(value.Field.DatabaseName, value)
        MyBase.Insert(index, value)
    End Sub

    'Public Function IndexOf(ByVal value As clsSort) As Integer
    '    Return MyBase.IndexOf(value)
    'End Function

    'Public Function Contains(ByVal value As clsSort) As Boolean
    '    Return (m_colSorts(value.Field.DatabaseName) IsNot Nothing)
    'End Function

    Public Overloads Sub Remove(ByVal value As clsSort)
        For Each objSort As clsSort In Me
            If objSort.SortOrder > value.SortOrder Then
                value.SortOrder -= 1
            End If
        Next
        m_colSorts.Remove(value.Field.DatabaseName)
        MyBase.Remove(value)
    End Sub

    Public Overloads Sub RemoveAt(ByVal index As Integer)
        Dim value As clsSort = Me(index)
        For Each objSort As clsSort In Me
            If objSort.SortOrder > value.SortOrder Then
                value.SortOrder -= 1
            End If
        Next
        m_colSorts.Remove(value.Field.DatabaseName)
        MyBase.RemoveAt(index)
    End Sub

    Public Function GetSortString() As String
        Dim strSort As String = ""
        Dim blnFoundIDField As Boolean = False

        For intLoop As Integer = 0 To Me.Count - 1
            Dim objSort As clsSort = Me(intLoop)

            AppendToCommaString(strSort, objSort.Field.DatabaseName)
            
            If objSort.Field.IsIdentityField Then
                blnFoundIDField = True
            End If

            If Not objSort.Ascending Then
                strSort &= " DESC"
            End If
        Next

        'the sort string must include the id
        If Not blnFoundIDField Then
            AppendToCommaString(strSort, clsDBConstants.Fields.cID)
        End If

        Return strSort
    End Function

#End Region

#End Region

End Class
