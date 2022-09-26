Namespace FrameworkCollections

    ''' <summary>
    ''' A collection of associated System.String keys and System.Object values 
    ''' that can be accessed either with the key or with the index.
    ''' </summary>
    Public Class K1NameCollection
        Inherits System.Collections.Specialized.NameObjectCollectionBase

#Region " Properties "

#Region " All Keys "

        ''' <summary>
        ''' Returns a System.String array that contains all the keys in the collection.
        ''' </summary>
        Public ReadOnly Property AllKeys() As String()
            Get
                Return MyBase.BaseGetAllKeys
            End Get
        End Property

#End Region

#Region " All Values "

        ''' <summary>
        ''' Returns an System.Object array that contains all the values in the collection.
        ''' </summary>
        Public ReadOnly Property AllValues() As Object()
            Get
                Return MyBase.BaseGetAllValues
            End Get
        End Property

        ''' <summary>
        ''' Returns an array of the specified type that contains all the values in the collection.
        ''' </summary>
        ''' <param name="type">A System.Type that represents the type of array to return.</param>
        ''' <exception cref="System.ArgumentNullException">type is null.</exception>
        ''' <exception cref="System.ArgumentException">type is not a valid System.Type.</exception>
        Public ReadOnly Property AllValues(ByVal type As System.Type) As Object()
            Get
                Return MyBase.BaseGetAllValues(type)
            End Get
        End Property

#End Region

#Region " Item "

        ''' <summary>
        ''' Gets a single object by index.
        ''' </summary>
        ''' <param name="index">The numerical index of the object in the collection.</param>
        ''' <returns>The object referenced by index.</returns>
        Default Public ReadOnly Property Item(ByVal index As Integer) As Object
            Get
                Return Me.Get(index)
            End Get
        End Property

        ''' <summary>
        ''' Gets the value of a single object by name.
        ''' </summary>
        ''' <param name="name">The name of the object in the collection.</param>
        ''' <returns>The object referenced by name.</returns>
        Default Public Property Item(ByVal name As String) As Object
            Get
                Return Me.Get(name)
            End Get
            Set(ByVal value As Object)
                Me.Set(name, value)
            End Set
        End Property

#End Region

#End Region

#Region " Methods "

#Region " Add "

        ''' <summary>
        ''' Adds an entry with the specified key and value.
        ''' </summary>
        ''' <param name="name">The System.String key of the entry to add.</param>
        ''' <param name="value">The System.Object value of the entry to add. The value can be null.</param>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Sub Add(ByVal name As String, ByVal value As Object)
            If name Is Nothing Then
                Throw New ArgumentException("name")
            End If

            MyBase.BaseAdd(name, value)
        End Sub

#End Region

#Region " Clear "

        ''' <summary>
        ''' Removes all entries from the collection.
        ''' </summary>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        Public Sub Clear()
            MyBase.BaseClear()
        End Sub

#End Region

#Region " Get "

#Region " Normal "

        ''' <summary>
        ''' Gets the value of the entry at the specified index of the collection.
        ''' </summary>
        ''' <param name="index">The zero-based index of the value to get.</param>
        ''' <returns>An System.Object that represents the value of the entry at the specified index.</returns>
        ''' <exception cref="System.ArgumentOutOfRangeException">index is outside the valid range of indexes for the collection.</exception>
        Public Function [Get](ByVal index As Integer) As Object
            Return MyBase.BaseGet(index)
        End Function

        ''' <summary>
        ''' Gets the value of the first entry with the specified key from the collection.
        ''' </summary>
        ''' <param name="name">The System.String key of the entry to get.</param>
        ''' <returns>An System.Object that represents the value of the first entry with the specified key, if found; otherwise, null.</returns>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Function [Get](ByVal name As String) As Object
            If name Is Nothing Then
                Throw New ArgumentException("name")
            End If

            Return MyBase.BaseGet(name)
        End Function

#End Region

#Region " Generic "

        ''' <summary>
        ''' Gets the value of the entry at the specified index of the collection.
        ''' </summary>
        ''' <param name="index">The zero-based index of the value to get.</param>
        ''' <returns>A TValue that represents the value of the entry at the specified index.</returns>
        ''' <exception cref="System.ArgumentOutOfRangeException">index is outside the valid range of indexes for the collection.</exception>
        Public Function [Get](Of TValue)(ByVal index As Integer) As TValue
            Try
                Return CType(MyBase.BaseGet(index), TValue)
            Catch ex As InvalidCastException
                Return Nothing
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Gets the value of the first entry with the specified key from the collection.
        ''' </summary>
        ''' <param name="name">The System.String key of the entry to get.</param>
        ''' <returns>A TValue that represents the value of the first entry with the specified key, if found; otherwise, null.</returns>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Function [Get](Of TValue)(ByVal name As String) As TValue
            Try
                If name Is Nothing Then
                    Throw New ArgumentException("name")
                End If

                Return CType(MyBase.BaseGet(name), TValue)
            Catch ex As InvalidCastException
                Return Nothing
            Catch ex As Exception
                Throw
            End Try
        End Function

#End Region

#End Region

#Region " Get Key "

        ''' <summary>
        ''' Gets the key of the entry at the specified index of the collection.
        ''' </summary>
        ''' <param name="index">The zero-based index of the key to get.</param>
        ''' <returns>A System.String that represents the key of the entry at the specified index.</returns>
        ''' <exception cref="System.ArgumentOutOfRangeException">index is outside the valid range of indexes for the collection.</exception>
        Public Function GetKey(ByVal index As Integer) As String
            Return MyBase.BaseGetKey(index)
        End Function

#End Region

#Region " Contains Key "

        ''' <summary>
        ''' Gets a value indicating whether the collection contains a specified key.
        ''' </summary>
        ''' <returns>true if the collection contains the keys; otherwise, false.</returns>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Function ContainsKey(ByVal name As String) As Boolean
            Return (Me.Get(name) Is Nothing)
        End Function

#End Region

#Region " Remove "

        ''' <summary>
        ''' Removes the entries with the specified key from the collection.
        ''' </summary>
        ''' <param name="name">The System.String key of the entries to remove. The key can be null.</param>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Sub Remove(ByVal name As String)
            If name Is Nothing Then
                Throw New ArgumentException("name")
            End If

            MyBase.BaseRemove(name)
        End Sub

#End Region

#Region " Remove At "

        ''' <summary>
        ''' Removes the entry at the specified index of the collection.
        ''' </summary>
        ''' <param name="index">The zero-based index of the entry to remove.</param>
        ''' <exception cref="System.ArgumentOutOfRangeException">index is outside the valid range of indexes for the collection.</exception>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        Public Sub RemoveAt(ByVal index As Integer)
            MyBase.BaseRemoveAt(index)
        End Sub

#End Region

#Region " Set "

        ''' <summary>
        ''' Sets the value of the entry at the specified index of the collection.
        ''' </summary>
        ''' <param name="index">The zero-based index of the entry to set.</param>
        ''' <param name="value">The System.Object that represents the new value of the entry to set. The value can be null.</param>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        ''' <exception cref="System.ArgumentOutOfRangeException">index is outside the valid range of indexes for the collection.</exception>
        Public Sub [Set](ByVal index As Integer, ByVal value As Object)
            MyBase.BaseSet(index, value)
        End Sub

        ''' <summary>
        ''' Sets the value of the first entry with the specified key in the collection.
        ''' </summary>
        ''' <param name="name">The System.String key of the entry to set.</param>
        ''' <param name="value">The System.Object that represents the new value of the entry to set. The value can be null.</param>
        ''' <exception cref="System.NotSupportedException">The collection is read-only.</exception>
        ''' <exception cref="System.ArgumentNullException">name is null.</exception>
        Public Sub [Set](ByVal name As String, ByVal value As Object)
            If name Is Nothing Then
                Throw New ArgumentException("name")
            End If

            MyBase.BaseSet(name, value)
        End Sub

#End Region

#End Region

    End Class

End Namespace

