Namespace FrameworkCollections

    ''' <summary>
    ''' Represents a generic collection of keys (String) and values (TValue) with index keys (IValue).
    ''' 
    ''' Use this if you have a collection of objects that you regulary use more then one key
    ''' to find it in the collection as it is faster then K1Dictionary and other ICollection classes
    ''' with a near O(1) 
    ''' </summary>    
    ''' <typeparam name="TValue">The datatype of the value you want to store in the Dictionary.</typeparam>
    ''' <typeparam name="IValue">The datatype of the index key you want to use in the Dictionary.</typeparam>
    '''<deprecated></deprecated>
    ''' <remark>This is bloody backwards!!! Should be Dictionary(TKey, TValue) Why is it backwards?!!!</remark>
    ''' <remark>Do not use this class anymore in future.</remark>
    Public Class K1DualKeyDictionary(Of TValue, IValue)
        Inherits System.Collections.Generic.Dictionary(Of String, TValue)
        Implements IDisposable

#Region " Members "

        Private m_colIndexHashTable As New Hashtable
        Private m_strIndexPropertyName As String = "ID"     '-- Default Property to use as the index
        Private m_blnDisposedValue As Boolean = False       '-- To detect redundant calls

#End Region

#Region " Constructor "

        ''' <summary>
        ''' Initializes a new instance of the Dictionary class that is empty, has the default initial capacity, 
        ''' uses StringComparer.Ordinal for comparing keys and uses ID property of the value as the index key.
        ''' </summary>
        Public Sub New()
            MyBase.New(StringComparer.Ordinal)
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the Dictionary class that is empty, has the default initial capacity, 
        ''' uses StringComparer.Ordinal for comparing keys and uses the specified property of the value as 
        ''' the index key.
        ''' </summary>
        ''' <param name="strIndexPropertyName">The property that will be use as the index key for a value</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal strIndexPropertyName As String)
            Me.New()
            m_strIndexPropertyName = strIndexPropertyName
        End Sub

#End Region

#Region " Property "

        ''' <summary>
        ''' Gets or sets the value associated with the specified key.
        ''' </summary>
        ''' <param name="strKey">The key of the value to get or set.</param>
        ''' <returns>
        ''' The value associated with the specified key. If the specified key is not found, 
        ''' a get operation throws a System.Collections.Generic.KeyNotFoundException, 
        ''' and a set operation creates a new element with the specified key.
        ''' </returns>
        ''' <exception cref="System.ArgumentNullException">The key is null.</exception>
        ''' <exception cref="System.Collections.Generic.KeyNotFoundException">
        ''' The property is retrieved and key does not exist in the collection.
        ''' </exception>
        Default Public Overloads Property Item(ByVal strKey As String) As TValue
            Get
                If Not strKey Is Nothing AndAlso Me.ContainsKey(strKey.ToUpper) Then
                    Return MyBase.Item(strKey.ToUpper)
                Else
                    Return Nothing
                End If
            End Get
            Set(ByVal value As TValue)
                strKey = strKey.ToUpper
                MyBase.Item(strKey) = value
                Dim indexKey As IValue = GetIndexKey(value)
                If m_colIndexHashTable.Contains(indexKey) Then
                    m_colIndexHashTable(indexKey) = strKey
                Else
                    m_colIndexHashTable.Add(indexKey, strKey)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the value associated with the specified index key.
        ''' </summary>
        ''' <param name="indexKey">The index key of the value to get or set.</param>
        ''' <returns>
        ''' The value associated with the specified index key. If the specified index key is not found, 
        ''' both get and set operation throw a System.Collections.Generic.KeyNotFoundException thrown, 
        ''' note this is different from Item when the key is specified.
        ''' </returns>
        ''' <exception cref="System.ArgumentNullException">The key is null.</exception>
        ''' <exception cref="System.Collections.Generic.KeyNotFoundException">
        ''' The key does not exist in the collection.
        ''' </exception>
        Default Public Overloads Property Item(ByVal indexKey As IValue) As TValue
            Get
                Return Me(CStr(m_colIndexHashTable(indexKey)))
            End Get
            Set(ByVal value As TValue)
                If m_colIndexHashTable.Contains(indexKey) Then
                    Me(CStr(m_colIndexHashTable(indexKey))) = value
                Else
                    Throw New System.Collections.Generic.KeyNotFoundException("Could not find the key '" & indexKey.ToString & "' in the dictionary.")
                End If
            End Set
        End Property

#End Region

#Region " Methods "

#Region " Add "

        ''' <summary>
        ''' Adds the specified key and value to the dictionary.
        ''' </summary>
        ''' <param name="strKey">The key of the element to add.</param>
        ''' <param name="objValue">The value of the element to add. The value can be null for reference types.</param>
        ''' <exception cref="System.ArgumentException">An element with the same key already exists in the Dictionary.</exception>
        ''' <exception cref="System.ArgumentNullException">The key is null.</exception>
        Public Overloads Sub Add(ByVal strKey As String, ByVal objValue As TValue)
            strKey = strKey.ToUpper
            MyBase.Add(strKey, objValue)
            m_colIndexHashTable.Add(GetIndexKey(objValue), strKey)
        End Sub

        ''' <summary>
        ''' Adds the specified key and value to the dictionary.
        ''' </summary>
        ''' <param name="strKey">The key of the element to add.</param>
        ''' <param name="objValue">The value of the element to add. The value can be null for reference types.</param>
        ''' <exception cref="System.ArgumentException">An element with the same key already exists in the Dictionary.</exception>
        ''' <exception cref="System.ArgumentNullException">The key is null.</exception>
        Public Overloads Sub Add(ByVal strKey As String, ByVal indexKey As IValue, ByVal objValue As TValue)
            strKey = strKey.ToUpper
            MyBase.Add(strKey, objValue)
            m_colIndexHashTable.Add(indexKey, strKey)
        End Sub

#End Region

#Region " Remove "

        ''' <summary>
        ''' Removes the value with the specified key from the Dictionary.
        ''' </summary>
        ''' <param name="strKey">The key of the element to remove.</param>
        ''' <returns>
        ''' true if the element is successfully found and removed; otherwise, false. 
        ''' This method returns false if key is not found in the Dictionary.
        ''' </returns>
        ''' <remarks></remarks>
        Public Overloads Function Remove(ByVal strKey As String) As Boolean
            Dim objValue As TValue = CType(Me(strKey), TValue)
            m_colIndexHashTable.Remove(GetIndexKey(objValue))
            Return MyBase.Remove(strKey.ToUpper)
        End Function

        ''' <summary>
        ''' Removes the value with the specified index key from the Dictionary.
        ''' </summary>
        ''' <param name="indexKey">The index key of the element to remove.</param>
        ''' <returns>
        ''' true if the element is successfully found and removed; otherwise, false. 
        ''' This method returns false if key is not found in the Dictionary.
        ''' </returns>
        ''' <remarks></remarks>
        Public Overloads Function Remove(ByVal indexKey As IValue) As Boolean
            Dim strKey As String = CStr(m_colIndexHashTable(indexKey))
            m_colIndexHashTable.Remove(indexKey)
            Return MyBase.Remove(strKey.ToUpper)
        End Function

#End Region

#Region " ContainsKey "

        ''' <summary>
        ''' Determines whether the Dictionary contains the specified key.
        ''' </summary>
        ''' <param name="strKey">The key to locate in the Dictionary.</param>
        ''' <returns>true if the Dictionary contains an element with the specified key; otherwise, false.</returns>
        Public Overloads Function ContainsKey(ByVal strKey As String) As Boolean
            Return MyBase.ContainsKey(strKey.ToUpper)
        End Function

#End Region

#Region " Contains Index Key "

        ''' <summary>
        ''' Determines whether the Dictionary contains the specified index key.
        ''' </summary>
        ''' <param name="indexKey">The index key to locate in the Dictionary.</param>
        ''' <returns>true if the Dictionary contains an element with the specified index key; otherwise, false.</returns>
        Public Overloads Function ContainsIndexKey(ByVal indexKey As IValue) As Boolean
            Return m_colIndexHashTable.ContainsKey(indexKey)
        End Function

#End Region

#Region " Try Get Value "

        ''' <summary>
        ''' Gets the value associated with the specified key. 
        ''' 
        ''' Use this method if your code frequently attempts to access keys that are not in the dictionary.  
        ''' Using this method is more efficient than catching the KeyNotFoundException thrown by the Item property.
        ''' </summary>
        ''' <param name="strKey">The key of the value to get.</param>
        ''' <param name="objValue">
        ''' When this method returns, contains the value associated with the specified key, 
        ''' if the key is found; otherwise, the default value for the type of the value parameter. 
        ''' This parameter is passed uninitialized.</param>
        ''' <returns>true if the Dictionary contains an element with the specified key; otherwise, false.</returns>
        ''' <remarks>
        ''' This method combines the functionality of the ContainsKey method and the Item property.
        ''' 
        ''' If the key is not found, then the value parameter gets the appropriate default value for the value type 
        ''' TValue; for example, 0 (zero) for integer types, false for Boolean types, and a Nothing for reference types. 
        '''</remarks>
        Public Overloads Function TryGetValue(ByVal strKey As String, ByRef objValue As TValue) As Boolean
            Return MyBase.TryGetValue(strKey.ToUpper, objValue)
        End Function

        ''' <summary>
        ''' Gets the value associated with the specified index key. 
        ''' 
        ''' Use this method if your code frequently attempts to access keys that are not in the dictionary.  
        ''' Using this method is more efficient than catching the KeyNotFoundException thrown by the Item property.
        ''' </summary>
        ''' <param name="indexKey">The index key of the value to get.</param>
        ''' <param name="objValue">
        ''' When this method returns, contains the value associated with the specified index key, 
        ''' if the index key is found; otherwise, the default value for the type of the value parameter. 
        ''' This parameter is passed uninitialized.</param>
        ''' <returns>true if the Dictionary contains an element with the specified index key; otherwise, false.</returns>
        ''' <remarks>
        ''' This method combines the functionality of the ContainsKey method and the Item property.
        ''' 
        ''' If the index key is not found, then the value parameter gets the appropriate default value for the value type 
        ''' TValue; for example, 0 (zero) for integer types, false for Boolean types, and a Nothing for reference types. 
        '''</remarks>
        Public Overloads Function TryGetValue(ByVal indexKey As IValue, ByRef objValue As TValue) As Boolean
            Dim strKey As String = CStr(m_colIndexHashTable(indexKey))
            Return MyBase.TryGetValue(strKey, objValue)
        End Function

#End Region

#Region " Get Index Key "

        ''' <summary>
        ''' Gets the value to be used as the index key from the property of value that matches index property string
        ''' </summary>
        ''' <param name="objValue">The value we want to get the index key for</param>
        ''' <returns>Returns the value of the property</returns>
        Private Function GetIndexKey(ByVal objValue As TValue) As IValue
            Dim objPropertyInfo As Reflection.PropertyInfo = _
                objValue.GetType.GetProperty(m_strIndexPropertyName, Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public)
            Return CType(objPropertyInfo.GetValue(objValue, Nothing), IValue)
        End Function

#End Region

#Region " Insert Update Value "

        ''' <summary>
        ''' Will check if a entry exists by the index key, if so then replaces the value.
        ''' </summary>
        ''' <param name="indexKey">Index key of the entry we want to update</param>
        ''' <param name="objNewValue">new value of entry</param>
        ''' <param name="strKey">Key to use if entry does not exist</param>
        Public Sub InsertUpdateValue(ByVal indexKey As IValue, ByVal objNewValue As TValue, _
        ByVal strKey As String)
            Try
                If Me.ContainsIndexKey(indexKey) Then
                    Me.Item(indexKey) = objNewValue
                Else
                    Me.Add(strKey, objNewValue)
                End If
            Catch ex As Exception

            End Try
        End Sub

#End Region

#End Region

#Region " IDisposable Support "

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not m_blnDisposedValue Then
                If disposing Then
                    'For Each objValue As TValue In Me.Values
                    '    If objValue IsNot Nothing AndAlso TypeOf objValue Is IDisposable Then
                    '        CType(objValue, IDisposable).Dispose()
                    '        objValue = Nothing
                    '    End If
                    'Next

                    Me.Clear()
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

End Namespace