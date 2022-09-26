Namespace FrameworkCollections

    ''' <summary>
    ''' Represents a generic collection of keys (String) and values (TValue).
    ''' 
    ''' If you have a collection of objects that you regulary use one of two keys to find it in the collection
    ''' use K1DualKeyDictionary as it is faster then K1Dictionary and other ICollection classes with a near O(1)
    ''' </summary>
    ''' <typeparam name="TValue">The datatype of the value you want to store in the Dictionary.</typeparam>
    Public Class K1Dictionary(Of TValue)
        Inherits System.Collections.Generic.Dictionary(Of String, TValue)
        Implements IDisposable

#Region " Members "

        Private m_blnDisposedValue As Boolean = False       '-- To detect redundant calls

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the Dictionary class that is empty, has the default initial capacity
        ''' and uses StringComparer.Ordinal for comparing keys
        ''' </summary>
        Public Sub New()
            MyBase.New(StringComparer.Ordinal)
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
                If Me.ContainsKey(strKey) Then
                    Return MyBase.Item(strKey.ToUpper)
                Else
                    Return Nothing
                End If
            End Get
            Set(ByVal value As TValue)
                MyBase.Item(strKey.ToUpper) = value
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
            Try
                MyBase.Add(strKey.ToUpper, objValue)
            Catch ex As Exception
                Trace.WriteInfo($"K1Dictionary.add ({strKey} {objValue} {Me})  {ex} ", "K1Dictionary")
                Throw
            End Try
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
            Return MyBase.Remove(strKey.ToUpper)
        End Function

#End Region

#Region " Contains Key "

        ''' <summary>
        ''' Determines whether the Dictionary contains the specified key.
        ''' </summary>
        ''' <param name="strKey">The key to locate in the Dictionary.</param>
        ''' <returns>true if the Dictionary contains an element with the specified key; otherwise, false.</returns>
        Public Overloads Function ContainsKey(ByVal strKey As String) As Boolean
            Return MyBase.ContainsKey(strKey.ToUpper)
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

#End Region

#End Region

#Region " IDisposable Support "

        Protected Overridable Sub DisposeObj()
        End Sub

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not m_blnDisposedValue Then
                If disposing Then
                    DisposeObj()

                    'TODO: I don't think we want to automatically dispose of objects, 
                    'especially if they are being used elsewhere for instance field objects

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