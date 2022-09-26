Namespace FrameworkCollections

    Public Class K1Collection(Of TValue)
        Inherits System.Collections.Generic.List(Of TValue)
        Implements IDisposable

#Region " Members "

        Private m_blnDisposedValue As Boolean = False       '-- To detect redundant calls
#End Region

#Region " IDisposable Support "

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not m_blnDisposedValue Then
                If disposing Then
                    'For Each objValue As TValue In Me
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
