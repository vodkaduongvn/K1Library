Imports System.Data.SqlClient

Public Class clsDBTransaction
    Implements IDisposable

#Region " Members "

    Protected m_objConnection As SqlConnection
    Protected m_objTransaction As SqlTransaction
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Properties "

    Public ReadOnly Property Connection() As SqlConnection
        Get
            Return m_objConnection
        End Get
    End Property

    Public ReadOnly Property Transaction() As SqlTransaction
        Get
            Return m_objTransaction
        End Get
    End Property
#End Region

#Region " Constructors "

    Public Sub New(ByVal objConnection As SqlConnection, ByVal objTransaction As SqlTransaction)
        m_objConnection = objConnection
        m_objTransaction = objTransaction
    End Sub
#End Region

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not Me.m_blnDisposedValue Then
            If blnDisposing Then
                If Not m_objTransaction Is Nothing Then
                    Try
                        m_objTransaction.Rollback()
                        m_objTransaction.Dispose()
                    Catch ex As Exception
                    End Try

                    m_objTransaction = Nothing
                End If

                If Not m_objConnection Is Nothing Then
                    Try
                        If m_objConnection.State = ConnectionState.Open Then
                            m_objConnection.Close()
                        End If
                        m_objConnection.Dispose()                        
                    Catch ex As Exception
                    End Try

                    m_objConnection = Nothing
                End If
            End If
        End If
        Me.m_blnDisposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Public Sub Commit()
        m_objTransaction.Commit()
        m_objTransaction = Nothing
    End Sub

    Public Sub Rollback()
        m_objTransaction.Rollback()
        m_objTransaction = Nothing
    End Sub

End Class
