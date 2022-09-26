Imports System.Data.SqlClient

''' <summary>
''' This class is a future candidate to replace SQL notifications (clsSqlDependency) in the current K1Library.
''' The current implementation in clsSqlDependency lacks abstraction and not very well thought out. The notification 
''' responsibility is spread out over too many classes for it to be maintainable.
''' </summary>
''' <remarks></remarks>
Public Class clsSqlNotifier

    Implements IDisposable

    Private objDataAccess As clsDB
    Private blnDisposed As Boolean
    Private strQuery As String

    Public Sub New(objDb As clsDB, strQuery As String)

        Me.strQuery = strQuery
        objDataAccess = objDb

        If (objDb.SqlDependency IsNot Nothing OrElse Not objDb.SqlDependency.Started) Then
            Throw New ApplicationException("SQL Server notification service is not yet started.")
        End If

    End Sub

    Public Sub Register(Optional strSqlQuery As String = "")

        If (Not String.IsNullOrEmpty(strSqlQuery)) Then
            strQuery = strSqlQuery
        End If

        If (String.IsNullOrEmpty(strQuery)) Then
            Throw New ArgumentException()
        End If

        objDataAccess.AddSqlNotification(strQuery, AddressOf OnChangeEventHandler)

    End Sub

    Private Sub OnChangeEventHandler(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)
        RemoveHandler CType(sender, SqlDependency).OnChange, AddressOf OnChangeEventHandler
        OnNotify(e)
    End Sub

    Public Event Notify(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)

    Protected Overridable Sub OnNotify(arg As SqlNotificationEventArgs)
        RaiseEvent Notify(Me, arg)
    End Sub

    Protected Overridable Sub Disposing(blnDisposing As Boolean)

        If (blnDisposing) Then
            'do some cleaning up here if needed...
            'well not really needed here now...
            If (Not blnDisposed) Then
                objDataAccess = Nothing
                strQuery = String.Empty
            End If
        End If

        blnDisposed = True

        GC.SuppressFinalize(Me)

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Disposing(True)
    End Sub

    Protected Overrides Sub Finalize()
        Disposing(False)
        MyBase.Finalize()
    End Sub

End Class
