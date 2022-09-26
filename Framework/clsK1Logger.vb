Public NotInheritable Class clsK1Logger
    Implements IDisposable

    Public Property strApplicationName As String = "K1Default"

    Private objWriter As StreamWriter = Nothing
    Private colMessages As New Concurrent.ConcurrentQueue(Of LogDetails)
    Private Event LogRaisedEvent()

    Private Sub New()
        AddHandler LogRaisedEvent, AddressOf WriteLog
    End Sub

    Private Shared ReadOnly lazy As Lazy(Of clsK1Logger) = New Lazy(Of clsK1Logger)(Function() New clsK1Logger())
    Public Shared ReadOnly Property Instance As clsK1Logger
        Get
            Return lazy.Value
        End Get
    End Property

    Public Sub Log(message As String,
                         <System.Runtime.CompilerServices.CallerLineNumber> Optional lineNumber As Integer = 0,
                         <System.Runtime.CompilerServices.CallerFilePath> Optional fileName As String = "",
                         <System.Runtime.CompilerServices.CallerMemberName> Optional memberName As String = "")
        colMessages.Enqueue(New LogDetails() With {
            .strFileName = fileName,
            .intLineNumber = lineNumber,
            .strMessage = message,
            .strMethod = memberName})
        RaiseEvent LogRaisedEvent()
    End Sub

    Private Sub WriteLog() '(source As Object, e As Timers.ElapsedEventArgs)
        ' Write the log to the file if the writer is free
        If colMessages?.Any() AndAlso objWriter Is Nothing Then
            Open()
            Dim copyDetails As LogDetails = Nothing

            While colMessages.TryDequeue(copyDetails)
                Log(copyDetails)
            End While

            Close()
        End If
    End Sub

    Private Sub Log(details As LogDetails)
        Try
            objWriter?.WriteLine($"{details.dtTimeStamp:dd MMMM yyyy HH:mm:ss.fff} - {details.strFileName} 
                                Function:{details.strMethod} Line:{details.intLineNumber}
                                {details.strMessage} {Environment.NewLine}")
        Catch ex As Exception
            ' Requeue the item.
            colMessages.Enqueue(details)
        End Try
    End Sub

    Private Sub Open()
        Dim logPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Knowledgeone Corporation", strApplicationName, "Logs")
        Try
            If Not Directory.Exists(logPath) Then
                Directory.CreateDirectory(logPath)
            End If
            objWriter = New StreamWriter(logPath + "\\" + "Logs-" + Date.Now.ToString("dd-MMM-yyyy") + ".log", True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Close()
        Try
            objWriter?.Close()
            objWriter?.Dispose()
            objWriter = Nothing
        Catch ex As Exception

        End Try
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                objWriter?.Dispose()
            Else
                RemoveHandler LogRaisedEvent, AddressOf WriteLog
                colMessages = Nothing
            End If
        End If
        disposedValue = True
    End Sub

    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
    End Sub
#End Region
    Private NotInheritable Class LogDetails

        Public Property strMessage As String
        Public Property intLineNumber As Integer
        Public Property strFileName As String
        Public Property strMethod As String
        Public Property dtTimeStamp As Date = Date.Now
    End Class
End Class
