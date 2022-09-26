Public Class Trace

#Region " Members "

    'Create a TraceSwitch to use in the entire application. 
    Private Shared g_objSwitch As New TraceSwitch("TraceSwitch", "Tracing switch for the entire application")
    Private Shared m_colIndents As New Generic.Dictionary(Of String, Integer)

#End Region

#Region " Enumerators "

    Public Enum TraceFormatting As Integer
        None = 0
        Indent
        UnIndent
    End Enum

#End Region

#Region " Properties "

#Region " Indent "

    ''' <summary>
    ''' Gets the indent level for a specific thread
    ''' </summary>
    ''' <param name="strThreadName">Name of thread to get the indent level for.</param>
    Private Shared Property Indent(ByVal strThreadName As String) As Integer
        Get
            If m_colIndents.ContainsKey(strThreadName) Then
                Return m_colIndents(strThreadName)
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then
                value = 0
            End If
            m_colIndents(strThreadName) = value
        End Set
    End Property

#End Region

#End Region

#Region " Public "

#Region " Write Error "

    Public Shared Sub WriteError(ByVal strMessage As String, ByVal strCategory As String, _
        ByVal blnCondition As Boolean, ByVal eTraceFormatting As TraceFormatting)
        If g_objSwitch.TraceError AndAlso blnCondition Then
            WriteLine(strMessage, strCategory, eTraceFormatting)
        End If
    End Sub

    Public Shared Sub WriteError(ByVal strMessage As String, ByVal strCategory As String)
        WriteError(strMessage, strCategory, True, TraceFormatting.None)
    End Sub

    Public Shared Sub WriteError(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal eTraceFormatting As TraceFormatting)
        WriteError(strMessage, strCategory, True, eTraceFormatting)
    End Sub

#End Region

#Region " Write Info "

    Public Shared Sub WriteInfo(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal blnCondition As Boolean, ByVal eTraceFormatting As TraceFormatting)
        If g_objSwitch.TraceInfo AndAlso blnCondition Then
            Trace.WriteLine(strMessage, strCategory, eTraceFormatting)
        End If
    End Sub

    Public Shared Sub WriteInfo(ByVal strMessage As String, ByVal strCategory As String)
        Trace.WriteInfo(strMessage, strCategory, True, TraceFormatting.None)
    End Sub

    Public Shared Sub WriteInfo(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal eTraceFormatting As TraceFormatting)
        Trace.WriteInfo(strMessage, strCategory, True, eTraceFormatting)
    End Sub

#End Region

#Region " Write Warning "

    Public Shared Sub WriteWarning(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal blnCondition As Boolean, ByVal eTraceFormatting As TraceFormatting)
        If g_objSwitch.TraceWarning AndAlso blnCondition Then
            WriteLine(strMessage, strCategory, eTraceFormatting)
        End If
    End Sub

    Public Shared Sub WriteWarning(ByVal strMessage As String, ByVal strCategory As String)
        WriteWarning(strMessage, strCategory, True, TraceFormatting.None)
    End Sub

    Public Shared Sub WriteWarning(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal eTraceFormatting As TraceFormatting)
        WriteWarning(strMessage, strCategory, True, eTraceFormatting)
    End Sub

#End Region

#Region " Write Verbose "

    Public Shared Sub WriteVerbose(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal blnCondition As Boolean, ByVal eTraceFormatting As TraceFormatting)
        If g_objSwitch.TraceVerbose AndAlso blnCondition Then
            WriteLine(strMessage, strCategory, eTraceFormatting)
        End If
    End Sub

    Public Shared Sub WriteVerbose(ByVal strMessage As String, ByVal strCategory As String)
        WriteVerbose(strMessage, strCategory, True, TraceFormatting.None)
    End Sub

    Public Shared Sub WriteVerbose(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal eTraceFormatting As TraceFormatting)
        WriteVerbose(strMessage, strCategory, True, eTraceFormatting)
    End Sub

#End Region

#End Region

#Region " Private "

#Region " Write Line "

    ''' <summary>
    ''' Writes a category name and message to the trace listeners in the 
    ''' System.Diagnostics.Trace.Listeners collection.
    ''' </summary>
    ''' <param name="strMessage">A message to write.</param>
    ''' <param name="strCategory">A category name used to organize the output.</param>
    ''' <remarks></remarks>
    Private Shared Sub WriteLine(ByVal strMessage As String, ByVal strCategory As String, _
    ByVal eTraceFormatting As TraceFormatting)
        Dim strThreadName As String = GetCurrentThreadName()

        If eTraceFormatting = TraceFormatting.UnIndent Then
            Trace.Indent(strThreadName) -= 1
        End If

        System.Diagnostics.Trace.WriteLine(
            String.Format("{4} {2}[{0}] {1}",
                strThreadName,
                strMessage,
                GetPadding(strThreadName),
                strCategory,
                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")))

        If eTraceFormatting = TraceFormatting.Indent Then
            Trace.Indent(strThreadName) += 1
        End If
    End Sub

#End Region

#Region " Get Current Thread Name "

    ''' <summary>
    ''' Gets the name or managed thread id of the current thread
    ''' </summary>
    Private Shared Function GetCurrentThreadName() As String
        Try
            Dim objThread As Threading.Thread = Threading.Thread.CurrentThread
            If String.IsNullOrEmpty(objThread.Name) Then
                Return objThread.ManagedThreadId.ToString
            Else
                Return objThread.Name
            End If
        Catch ex As Exception
            Return "ERR"
        End Try
    End Function

#End Region

#Region " Get Padding "

    ''' <summary>
    ''' Returns the padding required for a specified thread
    ''' </summary>
    ''' <param name="strThreadName">Name of thread we want the display padding for</param>
    Private Shared Function GetPadding(ByVal strThreadName As String) As String
        Dim intIndent As Integer = Trace.Indent(strThreadName) * 4
        Return "".PadLeft(intIndent)
    End Function

#End Region

#End Region

End Class
