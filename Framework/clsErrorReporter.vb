Public Class clsErrorReporter

#Region " Members "

    Private m_intErrorCount As Integer
    Private m_intMaxErrorCount As Integer
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.new()
        m_intErrorCount = 0
        m_intMaxErrorCount = 10
    End Sub
#End Region

#Region " Enumerations "

    Public Enum enumErrorLogType
        DATABASE = 0
        EVENTLOG = 1
        FILELOG = 2
    End Enum

    Public Enum enumErrorMsgType
        NORMAL_SERVICE = 0
        SHUTDOWN_SERVICE = 1
    End Enum
#End Region

#Region " Properties "

    Public Property ConcurrentErrorCount() As Integer
        Get
            Return m_intErrorCount
        End Get
        Set(ByVal Value As Integer)
            m_intErrorCount = Value
        End Set
    End Property

    Public Property MaxErrorCount() As Integer
        Get
            Return m_intMaxErrorCount
        End Get
        Set(ByVal Value As Integer)
            m_intMaxErrorCount = Value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Sub ReportError(ByVal objDB As clsDB, ByVal objException As Exception, _
    Optional ByVal eMsgType As enumErrorMsgType = enumErrorMsgType.NORMAL_SERVICE, _
    Optional ByVal objTable As clsTable = Nothing)
        Try
            Dim strMessage As String = Nothing
            Dim eLogType As enumErrorLogType = enumErrorLogType.DATABASE

            '-- Generate Message to be saved in event log
            strMessage &= "[" & Now.ToLongDateString & " " & Now.ToLongTimeString & "]" & vbCrLf & vbCrLf
            strMessage &= My.Application.Info.ProductName & ":  Concurrent Error Count = " & m_intErrorCount & vbCrLf
            strMessage &= "----------------------------------------" & vbCrLf

            Select Case eMsgType
                Case enumErrorMsgType.NORMAL_SERVICE
                    If m_intErrorCount >= m_intMaxErrorCount Then
                        strMessage &= "The service has been automatically shutdown because of too many concurrent errors." & vbCrLf & vbCrLf
                    End If
                Case enumErrorMsgType.SHUTDOWN_SERVICE
                    strMessage &= "The following error has caused the service to be automatically shutdown:" & vbCrLf & vbCrLf
            End Select

            strMessage &= objException.Message & vbCrLf
            strMessage &= "----------------------------------------" & vbCrLf
            strMessage &= objException.StackTrace & vbCrLf

            If eLogType = enumErrorLogType.DATABASE Then
                Try
                    Dim objErrTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cERROR)

                    Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(objErrTable)

                    colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, strMessage.Substring(0, Math.Min(100, strMessage.Length)))                
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cERRORDATE, objDB.GetCurrentTime)

                    If objDB.Profile Is Nothing Then
                        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, objDB.SysInfo.K1Configuration.DRMDefaultSecurityID)
                    Else
                        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, objDB.Profile.NominatedSecurityID)
                        colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cPERSONID, objDB.Profile.PersonID)
                    End If

                    If objTable IsNot Nothing Then
                        colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cTABLEID, objTable.ID)
                    End If
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cERROR, strMessage)
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cOCCURREDIN, My.Application.Info.ProductName)
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.Error.cSTACKTRACE, objException.StackTrace)

                    colMasks.Insert(objDB)
                Catch ex As Exception
                    eLogType = enumErrorLogType.EVENTLOG
                End Try
            End If

            If eLogType = enumErrorLogType.EVENTLOG Then
                Try
                    Dim objEventLogger As New EventLog
                    objEventLogger.Source = My.Application.Info.ProductName
                    '-- Writes Error to Custom Event Log
                    objEventLogger.WriteEntry(strMessage, EventLogEntryType.Error)
                Catch ex As Exception
                    eLogType = enumErrorLogType.FILELOG
                End Try
            End If

            If eLogType = enumErrorLogType.FILELOG Then
                Dim strErrorLogFile As String = ProperPath(My.Application.Info.DirectoryPath) & "Error.log"

                Dim objTW As New System.IO.StreamWriter(strErrorLogFile, True)

                objTW.WriteLine("===================================================================")
                objTW.Write(strMessage)
                objTW.WriteLine("===================================================================")
                objTW.WriteLine("")
                objTW.WriteLine("")

                objTW.Close()
            End If
        Catch ex As Exception
            'oh well, we tried
        End Try
    End Sub
#End Region

End Class
