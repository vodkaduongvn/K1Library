Imports K1Library.clsDBConstants

Public Class clsSystemFlags
    Implements IDisposable

#Region " Members "

    Private m_blnDRMRunning As Boolean
    Private m_intUserProfileID As Integer
    Private m_dtDRMCheck As Date
    Private m_dtDRMRefreshed As Date
    Private m_dtDRMForcedLogoff As Date
    Private m_blnDRMLocked As Boolean
    Private m_objDB As clsDB
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

    Public Const cTIMEOUT_MIN As Integer = 5

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB)
        m_objDB = objDB
        Refresh(objDB)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property DRMRunning() As Boolean
        Get
            Return m_blnDRMRunning
        End Get
    End Property

    Public ReadOnly Property DRMUserProfileID() As Integer
        Get
            Return m_intUserProfileID
        End Get
    End Property

    Public ReadOnly Property DRMCheckDate() As Date
        Get
            Return m_dtDRMCheck
        End Get
    End Property

    Public ReadOnly Property DRMRefreshedDate() As Date
        Get
            Return m_dtDRMRefreshed
        End Get
    End Property

    Public ReadOnly Property DRMForcedLogoffDate() As Date
        Get
            Return m_dtDRMForcedLogoff
        End Get
    End Property

    Public ReadOnly Property DRMLocked() As Boolean
        Get
            Return m_blnDRMLocked
        End Get
    End Property

#End Region

#Region " Methods "

#Region " DRM Methods "

    Public Shared Function StartDRM(ByVal objDB As clsDB, ByVal blnLocked As Boolean) As Boolean
        Dim objSF As clsSystemFlags
        Dim dtNow As Date = objDB.GetCurrentTime

        objSF = New clsSystemFlags(objDB)
        If objSF.DRMRunning AndAlso objSF.m_dtDRMCheck >= dtNow.Subtract(New TimeSpan(0, cTIMEOUT_MIN, 0)) Then
            '-- Can not start DRM if someone is already using it
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cUSERPROFILE)
            Dim strUser As String = objDB.GetRecordExternalID(objTable, objSF.m_intUserProfileID)
            Throw New clsK1Exception("DRM is currently in use by " & strUser)
        End If

        Dim strSQL As String
        Dim intCount As Integer = objDB.ExecuteScalar("SELECT COUNT(*) FROM [" & clsDBConstants.Tables.cK1SYSTEMFLAGS & "] ")
        If intCount > 0 Then
            strSQL = "UPDATE [{0}] SET [{1}] = 1, [{2}] = {5}, [{3}] = GETDATE(), [{4}] = {6}"
        Else
            strSQL = "INSERT INTO [{0}] ([{1}], [{2}], [{3}], [{4}]) VALUES (1, {5}, GETDATE(), {6})"
        End If

        objDB.ExecuteSQL(String.Format(strSQL, clsDBConstants.Tables.cK1SYSTEMFLAGS, _
                                                  clsDBConstants.Fields.K1SystemFlags.cDRM_RUNNING, _
                                                  clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, _
                                                  clsDBConstants.Fields.K1SystemFlags.cDRM_CHECK_DATE, _
                                                  clsDBConstants.Fields.K1SystemFlags.cDRM_LOCKED, _
                                                  objDB.Profile.ID, _
                                                  CInt(blnLocked)))
        Threading.Thread.Sleep(300)

        objSF.Refresh(objDB)
        If objSF.DRMUserProfileID = objDB.Profile.ID Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub UpdateDRM(ByVal objDB As clsDB)
        Dim strSQL As String = "UPDATE [{0}] SET [{1}] = @Date WHERE [{2}] = {3}"

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@Date", objDB.GetCurrentTime))

        objDB.ExecuteSQL(String.Format(strSQL, clsDBConstants.Tables.cK1SYSTEMFLAGS, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_CHECK_DATE, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, _
            objDB.Profile.ID), colParams)
    End Sub

    Public Shared Sub StopDRM(ByVal objDB As clsDB, ByVal blnRefresh As Boolean)
        Dim objSF As clsSystemFlags = New clsSystemFlags(objDB)
        Dim dtRefreshed As DateTime = objSF.DRMRefreshedDate
        Dim dtForceLogoff As DateTime = objSF.DRMForcedLogoffDate

        If blnRefresh Then
            dtRefreshed = objDB.GetCurrentTime
        ElseIf dtRefreshed = Date.MinValue Then
            dtRefreshed = New DateTime(2000, 1, 1) '-- set a null value in range of db date value
        End If

        If objSF.IsDRMLocked Then
            dtForceLogoff = objDB.GetCurrentTime
        ElseIf dtForceLogoff = Date.MinValue Then
            dtForceLogoff = New DateTime(2000, 1, 1) '-- set a null value in range of db date value
        End If

        Dim strSQL As String = "UPDATE [{0}] SET [{1}] = 0, [{2}] = 0, " & _
            "[{3}] = @Date1, [{4}] = @Date2 WHERE [{5}] = {6}"

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@Date1", dtRefreshed))
        colParams.Add(New clsDBParameter("@Date2", dtForceLogoff))

        objDB.ExecuteSQL(String.Format(strSQL, clsDBConstants.Tables.cK1SYSTEMFLAGS, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_RUNNING, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_LOCKED, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_REFRESH_DATE, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_FORCED_LOGOFF_DATE, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, _
            objDB.Profile.ID), colParams)
    End Sub

    Public Sub LockDRM(ByVal objDB As clsDB, ByVal blnUpdateForcedLogoffDate As Boolean)
        ToggleLockDRM(objDB, True, blnUpdateForcedLogoffDate)

        '-- wait a couple of seconds for notifications to do its thing
        Threading.Thread.Sleep(2000)

        If blnUpdateForcedLogoffDate Then
            objDB.ExecuteSQL("DELETE FROM " & clsDBConstants.Tables.cK1SESSION)
        End If
    End Sub

    Public Sub UnlockDRM(ByVal objDB As clsDB, ByVal blnUpdateForcedLogoffDate As Boolean)
        If IsDRMLocked() Then
            ToggleLockDRM(objDB, False, blnUpdateForcedLogoffDate)
        End If
    End Sub

    Private Sub ToggleLockDRM(ByVal objDB As clsDB, ByVal blnLocked As Boolean, ByVal blnUpdateForcedLogoffDate As Boolean)
        Dim dtForceLogoff As Date = Me.m_dtDRMForcedLogoff
        Dim strSQL As String = "UPDATE [{0}] SET [{1}] = {2}, [{3}] = @Date WHERE [{4}] = {5}"

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@Date", dtForceLogoff))

        If blnLocked AndAlso blnUpdateForcedLogoffDate Then
            dtForceLogoff = objDB.GetCurrentTime
        End If

        objDB.ExecuteSQL(String.Format(strSQL, clsDBConstants.Tables.cK1SYSTEMFLAGS, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_LOCKED, _
            CInt(blnLocked), _
            clsDBConstants.Fields.K1SystemFlags.cDRM_FORCED_LOGOFF_DATE, _
            clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, _
            objDB.Profile.ID), colParams)

        m_blnDRMLocked = blnLocked
        m_dtDRMForcedLogoff = dtForceLogoff
    End Sub

    Public Function IsDRMLocked() As Boolean
        Try
            'If Me.DRMRunning AndAlso Me.DRMLocked Then
            '    strMessage = "DRM is currently running, can not login at this point in time."
            'End If
            Dim dtNow As Date = m_objDB.GetCurrentTime

            Return (Me.DRMRunning AndAlso Me.DRMLocked) AndAlso (Me.DRMCheckDate > dtNow.Subtract(New TimeSpan(0, cTIMEOUT_MIN, 20)))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function IsDRMLocked(ByVal objDB As clsDB) As Boolean
        Try
            Dim objSystemFlags As New clsSystemFlags(objDB)
            Return objSystemFlags.IsDRMLocked() 'strMessage)
        Catch ex As Exception
            Return True
        End Try
    End Function

    Public Shared Sub DoDRMRefresh(ByVal objDB As clsDB)
        Dim strSQL As String = "UPDATE [{0}] SET [{1}] = GETDATE() WHERE [{2}] = {3}"

        objDB.ExecuteSQL(String.Format(strSQL, clsDBConstants.Tables.cK1SYSTEMFLAGS, _
                                                clsDBConstants.Fields.K1SystemFlags.cDRM_REFRESH_DATE, _
                                                clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, _
                                                objDB.Profile.ID))
    End Sub

#End Region

    Public Sub Refresh(ByVal objDB As clsDB)
        Dim objDT As DataTable = objDB.GetDataTableBySQL(
           "SELECT TOP 1 * FROM [" & clsDBConstants.Tables.cK1SYSTEMFLAGS & "]")

        If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
            Dim objDR As DataRow = objDT.Rows(0)

            m_blnDRMRunning = CBool(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_RUNNING, False))
            m_intUserProfileID = CInt(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID, clsDBConstants.cintNULL))
            m_dtDRMCheck = CDate(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_CHECK_DATE, DateTime.MinValue))
            m_dtDRMRefreshed = CDate(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_REFRESH_DATE, DateTime.MinValue))
            m_blnDRMLocked = CBool(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_LOCKED, False))
            m_dtDRMForcedLogoff = CDate(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1SystemFlags.cDRM_FORCED_LOGOFF_DATE, DateTime.MinValue))
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If
    End Sub

    Public Shared Sub Synchronise(objDB As clsDB)
        Dim dtRefreshed As Date = objDB.GetCurrentTime

        Dim strSQL = $"UPDATE [{clsDBConstants.Tables.cK1SYSTEMFLAGS}]
                       SET [{clsDBConstants.Fields.K1SystemFlags.cDRM_REFRESH_DATE}] = @Date1
                       WHERE [{clsDBConstants.Fields.K1SystemFlags.cDRM_USERPROFILEID}] = {objDB.Profile.ID}"

        Dim colParams As New clsDBParameterDictionary From {
            New clsDBParameter("@Date1", dtRefreshed)
        }

        objDB.ExecuteSQL(strSQL, colParams)
    End Sub

#End Region

#Region " IDisposable Support "
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
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
