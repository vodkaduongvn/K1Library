Public Class clsRecordLock

#Region " Members "

    Private m_intID As Integer
    Private m_intRecordID As Integer
    Private m_intTableID As Integer
    Private m_intProfileID As Integer
    Private m_strUser As String
    Private m_dtStart As Date
#End Region

#Region " Constructors "

    Public Sub New(ByVal intTableID As Integer, ByVal intRecordID As Integer, _
    ByVal intProfileID As Integer, ByVal strUser As String, ByVal dtTime As Date, _
    Optional ByVal intID As Integer = clsDBConstants.cintNULL)
        m_intID = intID
        m_intTableID = intTableID
        m_intRecordID = intRecordID
        m_intProfileID = intProfileID
        m_strUser = strUser
        m_dtStart = dtTime
    End Sub

    Private Sub New(ByVal objDR As DataRow)
        m_intID = CInt(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cID, clsDBConstants.cintNULL))
        m_intTableID = CInt(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1RecordLock.cTABLEID, clsDBConstants.cintNULL))
        m_intRecordID = CInt(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1RecordLock.cRECORDID, clsDBConstants.cintNULL))
        m_intProfileID = CInt(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1RecordLock.cUSERPROFILEID, clsDBConstants.cintNULL))
        m_dtStart = CDate(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1RecordLock.cTIMELOCKED, DateTime.MinValue))
        m_strUser = CStr(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1RecordLock.cUSERNAME, clsDBConstants.cstrNULL))
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property ID() As Integer
        Get
            Return m_intID
        End Get
    End Property

    Public ReadOnly Property RecordID() As Integer
        Get
            Return m_intRecordID
        End Get
    End Property

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property StartDate() As Date
        Get
            Return m_dtStart
        End Get
    End Property

    Public ReadOnly Property ProfileID() As Integer
        Get
            Return m_intProfileID
        End Get
    End Property

    Public ReadOnly Property User() As String
        Get
            Return m_strUser
        End Get
    End Property
#End Region

#Region " Methods "

    Public Shared Sub AddLock(ByVal objDB As clsDB, ByVal colLocks As K1Library.FrameworkCollections.K1Dictionary(Of clsRecordLock), _
    ByVal intRecordLockID As Integer, ByVal intTableID As Integer, ByVal intRecordID As Integer, ByVal dtTime As Date)
        Dim intProfileID As Integer
        Dim strUser As String

        If objDB.Profile Is Nothing Then
            intProfileID = objDB.SysInfo.K1Configuration.DefaultProfile.ID
            strUser = objDB.SysInfo.K1Configuration.DefaultProfile.ExternalID
        Else
            intProfileID = objDB.Profile.ID
            strUser = objDB.Profile.ExternalID
        End If

        Dim objLock As New clsRecordLock(intTableID, intRecordID, intProfileID, strUser, dtTime, intRecordLockID)

        colLocks.Add(intTableID & "_" & intRecordID, objLock)
    End Sub

    Public Shared Function GetLock(ByVal objDB As clsDB, ByVal colLocks As K1Library.FrameworkCollections.K1Dictionary(Of clsRecordLock), _
    ByVal intTableID As Integer, ByVal intRecordID As Integer) As String
        Dim intTimeOut As Integer
        Dim strReturn As String = Nothing
        Dim blnSuccess As Boolean = False
        Dim dtNow As Date = objDB.GetCurrentTime
        Dim intProfileID As Integer
        Dim strUser As String

        If objDB.Profile Is Nothing Then
            intProfileID = objDB.SysInfo.K1Configuration.DefaultProfile.ID
            strUser = objDB.SysInfo.K1Configuration.DefaultProfile.ExternalID
        Else
            intProfileID = objDB.Profile.ID
            strUser = objDB.Profile.ExternalID
        End If

        intTimeOut = objDB.SysInfo.K1Configuration.RecordLockTimeout

        While Not blnSuccess
            Dim objLock As clsRecordLock = clsRecordLock.GetItem(objDB, intTableID, intRecordID)

            If objLock IsNot Nothing Then
                'check whether or not the lock has expired
                If dtNow.Subtract(objLock.StartDate).TotalMinutes <= intTimeOut Then
                    If Not objLock.ProfileID = intProfileID Then
                        strReturn = objLock.User 'The record is locked by another user
                    Else
                        If colLocks(intTableID & "_" & intRecordID) Is Nothing Then
                            strReturn = objLock.User & " on another Session"
                        Else
                            strReturn = objLock.User & " on this Session"
                        End If
                    End If
                Else
                    'the lock has expired, release it
                    objLock.Delete(objDB)
                    colLocks.Remove(intTableID & "_" & intRecordID)
                    objLock = Nothing
                End If
            End If

            If objLock Is Nothing Then
                objLock = New clsRecordLock(intTableID, intRecordID, intProfileID, strUser, objDB.GetCurrentTime)
                If objLock.Insert(objDB) Then
                    colLocks.Add(intTableID & "_" & intRecordID, objLock)
                    blnSuccess = True
                Else
                    blnSuccess = False
                End If
            Else
                blnSuccess = True
            End If
        End While

        Return strReturn
    End Function

    Public Shared Function CheckLock(ByVal objDB As clsDB, ByVal colLocks As K1Library.FrameworkCollections.K1Dictionary(Of clsRecordLock), _
    ByVal intTableID As Integer, ByVal intRecordID As Integer) As Boolean
        Dim intTimeOut As Integer = objDB.SysInfo.K1Configuration.RecordLockTimeout
        Dim dtNow As Date = objDB.GetCurrentTime
        Dim intProfileID As Integer

        If objDB.Profile Is Nothing Then
            intProfileID = objDB.SysInfo.K1Configuration.DefaultProfile.ID
        Else
            intProfileID = objDB.Profile.ID
        End If

        Dim objLock As clsRecordLock = clsRecordLock.GetItem(objDB, intTableID, intRecordID)
        Dim objCompareLock As clsRecordLock = colLocks(intTableID & "_" & intRecordID)

        Dim blnReturn As Boolean = False

        If objLock Is Nothing OrElse _
        Not objLock.ProfileID = intProfileID OrElse _
        objCompareLock Is Nothing OrElse _
        Not objLock.ID = objCompareLock.ID Then
            Return blnReturn
        Else
            If dtNow.Subtract(objLock.StartDate).TotalMinutes > intTimeOut Then
                blnReturn = objLock.Update(objDB) 'Extend the timeout if possible
            Else
                blnReturn = True
            End If

            Return blnReturn
        End If
    End Function

    Public Shared Sub ReleaseLock(ByVal objDB As clsDB, ByVal colLocks As K1Library.FrameworkCollections.K1Dictionary(Of clsRecordLock), _
    ByVal intTableID As Integer, ByVal intRecordID As Integer)
        Try
            Dim objLock As clsRecordLock = colLocks(intTableID & "_" & intRecordID)

            'only remove the lock if it exists and was created by this session
            If objLock IsNot Nothing Then
                objLock.Delete(objDB)
                colLocks.Remove(intTableID & "_" & intRecordID)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Shared Function GetItem(ByVal objDB As clsDB, ByVal intTableID As Integer, _
    ByVal intRecordID As Integer) As clsRecordLock
        Dim objRecordLock As clsRecordLock = Nothing

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@TableID", intTableID))
        colParams.Add(New clsDBParameter("@RecordID", intRecordID))
        Dim strSQL As String = "SELECT * FROM [" & clsDBConstants.Tables.cK1RECORDLOCK & "] " & _
            "WHERE [" & clsDBConstants.Fields.K1RecordLock.cTABLEID & "] = @TableID " & _
            "AND [" & clsDBConstants.Fields.K1RecordLock.cRECORDID & "] = @RecordID"
        Dim objDT As DataTable = objDB.GetDataTableBySQL(strSQL, colParams)
        colParams.Dispose()

        If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
            objRecordLock = New clsRecordLock(objDT.Rows(0))
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If

        Return objRecordLock
    End Function

    Private Function Insert(ByVal objDB As clsDB) As Boolean
        Try
            Dim colParams As New clsDBParameterDictionary

            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
                clsDBConstants.Fields.cID), clsDBConstants.cintNULL, ParameterDirection.Output))
            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
                clsDBConstants.Fields.K1RecordLock.cTABLEID), m_intTableID))
            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
                clsDBConstants.Fields.K1RecordLock.cRECORDID), m_intRecordID))
            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
                clsDBConstants.Fields.K1RecordLock.cUSERPROFILEID), m_intProfileID))
            colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
                clsDBConstants.Fields.K1RecordLock.cUSERNAME), m_strUser))

            objDB.ExecuteProcedure(clsDBConstants.StoredProcedures.cK1_RECORDLOCK_INSERT, colParams)

            m_intID = CInt(colParams(clsDB_Direct.ParamName(clsDBConstants.Fields.cID)).Value)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Update(ByVal objDB As clsDB) As Boolean
        Dim colParams As New clsDBParameterDictionary

        colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
            clsDBConstants.Fields.cID), m_intID))

        Dim intRows As Integer = objDB.ExecuteSQL("UPDATE [" & clsDBConstants.Tables.cK1RECORDLOCK & "] " & _
            "SET [TimeLocked] = GETDATE() WHERE [" & clsDBConstants.Fields.cID & "] = @ID", colParams)

        Return (intRows > 0)
    End Function

    Public Sub Delete(ByVal objDB As clsDB)
        Dim colParams As New clsDBParameterDictionary

        colParams.Add(New clsDBParameter(clsDB_Direct.ParamName( _
            clsDBConstants.Fields.cID), m_intID))

        objDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cK1RECORDLOCK & "] " & _
            "WHERE [" & clsDBConstants.Fields.cID & "] = @ID", colParams)
    End Sub
#End Region

End Class
