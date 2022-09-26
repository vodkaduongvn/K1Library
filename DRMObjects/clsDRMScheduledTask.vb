Public Class clsDRMScheduledTask
    Inherits clsDRMBase

#Region " Members "

    Private m_blnActive As Boolean
    Private m_dtBeginDate As Date
    Private m_dtEndDate As Date
    Private m_dtTimeDue As Date
    Private m_intPeriodID As Integer
    Private m_blnReturnsMailList As Boolean
    Private m_strSenderEmail As String
    Private m_strSMTPServer As String
    Private m_strStoredProcedure As String
    Private m_strSQL As String
    Private m_blnBodyHTML As Boolean
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal objDB As clsDB, ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, ByVal blnActive As Boolean, ByVal dtBeginDate As Date, _
    ByVal dtEndDate As Date, ByVal dtTimeDue As Date, ByVal intPeriodID As Integer, _
    ByVal blnReturnsMailList As Boolean, ByVal strSenderEmail As String, _
    ByVal strSMTPServer As String, ByVal strStoredProcedure As String, ByVal strSQL As String, _
    ByVal blnBodyHTML As Boolean)
        MyBase.New(objDB, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_blnActive = blnActive
        m_dtBeginDate = dtBeginDate
        m_dtEndDate = dtEndDate
        m_dtTimeDue = dtTimeDue
        m_intPeriodID = intPeriodID
        m_blnReturnsMailList = blnReturnsMailList
        m_strSenderEmail = strSenderEmail
        m_strSMTPServer = strSMTPServer
        m_strStoredProcedure = strStoredProcedure
        m_strSQL = strSQL
        m_blnBodyHTML = blnBodyHTML
    End Sub
#End Region

#Region " Existing "

    ''' <summary>
    ''' Get an existing scheduled task
    ''' </summary>
    ''' <param name="objDB"></param>
    ''' <param name="strStoredProcedure"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objDB As clsDB, ByVal strStoredProcedure As String)
        MyBase.New(objDB, clsDBConstants.cstrNULL, clsDBConstants.cintNULL, clsDBConstants.cintNULL)

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@Name", strStoredProcedure))
        Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT [ID] FROM [" & _
            clsDBConstants.Tables.cSCHEDULEDTASK & "] " & _
            "WHERE [StoredProcedureName] = @Name", colParams)
        colParams.Dispose()

        If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
            Dim objST As clsScheduledTask = clsScheduledTask.GetItem(CInt(objDT.Rows(0)(0)), objDB)
            m_intID = objST.ID
            m_strExternalID = objST.ExternalID
            m_intSecurityID = objST.SecurityID
            m_intTypeID = objST.TypeID
            m_objDBObj = objST
            LoadFromScheduledTask(objST)
        Else
            Throw New Exception("scheduled task does not exist.")
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal objScheduledTask As clsScheduledTask)
        MyBase.New(objDB, objScheduledTask)
        LoadFromScheduledTask(objScheduledTask)
    End Sub

    Private Sub LoadFromScheduledTask(ByVal objScheduledTask As clsScheduledTask)
        m_blnActive = objScheduledTask.IsActive
        m_dtBeginDate = objScheduledTask.BeginDate
        m_dtEndDate = objScheduledTask.EndDate
        m_dtTimeDue = objScheduledTask.TimeDue
        m_intPeriodID = objScheduledTask.PeriodID
        m_blnReturnsMailList = objScheduledTask.ReturnsMailList
        m_strSenderEmail = objScheduledTask.SenderEmail
        m_strSMTPServer = objScheduledTask.SMTPServer
        m_strStoredProcedure = objScheduledTask.StoredProcedure
        m_strSQL = objScheduledTask.SQL
        m_blnBodyHTML = objScheduledTask.IsHTML
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property ScheduledTask() As clsScheduledTask
        Get
            Return CType(m_objDBObj, clsScheduledTask)
        End Get
        Set(ByVal value As clsScheduledTask)
            m_objDBObj = value
        End Set
    End Property

    Public Property IsActive() As Boolean
        Get
            Return m_blnActive
        End Get
        Set(ByVal value As Boolean)
            m_blnActive = value
        End Set
    End Property

    Public Property BeginDate() As Date
        Get
            Return m_dtBeginDate
        End Get
        Set(ByVal value As Date)
            m_dtBeginDate = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return m_dtEndDate
        End Get
        Set(ByVal value As Date)
            m_dtEndDate = value
        End Set
    End Property

    Public Property TimeDue() As Date
        Get
            Return m_dtTimeDue
        End Get
        Set(ByVal value As Date)
            m_dtTimeDue = value
        End Set
    End Property

    Public Property PeriodID() As Integer
        Get
            Return m_intPeriodID
        End Get
        Set(ByVal value As Integer)
            m_intPeriodID = value
        End Set
    End Property

    Public Property ReturnsMailList() As Boolean
        Get
            Return m_blnReturnsMailList
        End Get
        Set(ByVal value As Boolean)
            m_blnReturnsMailList = value
        End Set
    End Property

    Public Property SenderEmail() As String
        Get
            Return m_strSenderEmail
        End Get
        Set(ByVal value As String)
            m_strSenderEmail = value
        End Set
    End Property

    Public Property SMTPServer() As String
        Get
            Return m_strSMTPServer
        End Get
        Set(ByVal value As String)
            m_strSMTPServer = value
        End Set
    End Property

    Public ReadOnly Property StoredProcedure() As String
        Get
            Return m_strStoredProcedure
        End Get
    End Property

    Public Property SQL() As String
        Get
            Return m_strSQL
        End Get
        Set(ByVal value As String)
            m_strSQL = value
        End Set
    End Property

    Public Property IsHTML() As Boolean
        Get
            Return m_blnBodyHTML
        End Get
        Set(ByVal value As Boolean)
            m_blnBodyHTML = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Insert/Update Scheduled Task "

    ''' <summary>
    ''' Inserts or updates a scheduled task in to the system
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            m_objDB.DeleteStoredProcedure(m_strStoredProcedure)
            m_objDB.CreateStoredProcedure(m_strStoredProcedure, m_strSQL)

            Dim objTable As clsTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cSCHEDULEDTASK)

            Dim colMasks As clsMaskFieldDictionary
            Dim intID As Integer = clsDBConstants.cintNULL
            If ScheduledTask IsNot Nothing Then
                intID = ScheduledTask.ID
            End If

            colMasks = clsMaskField.CreateMaskCollection(objTable, intID)

            colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cTIMEDUE, m_dtTimeDue)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cBEGINDATE, m_dtBeginDate)
            If Not m_dtEndDate = Date.MinValue Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cENDDATE, m_dtEndDate)
            Else
                colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cENDDATE, DBNull.Value)
            End If
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cISACTIVE, m_blnActive)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cPERIODID, m_intPeriodID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cRETURNSMAILLIST, m_blnReturnsMailList)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cSENDEREMAIL, m_strSenderEmail)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cSMTPSERVER, m_strSMTPServer)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cSTOREDPROCEDURENAME, m_strStoredProcedure)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cSQL, m_strSQL)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ScheduledTask.cISHTML, m_blnBodyHTML)

            If Not intID = clsDBConstants.cintNULL Then
                colMasks.Update(m_objDB)
            Else
                intID = colMasks.Insert(m_objDB)
            End If

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

#End Region

#Region " Delete Scheduled Task "

    ''' <summary>
    ''' Deletes a scheduled task from the system
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            m_objDB.BeginTransaction()
            m_objDB.DeleteStoredProcedure(m_strStoredProcedure)
            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cSCHEDULEDTASK, ScheduledTask.ID)
            m_objDB.EndTransaction(True)

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
    End Sub

#End Region

End Class
