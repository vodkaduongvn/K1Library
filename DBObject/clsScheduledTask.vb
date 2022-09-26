Public Class clsScheduledTask
    Inherits clsDBObjBase

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
    Private m_intMailListID As Integer
    Private m_intSavedReportID As Integer
    Private m_intSavedReportFormat As Integer
    Private m_intSavedSearchID As Integer
    Private m_intReportTableID As Integer = clsDBConstants.cintNULL
    Private m_blnIsHTML As Boolean
    Private m_blnSendIfEmpty As Boolean

#End Region

#Region " Constructors "

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_blnActive = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cISACTIVE, False), Boolean)
        m_dtBeginDate = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cBEGINDATE, Date.MinValue), Date)
        m_dtEndDate = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cENDDATE, Date.MinValue), Date)
        m_dtTimeDue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cTIMEDUE, Date.MinValue), Date)
        m_intPeriodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cPERIODID, clsDBConstants.cintNULL), Integer)
        m_blnReturnsMailList = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cRETURNSMAILLIST, False), Boolean)
        m_strSenderEmail = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSENDEREMAIL, clsDBConstants.cstrNULL), String)
        m_strSMTPServer = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSMTPSERVER, clsDBConstants.cstrNULL), String)
        m_strStoredProcedure = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSTOREDPROCEDURENAME, clsDBConstants.cstrNULL), String)
        m_strSQL = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSQL, clsDBConstants.cstrNULL), String)
        m_intMailListID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cMAILLISTID, clsDBConstants.cintNULL), Integer)
        m_intSavedReportID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSAVEDREPORTID, clsDBConstants.cintNULL), Integer)
        m_intSavedReportFormat = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSAVEDREPORTFORMAT, clsDBConstants.cintNULL), Integer)
        m_intSavedSearchID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSAVEDSEARCHID, clsDBConstants.cintNULL), Integer)
        m_blnIsHTML = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cISHTML, False), Boolean)
        m_blnSendIfEmpty = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ScheduledTask.cSENDIFEMPTY, False), Boolean)
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property IsActive() As Boolean
        Get
            Return m_blnActive
        End Get
    End Property

    Public ReadOnly Property BeginDate() As Date
        Get
            Return m_dtBeginDate
        End Get
    End Property

    Public ReadOnly Property EndDate() As Date
        Get
            Return m_dtEndDate
        End Get
    End Property

    Public ReadOnly Property TimeDue() As Date
        Get
            Return m_dtTimeDue
        End Get
    End Property

    Public ReadOnly Property PeriodID() As Integer
        Get
            Return m_intPeriodID
        End Get
    End Property

    Public ReadOnly Property ReturnsMailList() As Boolean
        Get
            Return m_blnReturnsMailList
        End Get
    End Property

    Public ReadOnly Property SenderEmail() As String
        Get
            Return m_strSenderEmail
        End Get
    End Property

    Public ReadOnly Property SMTPServer() As String
        Get
            Return m_strSMTPServer
        End Get
    End Property

    Public ReadOnly Property StoredProcedure() As String
        Get
            Return m_strStoredProcedure
        End Get
    End Property

    Public ReadOnly Property SQL() As String
        Get
            Return m_strSQL
        End Get
    End Property

    Public ReadOnly Property SavedReportID() As Integer
        Get
            Return m_intSavedReportID
        End Get
    End Property

    Public ReadOnly Property SavedReportFormat() As Integer
        Get
            Return m_intSavedReportFormat
        End Get
    End Property

    Public ReadOnly Property SavedSearchID() As Integer
        Get
            Return m_intSavedSearchID
        End Get
    End Property

    Public ReadOnly Property MailListID() As Integer
        Get
            Return m_intMailListID
        End Get
    End Property

    Public ReadOnly Property ReportTableID() As Integer
        Get
            If m_intReportTableID = clsDBConstants.cintNULL AndAlso _
            Not m_intSavedReportID = clsDBConstants.cintNULL Then
                Dim objDT As DataTable = m_objDB.GetItem(clsDBConstants.Tables.cSAVEDREPORT, m_intSavedReportID)

                If objDT.Rows.Count = 1 Then
                    Return CInt(clsDB.NullValue(objDT.Rows(0)(clsDBConstants.Fields.SavedReport.cTABLEID), clsDBConstants.cintNULL))
                Else
                    Return m_intReportTableID
                End If
            Else
                Return m_intReportTableID
            End If
        End Get
    End Property

    Public ReadOnly Property IsHTML() As Boolean
        Get
            Return m_blnIsHTML
        End Get
    End Property

    Public ReadOnly Property SendIfEmpty() As Boolean
        Get
            Return m_blnSendIfEmpty
        End Get
    End Property

#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsScheduledTask
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cSCHEDULEDTASK, intID)

            Return New clsScheduledTask(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsScheduledTask)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsScheduledTask)

        Try
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cSCHEDULEDTASK)
            Dim objDT As DataTable = objDB.GetDataTable(objTable)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsScheduledTask)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsScheduledTask(objDR, objDB)
                If objItem.SavedReportID = clsDBConstants.cintNULL Then
                    colObjects.Add(CStr(objItem.StoredProcedure), objItem)
                Else
                    colObjects.Add(CStr(objItem.ID), objItem)
                End If
            Next

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
