#Region " File Information "
'=====================================================================
' This class represents the table Security in the Database.
'=====================================================================
#End Region


Public Class clsDRMFunction
    Inherits clsDBObjBase


#Region " Members "
    Dim m_intUIID As Integer
#End Region

#Region " Properties "
    Public ReadOnly Property UIID() As Integer
        Get
            Return m_intUIID
        End Get
    End Property

#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub


    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cUIID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsDRMFunction
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cDRMFunction, intID)

            Return New clsDRMFunction(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsDRMFunction)
        Dim colDRMFunctions As FrameworkCollections.K1Dictionary(Of clsDRMFunction)

        Try
            Dim strSP As String = clsDBConstants.Tables.cDRMFunction & _
                clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colDRMFunctions = New FrameworkCollections.K1Dictionary(Of clsDRMFunction)
            For Each objDR As DataRow In objDT.Rows
                Dim objDRMFunction As New clsDRMFunction(objDR, objDB)
                colDRMFunctions.Add(CStr(objDRMFunction.UIID), objDRMFunction)
            Next

            Return colDRMFunctions
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Enumerations "
    Public Enum enumDRMFunction
        Configure_Tables_and_Fields = 1
        ConfigureLink_Tables = 2
        Configure_Security_Codes = 3
        Configure_Security_Groups = 4
        Configure_Users = 5
        Configure_Error_Messages = 6
        Configure_Warning_Messages = 7
        Configure_Triggers = 8
        Configure_Filters = 9
        Configure_Scheduled_Tasks = 10
        Configure_Language_Settings = 11
        Configure_MetadataProfile_Type_Associtations = 12
        Configure_Auto_Numbers = 13
        Configure_Processes = 14
        Backup_The_Database = 15
        ReIndex_The_Database = 16
        Synchronize_With_Active_Directory = 17
        Change_Configuration_Settings = 18
        Change_Calendar_Settings = 19
        Configure_Audit_Trail_Settings = 20
        Manage_Licences = 21
        Configure_ToolBar = 22
        Configure_Scheduled_Reports = 23
        Configure_EDOC_Archive = 24
    End Enum
#End Region


End Class

