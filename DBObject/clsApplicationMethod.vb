Public Class clsApplicationMethod
    Inherits clsDBObjBase

#Region " Members "

    Private m_intUIID As Integer
    Private m_intEDOCID As Integer
    Private m_objEDOC As clsEDOC
    Private m_intStringID As Integer
    Private m_objString As clsString
    Private m_intParentAppMethodID As Integer
    Private m_intSortOrder As Integer
    Private m_eMethodType As enumMethodType
    Private m_intTableID As Integer
    Private m_intAppliesToTypeID As Integer
#End Region

#Region " Enumerations "

    Public Enum enumMethodType
        OTHER = 0
        MAINTENANCE = 1
        METADATA_SEARCH = 2
        BOOLEAN_SEARCH = 3
    End Enum

    Public Enum enumAppMethod
        Search = 1
        Pick_List = 2
        Maintenance = 4
        Barcode_PortableReader = 5
        WorkFlow = 6
        Code = 7
        Audit = 8
        People = 9
        User_Profile = 10
        Object_Numbering = 11
        File_Plan = 12
        Search_Text = 13
        Search_External_ID = 14
        Search_Metadata = 15
        Search_Boolean = 16
        Search_Saved_Search = 17
        Search_Any_Table = 18
        Metadata_File_Folder_Profile = 19
        Metadata_Document_Profile = 20
        Metadata_Attachment = 21
        Metadata_Archive_Box = 22
        Metadata_Storage_Space = 23
        Metadata_Movement = 24
        Metadata_Resubmit = 25
        Boolean_File_Folder_Profile = 26
        Boolean_Document_Profile = 27
        Boolean_Attachment = 28
        Boolean_Archive_Box = 29
        Boolean_Storage_Space = 30
        Boolean_Movement = 31
        Boolean_Resubmit = 32
        Maintenance_File_Folder_Profile = 33
        Maintenance_Document_Profile = 34
        Maintenance_Attachment = 35
        Maintenance_Archive_Box = 36
        Maintenance_Storage_Space = 37
        Workflow_Workflow = 38
        Workflow_Task = 39
        Workflow_To_Do_List = 40
        Codes_File_Folder_Profile = 41
        Codes_Document_Profile = 42
        Codes_Archive_Box = 43
        Codes_Organization = 44
        Codes_Vital_Record = 45
        Codes_Media_Type = 46
        Codes_Attachment_Format = 47
        Codes_Supplemental_Markup = 48
        CodesFile_Type_Code = 49
        CodesFile_Disposition = 50
        CodesFile_User_Code_1 = 51
        CodesFile_User_Code_2 = 52
        CodesFile_User_Code_3 = 53
        CodesDoc_Document_Type = 54
        CodesDoc_Document_Status = 55
        CodesDoc_Author_Type = 56
        CodesDoc_User_Code_1 = 57
        CodesDoc_User_Code_2 = 58
        CodesDoc_User_Code_3 = 59
        CodesBox_Box_Type = 60
        CodesBox_Box_Status = 61
        CodesOrg_Organization = 62
        CodesOrg_Location = 63
        CodesOrg_Department = 64
        AN_Auto_Number_Format = 65
        AN_Multiple_Sequence = 66
        FilePlan_Record_Category = 67
        FilePlan_Retention = 68
        FilePlan_File_Titles = 69
        FilePlan_Corporate_Vocabulary = 70
        Maintenance_Movement = 71
        Maintenance_Resubmit = 72
        FilePlan_Series = 73
        Barcode = 74
        Barcode_BarcodeMovements = 75
        Pick_List_Hold = 76
        Pick_List_Remove = 77
        Pick_List_Active = 78
        Pick_List_Close = 79
        Pick_List_Destroy = 80
        Pick_List_Vital_Record = 81
        Pick_List_To_Box = 82
        Pick_List_To_Storage = 83
        AN_Modify_Auto_Number = 84
        Workflow_Designer = 85
        LegalHold = 86
        Archive = 87
    End Enum
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cUIID, clsDBConstants.cintNULL), Integer)
        m_intEDOCID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cEDOCID, clsDBConstants.cintNULL), Integer)
        m_intStringID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cSTRINGID, clsDBConstants.cintNULL), Integer)
        m_intParentAppMethodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cPARENTAPPMETHODID, clsDBConstants.cintNULL), Integer)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cSORTORDER, clsDBConstants.cintNULL), Integer)
        m_eMethodType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cMETHODTYPE, enumMethodType.OTHER), enumMethodType)
        m_intTableID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cTABLEID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.ApplicationMethod.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property UIID() As Integer
        Get
            Return m_intUIID
        End Get
    End Property

    Public ReadOnly Property EDOC() As clsEDOC
        Get
            If m_objEDOC Is Nothing Then
                If Not m_intEDOCID = clsDBConstants.cintNULL Then
                    m_objEDOC = clsEDOC.GetItem(m_intEDOCID, Me.Database)
                End If
            End If
            Return m_objEDOC
        End Get
    End Property

    Public ReadOnly Property ParentApplicationMethodID() As Integer
        Get
            Return m_intParentAppMethodID
        End Get
    End Property

    Public ReadOnly Property StringObj() As clsString
        Get
            If m_objString Is Nothing Then
                If Not m_intStringID = clsDBConstants.cintNULL Then
                    m_objString = clsString.GetItem(m_intStringID, Me.Database)
                End If
            End If
            Return m_objString
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property

    Public ReadOnly Property MethodType() As enumMethodType
        Get
            Return m_eMethodType
        End Get
    End Property

    Public ReadOnly Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
    End Property

    Public ReadOnly Property Table() As clsTable
        Get
            Return m_objDB.SysInfo.Tables(m_intTableID)
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsApplicationMethod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cAPPLICATIONMETHOD, intID)

            Return New clsApplicationMethod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal objK1Config As clsK1Configuration) As FrameworkCollections.K1DualKeyDictionary(Of clsApplicationMethod, Integer)
        Dim colMethods As FrameworkCollections.K1DualKeyDictionary(Of clsApplicationMethod, Integer)
        Dim objMethod As clsApplicationMethod

        Try
            Dim strSP As String = clsDBConstants.Tables.cAPPLICATIONMETHOD & clsDBConstants.StoredProcedures.cGETLIST
            Dim colParams As New clsDBParameterDictionary

            If objK1Config.DbVersion >= 11.04 Then
                colParams.Add(New clsDBParameter("@MethodType", 0))
            Else
                colParams = Nothing
            End If

            Dim objDT As DataTable = objDB.GetDataTable(strSP, colParams)

            colMethods = New FrameworkCollections.K1DualKeyDictionary(Of clsApplicationMethod, Integer)
            For Each objDR As DataRow In objDT.Rows
                objMethod = New clsApplicationMethod(objDR, objDB)
                colMethods.Add(CStr(objMethod.UIID), objMethod)
            Next

            Return colMethods
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal eMethodType As enumMethodType) As FrameworkCollections.K1Collection(Of clsApplicationMethod)
        Dim colMethods As FrameworkCollections.K1Collection(Of clsApplicationMethod)
        Dim objMethod As clsApplicationMethod

        If objDB.SysInfo.K1Configuration.DbVersion < 11.04 Then
            Return New FrameworkCollections.K1Collection(Of clsApplicationMethod)
        End If

        Try
            Dim strSP As String = clsDBConstants.Tables.cAPPLICATIONMETHOD & clsDBConstants.StoredProcedures.cGETLIST
            Dim colParams As New clsDBParameterDictionary

            colParams.Add(New clsDBParameter("@MethodType", CInt(eMethodType)))

            Dim objDT As DataTable = objDB.GetDataTable(strSP, colParams)
            objDT.DefaultView.Sort = clsDBConstants.Fields.ApplicationMethod.cSORTORDER

            colMethods = New FrameworkCollections.K1Collection(Of clsApplicationMethod)
            For intLoop As Integer = 0 To objDT.DefaultView.Count - 1
                Dim objDR As DataRow = objDT.DefaultView(intLoop).Row
                objMethod = New clsApplicationMethod(objDR, objDB)
                colMethods.Add(objMethod)
            Next

            Return colMethods
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_objEDOC IsNot Nothing Then
                m_objEDOC.Dispose()
                m_objEDOC = Nothing
            End If

            If m_objString IsNot Nothing Then
                m_objString.Dispose()
                m_objString = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
