Public Class clsDRMApplicationMethod
    Inherits clsDRMBase

#Region " Members "

    Private m_strCaption As String
    Private m_intStringID As Integer
    Private m_eMethodType As clsApplicationMethod.enumMethodType
    Private m_intSortOrder As Integer
    Private m_intTableID As Integer
    Private m_intAppliesToTypeID As Integer
    Private m_strFile As String
    Private m_intEDOCID As Integer
    Private m_intUIID As Integer
    Private m_intParentAppMethodID As Integer
#End Region

#Region " Constructors "

#Region " New "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer, _
    ByVal strCaption As String, _
    ByVal eMethodType As clsApplicationMethod.enumMethodType, _
    ByVal intTableID As Integer, _
    ByVal intSortOrder As Integer, _
    ByVal intUIID As Integer, _
    ByVal intParentAppMethodID As Integer, _
    Optional ByVal intAppliesToTypeID As Integer = clsDBConstants.cintNULL, _
    Optional ByVal strFile As String = Nothing)
        MyBase.New(objDB, strExternalID, intSecurityID, intTypeID)
        m_strCaption = strCaption
        m_eMethodType = eMethodType
        m_intTableID = intTableID
        m_intSortOrder = intSortOrder
        m_intAppliesToTypeID = intAppliesToTypeID
        m_strFile = strFile
        m_intEDOCID = clsDBConstants.cintNULL
        m_intUIID = intUIID
        m_intParentAppMethodID = intParentAppMethodID
    End Sub
#End Region

#Region " From Existing "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal objAppMethod As clsApplicationMethod)
        MyBase.New(objDB, objAppMethod)
        If Not objAppMethod.StringObj Is Nothing Then
            m_intStringID = objAppMethod.StringObj.ID
            m_strCaption = objAppMethod.StringObj.GetLanguageString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
        Else
            m_strCaption = objAppMethod.ExternalID
        End If
        If Not objAppMethod.EDOC Is Nothing Then
            m_intEDOCID = objAppMethod.EDOC.ID
        Else
            m_intEDOCID = clsDBConstants.cintNULL
        End If
        m_intTableID = objAppMethod.TableID
        m_eMethodType = objAppMethod.MethodType
        m_intSortOrder = objAppMethod.SortOrder
        m_intAppliesToTypeID = objAppMethod.AppliesToTypeID
        m_intUIID = objAppMethod.UIID
        m_intParentAppMethodID = objAppMethod.ParentApplicationMethodID
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property ApplicationMethod() As clsApplicationMethod
        Get
            Return CType(m_objDBObj, clsApplicationMethod)
        End Get
    End Property

    Public Property IconFile() As String
        Get
            Return m_strFile
        End Get
        Set(ByVal value As String)
            m_strFile = value
        End Set
    End Property

    Public Property Caption() As String
        Get
            Return m_strCaption
        End Get
        Set(ByVal value As String)
            m_strCaption = value
        End Set
    End Property

    Public Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
        Set(ByVal value As Integer)
            m_intSortOrder = value
        End Set
    End Property

    Public Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
        Set(ByVal value As Integer)
            m_intAppliesToTypeID = value
        End Set
    End Property

    Public Property ParentAppMethodID() As Integer
        Get
            Return m_intParentAppMethodID
        End Get
        Set(ByVal value As Integer)
            m_intParentAppMethodID = value
        End Set
    End Property

    Public Property MethodType() As clsApplicationMethod.enumMethodType
        Get
            Return m_eMethodType
        End Get
        Set(ByVal value As clsApplicationMethod.enumMethodType)
            m_eMethodType = value
        End Set
    End Property

#End Region

#Region " Methods "

    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False
        Dim objEDOC As clsEDOC

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            '================================================================================
            'Create the application method's language string
            '================================================================================
            Dim objStringObj As clsString = Nothing
            Dim intLanguageStringID As Integer = clsDBConstants.cintNULL
            Dim objLanguageString As clsLanguageString

            If ApplicationMethod IsNot Nothing Then
                objStringObj = ApplicationMethod.StringObj

                If objStringObj IsNot Nothing Then
                    objLanguageString = ApplicationMethod.StringObj.LinkLanguageStrings( _
                        CType(m_objDB.Profile.LanguageID, String))
                    intLanguageStringID = objLanguageString.ID
                End If
            End If

            If objStringObj Is Nothing Then
                objStringObj = New clsString(m_objDB, clsDBConstants.cintNULL, m_strExternalID, m_intSecurityID)
                objStringObj.InsertUpdate()
            End If

            '================================================================================
            'Create the icon and icon edoc
            '================================================================================
            If Not (m_strFile Is Nothing OrElse m_strFile.Trim.Length = 0) Then
                objEDOC = New clsEDOC(m_objDB, m_intEDOCID, "ApplicationMethod - " & m_strExternalID, _
                    m_intSecurityID, m_strFile)
                objEDOC.InsertUpdate()
                m_intEDOCID = objEDOC.ID

                m_objDB.WriteBLOB(clsDBConstants.Tables.cEDOC, clsDBConstants.Fields.EDOC.cIMAGE, _
                    SqlDbType.Image, objEDOC.Size, objEDOC.ID, m_strFile)
            End If

            '================================================================================
            'Create the application method
            '================================================================================
            Dim objTable As clsTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cAPPLICATIONMETHOD)

            Dim intID As Integer = clsDBConstants.cintNULL
            If Not ApplicationMethod Is Nothing Then
                intID = ApplicationMethod.ID
            End If

            Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(objTable, intID)

            colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cMETHODTYPE, CInt(m_eMethodType))
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cSORTORDER, m_intSortOrder)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cTABLEID, m_intTableID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cEDOCID, m_intEDOCID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cSTRINGID, objStringObj.ID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cUIID, m_intUIID)

            Dim objParentAppMethod As clsApplicationMethod
            Select Case m_eMethodType
                Case clsApplicationMethod.enumMethodType.MAINTENANCE
                    objParentAppMethod = m_objDB.SysInfo.ApplicationMethods(clsApplicationMethod.enumAppMethod.Maintenance)
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cPARENTAPPMETHODID, objParentAppMethod.ID)
                Case clsApplicationMethod.enumMethodType.METADATA_SEARCH
                    objParentAppMethod = m_objDB.SysInfo.ApplicationMethods(clsApplicationMethod.enumAppMethod.Search_Metadata)
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cPARENTAPPMETHODID, objParentAppMethod.ID)
                Case clsApplicationMethod.enumMethodType.BOOLEAN_SEARCH
                    objParentAppMethod = m_objDB.SysInfo.ApplicationMethods(clsApplicationMethod.enumAppMethod.Search_Boolean)
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cPARENTAPPMETHODID, objParentAppMethod.ID)
                Case Else
                    colMasks.UpdateMaskObj(clsDBConstants.Fields.ApplicationMethod.cPARENTAPPMETHODID, m_intParentAppMethodID)
            End Select

            If ApplicationMethod Is Nothing Then
                intID = colMasks.Insert(m_objDB)
            Else
                colMasks.Update(m_objDB)
            End If

            objLanguageString = New clsLanguageString(m_objDB, _
                intLanguageStringID, m_strExternalID & " (" & m_objDB.Profile.Language.ExternalID & ")", _
                m_intSecurityID, m_objDB.Profile.LanguageID, objStringObj.ID, m_strCaption)

            objLanguageString.InsertUpdate()
            objStringObj.LinkLanguageStrings(CStr(objLanguageString.LanguageID)) = objLanguageString

            m_objDBObj = clsApplicationMethod.GetItem(intID, m_objDB)
            m_intID = m_objDBObj.ID

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cAPPLICATIONMETHOD, ApplicationMethod.ID)

            If ApplicationMethod.StringObj IsNot Nothing Then
                SystemDB.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLANGUAGESTRING & "] " & _
                    "WHERE [" & clsDBConstants.Fields.LanguageString.cSTRINGID & "] = " & ApplicationMethod.StringObj.ID)
                SystemDB.DeleteRecord(clsDBConstants.Tables.cSTRING, ApplicationMethod.StringObj.ID)
            End If

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()

    End Sub
#End Region

End Class
