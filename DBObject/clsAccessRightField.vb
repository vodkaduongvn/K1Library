Public Class clsAccessRightField
    Inherits clsDBObjBase

#Region " Members "

    Private m_intSecurityGroupID As Integer
    Private m_intFieldID As Integer
    Private m_intAppliesToTypeID As Integer
    Private m_blnVisible As Boolean
    Private m_blnReadOnly As Boolean
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intSecurityGroupID As Integer, _
    ByVal intFieldID As Integer, _
    ByVal intAppliesToTypeID As Integer, _
    ByVal blnVisible As Boolean, _
    ByVal blnReadOnly As Boolean)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intSecurityGroupID = intSecurityGroupID
        m_intFieldID = intFieldID
        m_intAppliesToTypeID = intAppliesToTypeID
        m_blnVisible = blnVisible
        m_blnReadOnly = blnReadOnly
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intSecurityGroupID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightField.cSECURITYGROUPID, clsDBConstants.cintNULL), Integer)
        m_intFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightField.cFIELDID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightField.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
        m_blnVisible = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightField.cISVISIBLE, False), Boolean)
        m_blnReadOnly = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightField.cISREADONLY, False), Boolean)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property FieldID() As Integer
        Get
            Return m_intFieldID
        End Get
    End Property

    Public ReadOnly Property Field() As clsField
        Get
            Return m_objDB.SysInfo.Fields(m_intFieldID)
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property

    Public ReadOnly Property SecurityGroupID() As Integer
        Get
            Return m_intSecurityGroupID
        End Get
    End Property

    Public Property Visible() As Boolean
        Get
            Return m_blnVisible
        End Get
        Set(value As Boolean)
            m_blnVisible = value
        End Set
    End Property

    Public Property IsReadOnly() As Boolean
        Get
            Return m_blnReadOnly
        End Get
        Set(value As Boolean)
            m_blnReadOnly = value
        End Set
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsAccessRightField
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cACCESSRIGHTFIELD, intID)

            Return New clsAccessRightField(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal objSG As clsSecurityGroup) As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Dim colItems As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Dim objItem As clsAccessRightField

        Try
            Dim strSP As String = clsDBConstants.Tables.cACCESSRIGHTFIELD & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTableByField(clsDBConstants.Tables.cACCESSRIGHTFIELD, _
                clsDBConstants.Fields.AccessRightField.cSECURITYGROUPID, objSG.ID)

            colItems = New FrameworkCollections.K1Dictionary(Of clsAccessRightField)
            For Each objDR As DataRow In objDT.Rows
                objItem = New clsAccessRightField(objDR, objDB)
                colItems.Add(objItem.FieldID & "_" & objItem.AppliesToTypeID, objItem)
            Next

            Return colItems
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal objUP As clsUserProfile) As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Dim colItems As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Dim objItem As clsAccessRightField

        Try
            Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT AccessRightField.* " & _
                "FROM AccessRightField INNER JOIN " & _
                "Field ON AccessRightField.FieldID = Field.ID INNER JOIN " & _
                "LinkUserProfileSecurityGroup ON AccessRightField.SecurityGroupID = LinkUserProfileSecurityGroup.SecurityGroupID INNER JOIN " & _
                "LinkSecurityGroupTable ON Field.TableID = LinkSecurityGroupTable.TableID AND " & _
                "LinkUserProfileSecurityGroup.SecurityGroupID = LinkSecurityGroupTable.SecurityGroupID " & _
                "WHERE (LinkUserProfileSecurityGroup.UserProfileID = " & objUP.ID & ")")

            colItems = New FrameworkCollections.K1Dictionary(Of clsAccessRightField)
            For Each objDR As DataRow In objDT.Rows
                objItem = New clsAccessRightField(objDR, objDB)
                '2016-09-22 -- Peter & James -- Bug fix for #1600003206
                If Not colItems.ContainsKey(objItem.FieldID & "_" & objItem.AppliesToTypeID & "_" & objItem.SecurityGroupID) Then
                    colItems.Add(objItem.FieldID & "_" & objItem.AppliesToTypeID & "_" & objItem.SecurityGroupID, objItem)
                End If
            Next

            Return colItems
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cACCESSRIGHTFIELD), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightField.cFIELDID, m_intFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightField.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightField.cSECURITYGROUPID, m_intSecurityGroupID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightField.cISVISIBLE, m_blnVisible)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightField.cISREADONLY, m_blnReadOnly)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
