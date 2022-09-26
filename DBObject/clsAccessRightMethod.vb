Public Class clsAccessRightMethod
    Inherits clsDBObjBase

#Region " Members "

    Private m_intSecurityGroupID As Integer
    Private m_intTableMethodID As Integer
    Private m_intAppliesToTypeID As Integer
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
    ByVal intTableMethodID As Integer, _
    ByVal intAppliesToTypeID As Integer)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intSecurityGroupID = intSecurityGroupID
        m_intTableMethodID = intTableMethodID
        m_intAppliesToTypeID = intAppliesToTypeID
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intSecurityGroupID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightMethod.cSECURITYGROUPID, clsDBConstants.cintNULL), Integer)
        m_intTableMethodID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AccessRightMethod.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
        End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property TableMethodID() As Integer
        Get
            Return m_intTableMethodID
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
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsAccessRightMethod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cACCESSRIGHTMETHOD, intID)

            Return New clsAccessRightMethod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal objSG As clsSecurityGroup) As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Dim colItems As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Dim objItem As clsAccessRightMethod

        Try
            Dim objDT As DataTable = objDB.GetDataTableByField(clsDBConstants.Tables.cACCESSRIGHTMETHOD, _
                clsDBConstants.Fields.AccessRightField.cSECURITYGROUPID, objSG.ID)

            colItems = New FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
            For Each objDR As DataRow In objDT.Rows
                objItem = New clsAccessRightMethod(objDR, objDB)
                colItems.Add(objItem.TableMethodID & "_" & objItem.AppliesToTypeID, objItem)
            Next

            Return colItems
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB, ByVal objUP As clsUserProfile) As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Dim colItems As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Dim objItem As clsAccessRightMethod

        Try
            '2016-09-23 -- Peter Melisi -- Bug fix for #1600003208
            Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT [" & clsDBConstants.Fields.cEXTERNALID & "], [" & clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID & "], [" & clsDBConstants.Fields.AccessRightMethod.cAPPLIESTOTYPEID & "]" & _
            "FROM [" & clsDBConstants.Tables.cACCESSRIGHTMETHOD & "] INNER JOIN [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "] " & _
            "ON [" & clsDBConstants.Tables.cACCESSRIGHTMETHOD & "].[" & clsDBConstants.Fields.AccessRightMethod.cSECURITYGROUPID & "] = [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "].[" & clsDBConstants.Fields.LinkUserProfileSecurityGroup.cSECURITYGROUPID & "] " & _
            "WHERE [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "].[" & clsDBConstants.Fields.LinkUserProfileSecurityGroup.cUSERPROFILEID & "] = " & objUP.ID & " " & _
            "GROUP BY [" & clsDBConstants.Fields.cEXTERNALID & "], [" & clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID & "], [" & clsDBConstants.Fields.AccessRightMethod.cAPPLIESTOTYPEID & "] " & _
            "HAVING COUNT([" & clsDBConstants.Tables.cACCESSRIGHTMETHOD & "].[" & clsDBConstants.Fields.cID & "]) = " & _
            "(SELECT COUNT([" & clsDBConstants.Fields.cID & "]) FROM [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "] " & _
            "WHERE [" & clsDBConstants.Tables.cLINKUSERPROFILESECURITYGROUP & "].[" & clsDBConstants.Fields.LinkUserProfileSecurityGroup.cUSERPROFILEID & "] = " & objUP.ID & ")")

            colItems = New FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)

            For Each objDR As DataRow In objDT.Rows
                objItem = New clsAccessRightMethod(objDR, objDB)
                If Not colItems.ContainsKey(objItem.TableMethodID & "_" & objItem.AppliesToTypeID) Then
                    colItems.Add(objItem.TableMethodID & "_" & objItem.AppliesToTypeID, objItem)
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
             m_objDB.SysInfo.Tables(clsDBConstants.Tables.cACCESSRIGHTMETHOD), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID, m_intTableMethodID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightMethod.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AccessRightMethod.cSECURITYGROUPID, m_intSecurityGroupID)

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
