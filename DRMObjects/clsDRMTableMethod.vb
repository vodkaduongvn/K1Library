Public Class clsDRMTableMethod
    Inherits clsDRMBase

#Region " Members "

    Private m_intTableID As Integer
    Private m_intMethodID As Integer
    Private m_blnAudit As Boolean
    Private m_blnAuditData As Boolean
    Private m_colSecurityGroupIds As New List(Of Integer)

#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new table method
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objTable As clsTable, ByVal objMethod As clsMethod, _
                   ByVal blnAudit As Boolean, ByVal blnAuditData As Boolean, ByVal colSecGroupIDs As List(Of Integer))
        MyBase.New(objDB, objTable.DatabaseName & "." & objMethod.ExternalID, objTable.SecurityID, clsDBConstants.cintNULL)
        m_intTableID = objTable.ID
        m_intMethodID = objMethod.ID
        m_blnAudit = blnAudit
        m_blnAuditData = blnAuditData
        m_colSecurityGroupIds = colSecGroupIDs
    End Sub

    ''' <summary>
    ''' Creates a new table method
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objTable As clsTable, ByVal eMethod As clsMethod.enumMethods)
        MyBase.New(objDB, "", objTable.SecurityID, clsDBConstants.cintNULL)

        Dim objMethod As clsMethod = objDB.SysInfo.Methods(CStr(eMethod))
        m_strExternalID = objTable.DatabaseName & "." & objMethod.ExternalID
        m_intTableID = objTable.ID
        m_intMethodID = objMethod.ID
        m_blnAudit = False
        m_blnAuditData = False
    End Sub

#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRMTableMethod from an existing tablemethod database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objTableMethod As clsTableMethod, ByVal intTableID As Integer)
        MyBase.New(objDB, objTableMethod)
        m_intTableID = intTableID
        m_intMethodID = objTableMethod.MethodID
        m_blnAudit = objTableMethod.Audit
        m_blnAuditData = objTableMethod.AuditData
    End Sub

#End Region

#End Region

#Region " Properties "

    Public Property TableID() As Integer
        Get
            Return m_intTableID
        End Get
        Set(ByVal value As Integer)
            m_intTableID = value
        End Set
    End Property

    Public Property MethodID() As Integer
        Get
            Return m_intMethodID
        End Get
        Set(ByVal value As Integer)
            m_intMethodID = value
        End Set
    End Property

    Public Property Audit() As Boolean
        Get
            Return m_blnAudit
        End Get
        Set(ByVal value As Boolean)
            m_blnAudit = value
        End Set
    End Property

    Public Property AuditData() As Boolean
        Get
            Return m_blnAuditData
        End Get
        Set(ByVal value As Boolean)
            m_blnAuditData = value
        End Set
    End Property

    Public ReadOnly Property TableMethod() As clsTableMethod
        Get
            Return CType(m_objDBObj, clsTableMethod)
        End Get
    End Property

    Public Property SecurityGroupIds() As List(Of Integer)
        Get
            Return m_colSecurityGroupIds
        End Get
        Set(ByVal value As List(Of Integer))
            m_colSecurityGroupIds = value
        End Set
    End Property

#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate(ByVal objDb As clsDB_System)
        Try
            Dim colParameters As New clsDBParameterDictionary
            Dim eDirection As Data.ParameterDirection = ParameterDirection.Input
            Dim intID As Integer = clsDBConstants.cintNULL

            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cEXTERNALID, Me.ExternalID, eDirection, SqlDbType.NVarChar))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cSECURITYID, Me.SecurityID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cTABLEID, Me.TableID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cMETHODID, Me.MethodID, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cAUDIT, Me.Audit, eDirection, SqlDbType.Int))
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.TableMethod.cAUDITDATA, Me.AuditData, eDirection, SqlDbType.Int))

            If Me.ID = clsDBConstants.cintNULL Then
                colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cID, Me.ID, ParameterDirection.Output, SqlDbType.Int))
                objDb.ExecuteProcedure(clsDBConstants.Tables.cTABLEMETHOD & clsDBConstants.StoredProcedures.cINSERT, colParameters)

                m_intID = CInt(colParameters(clsDBConstants.Fields.cID).Value)
            Else
                colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cID, Me.ID, ParameterDirection.Input, SqlDbType.Int))
                objDb.ExecuteProcedure(clsDBConstants.Tables.cTABLEMETHOD & clsDBConstants.StoredProcedures.cUPDATE, colParameters)
            End If

            If m_colSecurityGroupIds IsNot Nothing AndAlso m_colSecurityGroupIds.Count >= 1 Then
                objDb.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD & "] " & _
                                 "WHERE [" & clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID & "] = " & Me.ID)

                For Each intSecID As Integer In m_colSecurityGroupIds
                    m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                                       clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID, Me.ID, _
                                       clsDBConstants.Fields.LinkSecurityGroupTableMethod.cSECURITYGROUPID, intSecID)
                Next
            ElseIf Me.ID = clsDBConstants.cintNULL Then
                If m_objDB.Profile.SecurityGroups Is Nothing AndAlso m_objDB.Profile.SecurityGroup IsNot Nothing Then
                    m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                                                       clsDBConstants.Fields.LinkSecurityGroupTableMethod.cSECURITYGROUPID, _
                                                       objDb.Profile.SecurityGroup.ID, _
                                                       clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID, Me.ID)
                Else
                    For Each intSecurityGroupID As Integer In objDb.Profile.LinkSecurityGroups.Values
                        m_objDB.InsertLink(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                                           clsDBConstants.Fields.LinkSecurityGroupTableMethod.cSECURITYGROUPID, _
                                           intSecurityGroupID, _
                                           clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID, Me.ID)
                    Next
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub Delete(ByVal objDb As clsDB_System)
        Try
            objDb.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cACCESSRIGHTMETHOD & "] " & _
                               "WHERE [" & clsDBConstants.Fields.AccessRightMethod.cTABLEMETHODID & "] = " & Me.m_intID)

            '-- Delete dependent record first
            objDb.ExecuteSQL("DELETE FROM [" & clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD & "] " & _
                               "WHERE [" & clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID & "] = " & Me.m_intID)

            '-- Delete TableMethod record
            Dim colParameters As New clsDBParameterDictionary
            colParameters.Add(New clsDBParameter(clsDBConstants.Fields.cID, Me.m_intID, ParameterDirection.Input, SqlDbType.Int))
            objDb.ExecuteProcedure(clsDBConstants.Tables.cTABLEMETHOD & clsDBConstants.StoredProcedures.cDELETE, colParameters)
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

    Protected Overrides Sub DisposeDRMObject()

    End Sub
End Class
