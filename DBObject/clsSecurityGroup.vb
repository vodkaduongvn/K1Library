Public Class clsSecurityGroup
    Inherits clsDBObjBase

#Region " Members "

    Private m_colAppMethods As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colFields As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colReadOnlyFields As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colTables As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colSecurities As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colMethods As FrameworkCollections.K1Dictionary(Of Object)
    Private m_colTypeFields As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
    Private m_colTypeMethods As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
    Private m_colTypesByTableMethod As FrameworkCollections.K1Dictionary(Of Hashtable)
    '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
    'Private m_intNominatedSecurityID As Integer
    'Private m_objNominatedSecurity As clsSecurity
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
        'm_intNominatedSecurityID = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
        'm_intNominatedSecurityID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.SecurityGroup.cNOMINATEDSECURITYID, clsDBConstants.cintNULL), Integer)
    End Sub
#End Region

#Region " Properties "


    Public ReadOnly Property LinkApplicationMethods() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colAppMethods Is Nothing Then
                m_colAppMethods = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPAPPMETHOD, _
                    clsDBConstants.Fields.LinkSecurityGroupAppMethod.cAPPMETHODID)
            End If
            Return m_colAppMethods
        End Get
    End Property

    Public ReadOnly Property LinkFields() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colFields Is Nothing Then
                m_colFields = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPFIELD, _
                    clsDBConstants.Fields.LinkSecurityGroupField.cFIELDID)
            End If
            Return m_colFields
        End Get
    End Property

    Public ReadOnly Property LinkReadOnlyFields() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colReadOnlyFields Is Nothing Then
                m_colReadOnlyFields = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPREADONLYFIELDS, _
                    clsDBConstants.Fields.LinkSecurityGroupReadOnlyField.cFIELDID)
            End If
            Return m_colReadOnlyFields
        End Get
    End Property

    Public ReadOnly Property LinkSecurities() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colSecurities Is Nothing Then
                m_colSecurities = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPSECURITY, _
                    clsDBConstants.Fields.cSECURITYID)
            End If
            Return m_colSecurities
        End Get
    End Property

    Public ReadOnly Property LinkTables() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colTables Is Nothing Then
                m_colTables = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPTABLE, _
                    clsDBConstants.Fields.LinkSecurityGroupTable.cTABLEID)
            End If
            Return m_colTables
        End Get
    End Property

    Public ReadOnly Property LinkMethods() As FrameworkCollections.K1Dictionary(Of Object)
        Get
            If m_colMethods Is Nothing Then
                m_colMethods = LoadHashTable(clsDBConstants.Tables.cLINKSECURITYGROUPTABLEMETHOD, _
                    clsDBConstants.Fields.LinkSecurityGroupTableMethod.cTABLEMETHODID)
            End If
            Return m_colMethods
        End Get
    End Property

    Public ReadOnly Property TypeFields() As FrameworkCollections.K1Dictionary(Of clsAccessRightField)
        Get
            If m_colTypeFields Is Nothing Then
                m_colTypeFields = clsAccessRightField.GetList(m_objDB, Me)
            End If
            Return m_colTypeFields
        End Get
    End Property

    Public ReadOnly Property TypeMethods() As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod)
        Get
            If m_colTypeMethods Is Nothing Then
                m_colTypeMethods = clsAccessRightMethod.GetList(m_objDB, Me)
            End If
            Return m_colTypeMethods
        End Get
    End Property

    '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
    'Public ReadOnly Property NominatedSecurityID() As Integer
    '    Get
    '        Return m_intNominatedSecurityID
    '    End Get
    'End Property

    '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
    'Public ReadOnly Property NominatedSecurity() As clsSecurity
    '    Get
    '        If m_objNominatedSecurity Is Nothing Then
    '            m_objNominatedSecurity = clsSecurity.GetItem(m_intNominatedSecurityID, m_objDB)
    '        End If
    '        Return m_objNominatedSecurity
    '    End Get
    'End Property

    'Public ReadOnly Property TypesByTableMethod() As FrameworkCollections.K1Dictionary(Of Hashtable)
    '    Get
    '        If m_colTypesByTableMethod Is Nothing Then
    '            AssignTypesByTableMethod()
    '        End If
    '        Return m_colTypesByTableMethod
    '    End Get
    'End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsSecurityGroup
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cSECURITYGROUP, intID)

            Return New clsSecurityGroup(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsSecurityGroup)
        Dim colSecurityGroups As FrameworkCollections.K1Dictionary(Of clsSecurityGroup)
        Dim objSecurityGroup As clsSecurityGroup

        Try
            Dim strSP As String = clsDBConstants.Tables.cSECURITYGROUP & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colSecurityGroups = New FrameworkCollections.K1Dictionary(Of clsSecurityGroup)
            For Each objDR As DataRow In objDT.Rows
                objSecurityGroup = New clsSecurityGroup(objDR, objDB)
                colSecurityGroups.Add(objSecurityGroup.ID.ToString(), objSecurityGroup)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colSecurityGroups
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Methods "

    Private Function LoadHashTable(ByVal strTable As String, ByVal strField As String) As FrameworkCollections.K1Dictionary(Of Object)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of Object)

        Try
            Dim objDT As DataTable = m_objDB.GetDataTableBySQL("SELECT [" & strField & "] AS [X] " & _
                "FROM [" & strTable & "] WHERE [" & _
                clsDBConstants.Fields.LinkSecurityGroupField.cSECURITYGROUPID & "] = " & m_intID)

            colObjects = New FrameworkCollections.K1Dictionary(Of Object)
            For Each objDR As DataRow In objDT.Rows
                colObjects.Add(CStr(objDR(0)), CInt(objDR(0)))
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function

    'Private Sub AssignTypesByTableMethod()
    '    m_colTypesByTableMethod = New FrameworkCollections.K1Dictionary(Of Hashtable)
    '    Dim colTypeMethods As FrameworkCollections.K1Dictionary(Of clsAccessRightMethod) = TypeMethods

    '    For Each objAccessMethod As clsAccessRightMethod In colTypeMethods.Values
    '        Dim colIDs As Hashtable

    '        colIDs = m_colTypesByTableMethod(CStr(objAccessMethod.TableMethodID))

    '        If colIDs Is Nothing Then
    '            colIDs = New Hashtable
    '            m_colTypesByTableMethod.Add(CStr(objAccessMethod.TableMethodID), colIDs)
    '        End If

    '        If colIDs(CStr(objAccessMethod.AppliesToTypeID)) Is Nothing Then
    '            colIDs.Add(CStr(objAccessMethod.AppliesToTypeID), objAccessMethod.AppliesToTypeID)
    '        End If
    '    Next
    'End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If m_colFields IsNot Nothing Then
                m_colFields.Dispose()
                m_colFields = Nothing
            End If

            If m_colTables IsNot Nothing Then
                m_colTables.Dispose()
                m_colTables = Nothing
            End If

            If m_colSecurities IsNot Nothing Then
                m_colSecurities.Dispose()
                m_colSecurities = Nothing
            End If

            If m_colMethods IsNot Nothing Then
                m_colMethods.Dispose()
                m_colMethods = Nothing
            End If

            '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
            'If m_objNominatedSecurity IsNot Nothing Then
            '    m_objNominatedSecurity.Dispose()
            '    m_objNominatedSecurity = Nothing
            'End If

            If m_colTypeFields IsNot Nothing Then
                m_colTypeFields.Dispose()
                m_colTypeFields = Nothing
            End If

            If m_colTypeMethods IsNot Nothing Then
                m_colTypeMethods.Dispose()
                m_colTypeMethods = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
