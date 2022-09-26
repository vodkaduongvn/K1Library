#Region " File Information "

'=====================================================================
' This class represents the table Table in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      6/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsTable
    Inherits clsDBObjBase

#Region " Members "

    Private m_eClass As clsDBConstants.enumTableClass
    Private m_intIconID As Integer
    Private m_objIcon As clsIcon
    Private m_colFields As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_colTableMethod As FrameworkCollections.K1Dictionary(Of clsTableMethod)
    Private m_colFieldLinks As FrameworkCollections.K1Dictionary(Of clsFieldLink)
    Private m_colOneToManyLinks As FrameworkCollections.K1Dictionary(Of clsFieldLink)
    Private m_colManyToManyLinks As FrameworkCollections.K1Dictionary(Of clsFieldLink)
    Private m_colProcesses As FrameworkCollections.K1Dictionary(Of clsProcess)
    Private m_colWarnings As FrameworkCollections.K1Collection(Of clsWarningMessage)
    Private m_strDatabaseName As String
    Private m_blnTypeDependent As Boolean
    Private m_blnShowIcon As Boolean
    Private m_colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn)
    Private m_colTypeListColumns As FrameworkCollections.K1Dictionary(Of FrameworkCollections.K1Dictionary(Of clsListColumn))
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal intID As Integer, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal strDatabaseName As String, _
    ByVal eClass As clsDBConstants.enumTableClass, _
    ByVal intIconID As Integer, _
    ByVal intCaptionID As Integer, _
    ByVal blnTypeDependent As Boolean, _
    ByVal blnShowIcon As Boolean)
        MyBase.New(objDB, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_strDatabaseName = strDatabaseName
        m_eClass = eClass
        m_intIconID = intIconID
        m_intCaptionID = intCaptionID
        m_blnTypeDependent = blnTypeDependent
        m_blnShowIcon = blnShowIcon
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_eClass = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cCLASS, clsDBConstants.cintNULL), clsDBConstants.enumTableClass)
        m_intIconID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cICONID, clsDBConstants.cintNULL), Integer)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_strDatabaseName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cDATABASENAME, clsDBConstants.cstrNULL), String)
        m_blnTypeDependent = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cISTYPEDEPENDENT, False), Boolean)
        m_blnShowIcon = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Table.cSHOWICON, True), Boolean)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property ClassType() As clsDBConstants.enumTableClass
        Get
            Return m_eClass
        End Get
    End Property

    Public Property Icon() As clsIcon
        Get
            If m_objIcon Is Nothing Then
                If Not m_intIconID = clsDBConstants.cintNULL Then
                    m_objIcon = clsIcon.GetItem(m_intIconID, Me.Database)
                End If
            End If
            Return m_objIcon
        End Get
        Set(ByVal value As clsIcon)
            m_objIcon = value
        End Set
    End Property

    Public ReadOnly Property IconID() As Integer
        Get
            Return m_intIconID
        End Get
    End Property

    Public ReadOnly Property CaptionID() As Integer
        Get
            Return m_intCaptionID
        End Get
    End Property

    Public ReadOnly Property CaptionText() As String
        Get
            Return m_objDB.SysInfo.GetCaptionString(m_intCaptionID)
        End Get
    End Property

    Public ReadOnly Property Caption() As clsCaption
        Get
            If m_objCaption Is Nothing Then
                If Not m_intCaptionID = clsDBConstants.cintNULL Then
                    m_objCaption = clsCaption.GetItem(m_intCaptionID, Me.Database)
                End If
            End If

            Return m_objCaption
        End Get
    End Property

    Public ReadOnly Property TableMethods() As FrameworkCollections.K1Dictionary(Of clsTableMethod)
        Get
            If m_colTableMethod Is Nothing Then
                m_colTableMethod = clsTableMethod.GetList(Me, Me.Database)
            End If
            Return m_colTableMethod
        End Get
    End Property

    Public ReadOnly Property Fields() As FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
        Get
            If m_colFields Is Nothing Then
                m_colFields = New FrameworkCollections.K1DualKeyDictionary(Of clsField, Integer)
            End If
            Return m_colFields
        End Get
    End Property

    Public ReadOnly Property FieldLinks() As FrameworkCollections.K1Dictionary(Of clsFieldLink)
        Get
            If m_colFieldLinks Is Nothing Then
                m_colFieldLinks = New FrameworkCollections.K1Dictionary(Of clsFieldLink)
            End If
            Return m_colFieldLinks
        End Get
    End Property

    Public ReadOnly Property OneToManyLinks() As FrameworkCollections.K1Dictionary(Of clsFieldLink)
        Get
            If m_colOneToManyLinks Is Nothing Then
                m_colOneToManyLinks = New FrameworkCollections.K1Dictionary(Of clsFieldLink)
            End If
            Return m_colOneToManyLinks
        End Get
    End Property

    Public ReadOnly Property ManyToManyLinks() As FrameworkCollections.K1Dictionary(Of clsFieldLink)
        Get
            If m_colManyToManyLinks Is Nothing Then
                m_colManyToManyLinks = New FrameworkCollections.K1Dictionary(Of clsFieldLink)
            End If
            Return m_colManyToManyLinks
        End Get
    End Property

    Public ReadOnly Property DatabaseName() As String
        Get
            Return m_strDatabaseName
        End Get
    End Property

    Public ReadOnly Property TypeDependent() As Boolean
        Get
            Return m_blnTypeDependent
        End Get
    End Property

    Public ReadOnly Property ShowIcon() As Boolean
        Get
            Return m_blnShowIcon
        End Get
    End Property

    Public ReadOnly Property IsLinkTable() As Boolean
        Get
            If m_eClass = clsDBConstants.enumTableClass.LINK_TABLE OrElse _
            m_eClass = clsDBConstants.enumTableClass.LINK_TABLE_ESSENTIAL Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property Processes() As FrameworkCollections.K1Dictionary(Of clsProcess)
        Get
            If m_colProcesses Is Nothing Then
                m_colProcesses = clsProcess.GetList(Me, Me.Database)
            End If

            Return m_colProcesses
        End Get
    End Property

    Public ReadOnly Property ProcessCount() As Integer
        Get
            Dim intCount = (From p In Me.Processes.Values _
                           Where m_objDB.Profile.HasAccess(p.SecurityID) _
                           Select p.ID).Count

            Return intCount
        End Get
    End Property

    Public ReadOnly Property Warnings() As FrameworkCollections.K1Collection(Of clsWarningMessage)
        Get
            If m_colWarnings Is Nothing Then
                m_colWarnings = clsWarningMessage.GetList(Me, Me.Database)
            End If
            Return m_colWarnings
        End Get
    End Property

    Public ReadOnly Property ListColumns() As FrameworkCollections.K1Dictionary(Of clsListColumn)
        Get
            If m_colListColumns Is Nothing Then
                m_colListColumns = New FrameworkCollections.K1Dictionary(Of clsListColumn)
            End If
            Return m_colListColumns
        End Get
    End Property

    Public ReadOnly Property TypeListColumns() As FrameworkCollections.K1Dictionary(Of FrameworkCollections.K1Dictionary(Of clsListColumn))
        Get
            If m_colTypeListColumns Is Nothing Then
                m_colTypeListColumns = New FrameworkCollections.K1Dictionary(Of FrameworkCollections.K1Dictionary(Of clsListColumn))
            End If
            Return m_colTypeListColumns
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsTable
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cTABLE, intID)

            Return New clsTable(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsTable, Integer)
        Dim colTables As FrameworkCollections.K1DualKeyDictionary(Of clsTable, Integer)

        Try
            Dim strSP As String = clsDBConstants.Tables.cTABLE & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colTables = New FrameworkCollections.K1DualKeyDictionary(Of clsTable, Integer)
            For Each objDR As DataRow In objDT.Rows
                Dim objTable As New clsTable(objDR, objDB)
                colTables.Add(objTable.DatabaseName, objTable)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colTables
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTABLE), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cDATABASENAME, m_strDatabaseName)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cCLASS, m_eClass)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cICONID, m_intIconID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cCAPTIONID, m_intCaptionID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cISTYPEDEPENDENT, m_blnTypeDependent)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.Table.cSHOWICON, m_blnShowIcon)

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

#Region " Business Logic "

    Public Function GetListColumns(Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As FrameworkCollections.K1Dictionary(Of clsListColumn)
        If m_blnTypeDependent Then

            If TypeDependent AndAlso _
            Not intTypeID = clsDBConstants.cintNULL AndAlso _
            DatabaseName = clsDBConstants.Tables.cMETADATAPROFILE AndAlso _
            TypeListColumns(CStr(intTypeID)) Is Nothing AndAlso _
            m_objDB.SysInfo IsNot Nothing AndAlso _
            m_objDB.SysInfo.K1Groups IsNot Nothing AndAlso _
            m_objDB.SysInfo.K1Groups.TypeGroups IsNot Nothing Then
                'see if this type belongs to a group (if it does, try the default type for the group)
                If m_objDB.SysInfo.K1Groups.TypeGroups.ContainsKey(CStr(intTypeID)) Then

                    Dim eMDPType As clsDBConstants.enumMDPTypeCodes = m_objDB.SysInfo.K1Groups.TypeGroups(CStr(intTypeID))

                    intTypeID = m_objDB.SysInfo.K1Groups.GetDefaultType(eMDPType)
                End If
            End If

            Dim colListColumns As FrameworkCollections.K1Dictionary(Of clsListColumn) = Nothing

            If m_colTypeListColumns IsNot Nothing Then
                colListColumns = m_colTypeListColumns(CStr(intTypeID))
            End If

            If colListColumns IsNot Nothing Then
                Return colListColumns
            End If
        End If

        Return m_colListColumns
    End Function
#End Region

#Region " Security "

    Public Function HasAccess() As Boolean
        If m_objDB.Profile.HasAccess(m_intSecurityID) AndAlso _
        m_objDB.Profile.LinkTables(CStr(m_intID)) IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

    ''' <summary>
    ''' Returns the field that matches the specified database name.
    ''' </summary>
    ''' <param name="strDatabaseName">Database name of field that should be returned.</param>
    ''' <returns>Returns clsField if match found; otherwise Nothing</returns>
    Public Function GetField(ByVal strDatabaseName As String) As clsField
        For Each objField As clsField In Me.Fields.Values
            If objField.DatabaseName.ToLower = strDatabaseName.ToLower Then
                Return objField
            End If
        Next
        Return Nothing
    End Function


    ''' <summary>
    ''' Returns the field link whos foreign key field matches the specified database name.
    ''' </summary>
    ''' <param name="strDatabaseName">Database name of foreign key field of the field link that should be returned.</param>
    ''' <returns>Returns clsFieldLink if match found; otherwise Nothing</returns>
    Public Function GetFieldLink(ByVal strDatabaseName As String) As clsFieldLink
        For Each objField As clsFieldLink In Me.FieldLinks.Values
            If objField.ForeignKeyField.DatabaseName.ToLower = strDatabaseName.ToLower Then
                Return objField
            End If
        Next
        Return Nothing
    End Function

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_objIcon Is Nothing Then
                m_objIcon.Dispose()
                m_objIcon = Nothing
            End If

            If m_colFields IsNot Nothing Then
                m_colFields.Dispose()
                m_colFields = Nothing
            End If

            If Not m_objCaption Is Nothing Then
                m_objCaption.Dispose()
                m_objCaption = Nothing
            End If

            If m_colTableMethod IsNot Nothing Then
                m_colTableMethod.Dispose()
                m_colTableMethod = Nothing
            End If

            If m_colFieldLinks IsNot Nothing Then
                m_colFieldLinks.Dispose()
                m_colFieldLinks = Nothing
            End If

            If m_colOneToManyLinks IsNot Nothing Then
                m_colOneToManyLinks.Dispose()
                m_colOneToManyLinks = Nothing
            End If

            If m_colManyToManyLinks IsNot Nothing Then
                m_colManyToManyLinks.Dispose()
                m_colManyToManyLinks = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
