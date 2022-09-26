''' <summary>
''' Used to create or update a field using the DRM
''' </summary>
Public Class clsDRMFieldLink
    Inherits clsDRMBase

#Region " Members "

    '-- General Values
    Private m_objTable As clsTable
    Private m_objField As clsField
    Private m_strForeignTable As String
    Private m_strCaption As String
    Private m_blnIsVisible As Boolean
    Private m_intSortOrder As Integer = 6
    Private m_blnDisplayASDropDown As Boolean
    Private m_blnIsExpanded As Boolean
#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new field
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal objTable As clsTable, _
    ByVal objField As clsField, _
    ByVal intSecurityID As Integer, _
    ByVal strForeignTable As String, _
    ByVal strCaption As String, _
    ByVal blnIsVisible As Boolean, _
    ByVal blnDisplayASDropDown As Boolean)
        MyBase.New(objDB, clsDBConstants.cstrNULL, intSecurityID, clsDBConstants.cintNULL)
        m_objTable = objTable
        m_objField = objField
        m_strForeignTable = strForeignTable
        m_strCaption = strCaption
        m_blnIsVisible = blnIsVisible
        m_blnDisplayASDropDown = blnDisplayASDropDown
    End Sub
#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRM Field from an existing field database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objFieldLink As clsFieldLink)
        MyBase.New(objDB, objFieldLink)
        m_objTable = objFieldLink.ForeignKeyTable
        m_objField = objFieldLink.ForeignKeyField
        m_strForeignTable = objFieldLink.IdentityTable.DatabaseName

        If objFieldLink.Caption IsNot Nothing Then
            m_strCaption = objFieldLink.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
        End If

        m_blnIsVisible = objFieldLink.IsVisible
        m_blnDisplayASDropDown = objFieldLink.DisplayAsDropDown
        m_blnIsExpanded = objFieldLink.IsExpanded
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property FieldLink() As clsFieldLink
        Get
            Return CType(m_objDBObj, clsFieldLink)
        End Get
        Set(ByVal value As clsFieldLink)
            m_objDBObj = value
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

    Public Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
        Set(ByVal value As Boolean)
            m_blnIsVisible = value
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

    Public Property DisplayAsDropDown() As Boolean
        Get
            Return m_blnDisplayASDropDown
        End Get
        Set(ByVal value As Boolean)
            m_blnDisplayASDropDown = value
        End Set
    End Property

    Public Property IsExpanded() As Boolean
        Get
            Return m_blnIsExpanded
        End Get
        Set(ByVal value As Boolean)
            m_blnIsExpanded = value
        End Set
    End Property

    Public Property Table() As clsTable
        Get
            Return m_objTable
        End Get
        Set(ByVal value As clsTable)
            m_objTable = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Insert Update "

    Public Sub InsertUpdate(Optional ByVal blnCreateAuditTrailRecord As Boolean = True)

        Dim blnCreatedTransaction As Boolean = False

        Try
            'Start a transaction
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            Dim objIdentityTable As clsTable = m_objDB.SysInfo.Tables(m_strForeignTable)
            Dim objIdentityField As clsField = m_objDB.SysInfo.Fields(objIdentityTable.ID & "_" & clsDBConstants.Fields.cID)

            If m_strExternalID = clsDBConstants.cstrNULL Then
                m_strExternalID = m_strForeignTable & " - " & MakePlural(m_objTable.DatabaseName)
            End If

            Dim intFKeyCaptionID As Integer

            If String.IsNullOrEmpty(m_strCaption) Then
                m_strCaption = MakePlural(m_objTable.DatabaseName)
            End If

            Dim objFieldLink As clsFieldLink = m_objField.FieldLink

            If objFieldLink Is Nothing OrElse objFieldLink.Caption Is Nothing Then
                intFKeyCaptionID = CreateCaption("FieldLink - " & m_strExternalID,
                                                 m_strCaption,
                                                 blnCreateAuditTrailRecord)
            Else
                intFKeyCaptionID = objFieldLink.Caption.ID
                Dim strFieldLinkCaption As String = objFieldLink.Caption.GetString(m_objDB.Profile.LanguageID,
                                                                                   m_objDB.Profile.DefaultLanguageID)

                If Not String.IsNullOrEmpty(m_strCaption) AndAlso Not m_strCaption = strFieldLinkCaption Then
                    intFKeyCaptionID = CreateCaption(objFieldLink.Caption,
                                                     m_strCaption,
                                                     blnCreateAuditTrailRecord)
                End If
            End If

            objFieldLink = New clsFieldLink(m_objDB, m_intID, m_strExternalID,
                                            m_intSecurityID, objIdentityField.ID,
                                            m_objField.ID, m_blnIsVisible,
                                            m_intSortOrder, intFKeyCaptionID,
                                            m_blnDisplayASDropDown, m_blnIsExpanded)

            objFieldLink.InsertUpdate(blnCreateAuditTrailRecord)

            If m_objField.FieldLink Is Nothing Then
                SystemDB.CreateForeignKey(m_objTable.DatabaseName, m_objField.DatabaseName, objIdentityTable.DatabaseName)

                SystemDB.DropDefaultConstraint(m_objTable.DatabaseName, m_objField.DatabaseName)
                clsTableIndex.CreateForeignKeyIndex(SystemDB, m_objTable.DatabaseName, m_objField.DatabaseName)
            Else
                For Each strKey As String In m_objField.FieldLink.TypeFieldLinkInfos.Keys
                    objFieldLink.TypeFieldLinkInfos(strKey) = m_objField.FieldLink.TypeFieldLinkInfos(strKey)
                Next
            End If

            m_objDB.SysInfo.DRMInsertUpdateFieldLink(m_objTable, objFieldLink)

            m_intID = objFieldLink.ID

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
        m_objTable = Nothing
        m_objField = Nothing
    End Sub
#End Region

End Class
