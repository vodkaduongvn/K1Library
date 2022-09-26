''' <summary>
''' Used to create or update a field using the DRM
''' </summary>
Public Class clsDRMTypeFieldLink
    Inherits clsDRMBase

#Region " Members "

    '-- General Values
    Private m_objFieldLink As clsFieldLink
    Private m_strCaption As String
    Private m_intCaptionID As Integer = clsDBConstants.cintNULL
    Private m_blnIsVisible As Boolean
    Private m_intAppliesToTypeID As Integer
    Private m_intSortOrder As Integer = 5
#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new field
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal objFieldLink As clsFieldLink, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intAppliesToTypeID As Integer, _
    ByVal strCaption As String, _
    ByVal blnIsVisible As Boolean)
        MyBase.New(objDB, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_objFieldLink = objFieldLink
        m_strCaption = strCaption
        m_intAppliesToTypeID = intAppliesToTypeID
        m_blnIsVisible = blnIsVisible
        m_intSortOrder = objFieldLink.SortOrder
    End Sub
#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRM Field from an existing field database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objFieldLink As clsFieldLink, ByVal objTFL As clsTypeFieldLink)
        MyBase.New(objDB, objTFL)
        m_objFieldLink = objFieldLink
        m_intAppliesToTypeID = objTFL.AppliesToTypeID
        m_blnIsVisible = objTFL.IsVisible
        m_intSortOrder = objTFL.SortOrder

        If objTFL.Caption IsNot Nothing Then
            m_strCaption = objTFL.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
            m_intCaptionID = objTFL.Caption.ID
        End If
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property TypeFieldLink() As clsTypeFieldLink
        Get
            Return CType(m_objDBObj, clsTypeFieldLink)
        End Get
        Set(ByVal value As clsTypeFieldLink)
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
#End Region

#Region " Methods "

#Region " Insert Update "

    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False
        Dim intCaptionID As Integer = clsDBConstants.cintNULL

        Try
            'Start a transaction
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            '================================================================================
            'Create the field and field heading captions
            '================================================================================
            If String.IsNullOrEmpty(m_strCaption) OrElse m_strCaption = m_objFieldLink.Caption.GetString(m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID) Then
                m_strCaption = Nothing
                If m_objDBObj IsNot Nothing AndAlso Not m_intCaptionID = clsDBConstants.cintNULL Then
                    DeleteCaption(CType(m_objDBObj, clsTypeFieldLink).Caption)
                    m_intCaptionID = clsDBConstants.cintNULL
                End If
            End If

            If m_intID = clsDBConstants.cintNULL OrElse m_intCaptionID = clsDBConstants.cintNULL Then
                If Not String.IsNullOrEmpty(m_strCaption) Then
                    intCaptionID = CreateCaption(m_strExternalID, m_strCaption)
                End If
            Else
                intCaptionID = m_intCaptionID
                Dim strCaption As String = TypeFieldLink.Caption.GetString( _
                    m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)

                If Not String.IsNullOrEmpty(m_strCaption) AndAlso _
                Not m_strCaption = strCaption Then
                    intCaptionID = CreateCaption(TypeFieldLink.Caption, m_strCaption)
                End If
            End If

            '================================================================================
            'Insert the field
            '================================================================================
            Dim objTypeFieldLink As New clsTypeFieldLink(m_objFieldLink, m_intID, m_strExternalID, _
                m_intSecurityID, m_intAppliesToTypeID, m_blnIsVisible, intCaptionID, m_intSortOrder)

            objTypeFieldLink.InsertUpdate()

            If m_objFieldLink.TypeFieldLinkInfos(CStr(m_intAppliesToTypeID)) Is Nothing Then
                m_objFieldLink.TypeFieldLinkInfos.Add(CStr(m_intAppliesToTypeID), objTypeFieldLink)
            Else
                m_objFieldLink.TypeFieldLinkInfos(CStr(m_intAppliesToTypeID)) = objTypeFieldLink
            End If

            m_intID = objTypeFieldLink.ID
            m_intCaptionID = intCaptionID

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub
#End Region

#Region " Delete "

    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cTYPEFIELDLINKINFO, TypeFieldLink.ID)

            DeleteCaption(TypeFieldLink.Caption)

            m_objFieldLink.TypeFieldLinkInfos.Remove(CStr(m_intAppliesToTypeID))

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
        m_objFieldLink = Nothing
    End Sub
#End Region

End Class
