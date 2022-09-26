''' <summary>
''' Used to create or update a field using the DRM
''' </summary>
Public Class clsDRMTypeField
    Inherits clsDRMBase

#Region " Members "

    '-- General Values
    Private m_objField As clsField
    Private m_strCaption As String
    Private m_intCaptionID As Integer = clsDBConstants.cintNULL
    Private m_blnIsVisible As Boolean
    Private m_blnIsMandatory As Boolean
    Private m_blnIsReadOnly As Boolean
    Private m_intAppliesToTypeID As Integer
    Private m_intSortOrder As Integer = 5
    Private m_blnIsMultipleSequenceField As Boolean
    Private m_blnDeterminesMultipleSequence As Boolean
    Private m_blnAllowFreeTextEntry As Boolean

    '-- Autonumber Values
    Private m_objAutoFill As clsAutoFillInfo
#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Creates a new field
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, _
    ByVal objField As clsField, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intAppliesToTypeID As Integer, _
    ByVal strCaption As String, _
    ByVal blnIsVisible As Boolean, _
    ByVal objAutoFillInfo As clsAutoFillInfo, _
    ByVal blnIsMultipleSequenceField As Boolean, _
    ByVal blnDeterminesMultipleSequence As Boolean, _
    ByVal blnAllowFreeTextEntry As Boolean, _
    ByVal blnMandatory As Boolean, _
    ByVal blnReadOnly As Boolean)
        MyBase.New(objDB, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_objField = objField
        m_strCaption = strCaption
        m_intAppliesToTypeID = intAppliesToTypeID
        m_blnIsVisible = blnIsVisible
        m_objAutoFill = objAutoFillInfo
        m_blnIsMultipleSequenceField = blnIsMultipleSequenceField
        m_blnDeterminesMultipleSequence = blnDeterminesMultipleSequence
        m_blnAllowFreeTextEntry = blnAllowFreeTextEntry
        m_blnIsMandatory = blnMandatory
        m_blnIsReadOnly = blnReadOnly
        m_intSortOrder = objField.SortOrder
    End Sub
#End Region

#Region " From Existing "

    ''' <summary>
    ''' Creates a DRM Field from an existing field database object
    ''' </summary>
    Public Sub New(ByVal objDB As clsDB, ByVal objField As clsField, ByVal objTF As clsTypeField)
        MyBase.New(objDB, objTF)
        m_objField = objField
        m_intAppliesToTypeID = objTF.AppliesToTypeID
        m_blnIsVisible = objTF.IsVisible
        m_blnIsMandatory = objTF.IsMandatory
        m_blnIsReadOnly = objTF.IsReadOnly
        m_intSortOrder = objTF.SortOrder
        m_objAutoFill = objTF.AutoFillInfo
        m_blnIsMultipleSequenceField = objTF.IsMultipleSequenceField
        m_blnDeterminesMultipleSequence = objTF.DeterminesMultipleSequence
        m_blnAllowFreeTextEntry = objTF.AllowFreeTextEntry

        If objTF.Caption IsNot Nothing Then
            m_strCaption = objTF.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
            m_intCaptionID = objTF.Caption.ID
        End If
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property TypeField() As clsTypeField
        Get
            Return CType(m_objDBObj, clsTypeField)
        End Get
        Set(ByVal value As clsTypeField)
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

    Public Property IsMandatory() As Boolean
        Get
            Return m_blnIsMandatory
        End Get
        Set(ByVal value As Boolean)
            m_blnIsMandatory = value
        End Set
    End Property

    Public Property IsReadOnly() As Boolean
        Get
            Return m_blnIsReadOnly
        End Get
        Set(ByVal value As Boolean)
            m_blnIsReadOnly = value
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

    Public Property AutoFillInfo() As clsAutoFillInfo
        Get
            Return m_objAutoFill
        End Get
        Set(ByVal value As clsAutoFillInfo)
            m_objAutoFill = value
        End Set
    End Property

    Public Property IsMultipleSequenceField() As Boolean
        Get
            Return m_blnIsMultipleSequenceField
        End Get
        Set(ByVal value As Boolean)
            m_blnIsMultipleSequenceField = value
        End Set
    End Property

    Public Property DeterminesMultipleSequence() As Boolean
        Get
            Return m_blnDeterminesMultipleSequence
        End Get
        Set(ByVal value As Boolean)
            m_blnDeterminesMultipleSequence = value
        End Set
    End Property

    Public Property AllowFreeTextEntry() As Boolean
        Get
            Return m_blnAllowFreeTextEntry
        End Get
        Set(ByVal value As Boolean)
            m_blnAllowFreeTextEntry = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Insert Update "

    Public Sub InsertUpdate()
        Dim intCaptionID As Integer = clsDBConstants.cintNULL
        Dim blnCreatedTransaction As Boolean = False

        Try
            'Start a transaction
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            '================================================================================
            'Create the field and field heading captions
            '================================================================================
            If String.IsNullOrEmpty(m_strCaption) OrElse m_strCaption = m_objField.Caption.GetString(m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID) Then
                m_strCaption = Nothing
                If m_objDBObj IsNot Nothing AndAlso Not m_intCaptionID = clsDBConstants.cintNULL Then
                    DeleteCaption(CType(m_objDBObj, clsTypeField).Caption)
                    m_intCaptionID = clsDBConstants.cintNULL
                End If
            End If

            If m_intID = clsDBConstants.cintNULL OrElse m_intCaptionID = clsDBConstants.cintNULL Then
                If Not String.IsNullOrEmpty(m_strCaption) Then
                    intCaptionID = CreateCaption(m_strExternalID, m_strCaption)
                End If
            Else
                intCaptionID = m_intCaptionID
                Dim strCaption As String = TypeField.Caption.GetString( _
                    m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)

                If Not String.IsNullOrEmpty(m_strCaption) AndAlso _
                Not m_strCaption = strCaption Then
                    intCaptionID = CreateCaption(TypeField.Caption, m_strCaption)
                End If
            End If

            '================================================================================
            'Insert the field
            '================================================================================
            Dim objTypeField As New clsTypeField(m_objField, m_intID, m_strExternalID, _
                m_intSecurityID, m_intAppliesToTypeID, m_blnIsVisible, intCaptionID, m_intSortOrder, _
                m_objAutoFill, m_blnIsMultipleSequenceField, m_blnDeterminesMultipleSequence, _
                m_blnAllowFreeTextEntry, m_blnIsMandatory, m_blnIsReadOnly)

            objTypeField.InsertUpdate()

            If m_objField.TypeFieldInfos(CStr(m_intAppliesToTypeID)) Is Nothing Then
                m_objField.TypeFieldInfos.Add(CStr(m_intAppliesToTypeID), objTypeField)
            Else
                m_objField.TypeFieldInfos(CStr(m_intAppliesToTypeID)) = objTypeField
            End If

            m_intID = objTypeField.ID
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

            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cTYPEFIELDINFO, TypeField.ID)

            DeleteCaption(TypeField.Caption)

            m_objField.TypeFieldInfos.Remove(CStr(m_intAppliesToTypeID))

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
        m_objField = Nothing
        m_objAutoFill = Nothing
    End Sub
#End Region

End Class
