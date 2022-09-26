Public Class clsDRMListColumn
    Inherits clsDRMBase

#Region " Members "

    Private m_intFieldID As Integer
    Private m_intAppliesToTypeID As Integer
    Private m_intWidth As Integer
    Private m_intCaptionID As Integer = clsDBConstants.cintNULL
    Private m_strCaption As String
    Private m_intSortOrder As Integer
#End Region

#Region " Constructors "

#Region " New Field "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, _
    ByVal intTypeID As Integer, _
    ByVal intFieldID As Integer, _
    ByVal intAppliesToTypeID As Integer, _
    ByVal intWidth As Integer, _
    ByVal strCaption As String, _
    ByVal intSortOrder As Integer)
        MyBase.New(objDB, strExternalID, intSecurityID, intTypeID)
        m_intFieldID = intFieldID
        m_intAppliesToTypeID = intAppliesToTypeID
        m_intWidth = intWidth
        m_strCaption = strCaption
        m_intSortOrder = intSortOrder
    End Sub
#End Region

#Region " From Existing "

    Public Sub New(ByVal objDB As clsDB, _
    ByVal objListColumn As clsListColumn)
        MyBase.New(objDB, objListColumn)
        m_intFieldID = objListColumn.FieldID
        m_intAppliesToTypeID = objListColumn.AppliesToTypeID
        m_intWidth = objListColumn.Width
        If objListColumn.Caption IsNot Nothing Then
            m_strCaption = objListColumn.Caption.GetString( _
                objDB.Profile.LanguageID, objDB.Profile.DefaultLanguageID)
            m_intCaptionID = objListColumn.Caption.ID
        End If
        m_intSortOrder = objListColumn.SortOrder
    End Sub
#End Region

#End Region

#Region " Properties "

    Public ReadOnly Property FieldID() As Integer
        Get
            Return m_intFieldID
        End Get
    End Property

    Public Property Width() As Integer
        Get
            Return m_intWidth
        End Get
        Set(ByVal value As Integer)
            m_intWidth = value
        End Set
    End Property

    Public Property CaptionText() As String
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

    Public ReadOnly Property ListColumn() As clsListColumn
        Get
            Return CType(m_objDBObj, clsListColumn)
        End Get
    End Property
#End Region

#Region " Methods "

    Public Sub InsertUpdate(Optional ByVal objTable As clsTable = Nothing)
        Dim intCaptionID As Integer = clsDBConstants.cintNULL
        Dim blnCreatedTransaction As Boolean = False
        Dim objListColumn As clsListColumn

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            Dim objField As clsField = m_objDB.SysInfo.Fields(m_intFieldID)
            Dim objDRMCaption As clsDRMCaption = Nothing

            If objField.DataType = SqlDbType.Image Then
                m_intWidth = 0
                m_strCaption = Nothing
            End If

            If m_intID = clsDBConstants.cintNULL OrElse m_intCaptionID = clsDBConstants.cintNULL Then
                If Not String.IsNullOrEmpty(m_strCaption) Then
                    intCaptionID = CreateCaption("Field Column Heading - " & objField.DatabaseName, m_strCaption)
                End If
            Else
                intCaptionID = m_intCaptionID
                Dim strCaption As String = ListColumn.Caption.GetString( _
                    m_objDB.Profile.LanguageID, m_objDB.Profile.DefaultLanguageID)

                If String.IsNullOrEmpty(m_strCaption) Then
                    objDRMCaption = New clsDRMCaption(m_objDB, ListColumn.Caption, "")
                    intCaptionID = clsDBConstants.cintNULL
                ElseIf Not m_strCaption = strCaption Then
                    intCaptionID = CreateCaption(ListColumn.Caption, m_strCaption)
                End If
            End If

            objListColumn = New clsListColumn(m_objDB, m_intID, m_strExternalID, m_intSecurityID, _
                 m_intFieldID, m_intAppliesToTypeID, m_intWidth, intCaptionID, m_intSortOrder)
            objListColumn.InsertUpdate()

            If objDRMCaption IsNot Nothing Then
                objDRMCaption.Delete()
            End If

            If objTable Is Nothing Then
                objTable = objField.Table
            End If

            m_objDB.SysInfo.DRMInsertUpdateListColumn(objTable, objListColumn)

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

            clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cLISTCOLUMN, ListColumn.ID)
            m_objDB.SysInfo.DRMDeleteListColumn(ListColumn)

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
