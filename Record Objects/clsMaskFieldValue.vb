Imports K1Library.clsDBConstants
Imports K1

#Region " File Information "

'==============================================================================
' This class contains information regarding the value of a mask field
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       12/07/2004  Implemented.
'==============================================================================

#End Region

#End Region

Public Class clsMaskFieldValue
    Implements IDisposable

#Region " Members "

    Private m_objValue As Object
    Private m_strDisplay As String
    Private m_strFileName As String
    Private m_intObjectSecurityID As Integer = clsDBConstants.cintNULL
    Private m_objExtraFieldValue As Object
    Private m_objMaskField As clsMaskField
    Private m_blnIsDirty As Boolean = False
    Private m_blnUseOCR As Boolean = False
    Private m_strFreeText As String
    Private m_objRollbackValue As Object
    Private m_objAutoNumber As clsAutoNumberFormat
    Private m_objMS As clsMaskFieldDictionary
    Private m_strMaskText As String
    Private m_blnAutoNumberGenerated As Boolean
    Private m_intSeqNum As Integer = clsDBConstants.cintNULL
    Private m_objModiArguments As clsModiArguments
    Private m_eUserFileType As enumImageTypes

    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Constructors "

    Public Sub New(ByVal objMaskField As clsMaskField)
        m_objMaskField = objMaskField
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' This is the database value of the mask object
    ''' </summary>
    Public Property Value() As Object
        Get
            Return m_objValue
        End Get
        Set(ByVal objValue As Object)
            m_objValue = objValue
            m_blnIsDirty = True
            If m_objMaskField IsNot Nothing Then
                m_objMaskField.UIUpdate = True
            End If
        End Set
    End Property

    ''' <summary>
    ''' If the mask field is a blob, this is the name of the file we are to store (used in adds)
    ''' </summary>
    Public Property FileName() As String
        Get
            Return m_strFileName
        End Get
        Set(ByVal Value As String)
            m_strFileName = Value
        End Set
    End Property

    Public Property UseOCR() As Boolean
        Get
            Return m_blnUseOCR
        End Get
        Set(value As Boolean)
            m_blnUseOCR = value
        End Set
    End Property

    Public Property ModiArguments As clsModiArguments
        Get
            Return m_objModiArguments
        End Get
        Set(value As clsModiArguments)
            m_objModiArguments = value
        End Set
    End Property

    Public Property ImageType As enumImageTypes
        Get
            Return m_eUserFileType
        End Get
        Set(value As enumImageTypes)
            m_eUserFileType = value
        End Set
    End Property

    ''' <summary>
    ''' If the mask field is a foreign key, this is the associated record's ExternalID
    ''' </summary>
    Public Property Display() As String
        Get
            Return m_strDisplay
        End Get
        Set(ByVal Value As String)
            m_strDisplay = Value
        End Set
    End Property

    Public Property ExtraFieldValue() As Object
        Get
            Return m_objExtraFieldValue
        End Get
        Set(ByVal value As Object)
            m_objExtraFieldValue = value
        End Set
    End Property

    ''' <summary>
    ''' If the mask field is a foreign key, this is the security ID of the associated record
    ''' </summary>
    Public Property ObjectSecurityID() As Integer
        Get
            Return m_intObjectSecurityID
        End Get
        Set(ByVal value As Integer)
            m_intObjectSecurityID = value
        End Set
    End Property

    ''' <summary>
    ''' This is a link back to the parent mask field object
    ''' </summary>
    Public ReadOnly Property MaskField() As clsMaskField
        Get
            Return m_objMaskField
        End Get
    End Property

    Public Property AutoNumber() As clsAutoNumberFormat
        Get
            Return m_objAutoNumber
        End Get
        Set(ByVal value As clsAutoNumberFormat)
            m_objAutoNumber = value
        End Set
    End Property

    Public ReadOnly Property MaskText() As String
        Get
            If m_objAutoNumber IsNot Nothing AndAlso _
            m_objAutoNumber.AutoNumberTokens IsNot Nothing AndAlso _
            m_objAutoNumber.AutoNumberTokens.Count > 0 AndAlso _
            m_strMaskText Is Nothing Then
                m_strMaskText = m_objAutoNumber.CreateMask(m_objAutoNumber.AutoNumberTokens)
            End If
            Return m_strMaskText
        End Get
    End Property

    Public Property RollbackValue() As Object
        Get
            Return m_objRollbackValue
        End Get
        Set(ByVal value As Object)
            m_objRollbackValue = value
        End Set
    End Property

    Public Property MultipleSequence() As clsMaskFieldDictionary
        Get
            Return m_objMS
        End Get
        Set(ByVal value As clsMaskFieldDictionary)
            m_objMS = value
        End Set
    End Property

    Public Property IsDirty() As Boolean
        Get
            Return m_blnIsDirty
        End Get
        Set(ByVal value As Boolean)
            m_blnIsDirty = value
        End Set
    End Property

    Public Property FreeText() As String
        Get
            Return m_strFreeText
        End Get
        Set(ByVal value As String)
            m_strFreeText = value
        End Set
    End Property

    Public Property AutoNumberGenerated() As Boolean
        Get
            Return m_blnAutoNumberGenerated
        End Get
        Set(ByVal value As Boolean)
            m_blnAutoNumberGenerated = value
        End Set
    End Property

    Public Property SequentialNumber() As Integer
        Get
            Return m_intSeqNum
        End Get
        Set(ByVal value As Integer)
            m_intSeqNum = value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Sub InitializeValue(ByVal objValue As Object)
        m_objValue = objValue
    End Sub

    Public Sub InitializeValue(ByVal objValue As Object, ByVal strDisplay As String, ByVal intSecurityID As Integer)
        m_objValue = objValue
        m_strDisplay = strDisplay
        m_intObjectSecurityID = intSecurityID
    End Sub

    Public Sub InitializeValue(ByVal objValue As Object, ByVal strDisplay As String, _
    ByVal intSecurityID As Integer, ByVal objExtraFieldValue As String)
        m_objValue = objValue
        m_strDisplay = strDisplay
        m_intObjectSecurityID = intSecurityID
        m_objExtraFieldValue = objExtraFieldValue
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objMaskField = Nothing
                m_objValue = Nothing
                m_objAutoNumber = Nothing
                m_objRollbackValue = Nothing
                m_objMS = Nothing
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
