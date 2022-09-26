#Region " File Information "

'=====================================================================
' This is the base class for any mask object (object which represents
' a field/field link on the mask screen)
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Kevin     08/03/2007   Implemented.
'=====================================================================

#End Region

#End Region

Public MustInherit Class clsMaskBase
    Implements IDisposable

#Region " Members "

    Protected m_objDB As clsDB

    Protected m_strCaption As String

    Protected m_eCheckState As Windows.Forms.CheckState = Windows.Forms.CheckState.Unchecked
    Protected m_blnIsVisible As Boolean = True
    Protected m_eObjType As enumMaskObjectType
    Protected m_eMaskType As clsTableMask.enumMaskType
    Protected m_eMaskSearchType As enumMaskSearchType = enumMaskSearchType.SINGLE
    Protected m_objTableMask As clsTableMask
    Protected m_objSearchFilter As clsSearchFilter
    Protected m_blnUIUpdate As Boolean = False

    Protected m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Enumerations "

    Public Enum enumMaskSearchType
        [SINGLE] = 1
        RANGE = 2
    End Enum

    Public Enum enumMaskObjectType
        FIELD = 1
        ONETOMANY = 2
        MANYTOMANY = 3
    End Enum
#End Region

#Region " Constructors "

    Protected Sub New(ByVal objDB As clsDB, _
    ByVal eObjType As enumMaskObjectType, ByVal eMaskType As clsTableMask.enumMaskType)
        m_objDB = objDB
        m_eObjType = eObjType
        m_eMaskSearchType = enumMaskSearchType.SINGLE
        m_eMaskType = eMaskType
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Database() As clsDB
        Get
            Return m_objDB
        End Get
    End Property

    Public Property Caption() As String
        Get
            Return m_strCaption
        End Get
        Set(ByVal Value As String)
            m_strCaption = Value
        End Set
    End Property

    Public Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
        Set(ByVal Value As Boolean)
            m_blnIsVisible = Value
        End Set
    End Property

    Public Property CheckState() As Windows.Forms.CheckState
        Get
            Return m_eCheckState
        End Get
        Set(ByVal value As Windows.Forms.CheckState)
            m_eCheckState = value
        End Set
    End Property

    Public ReadOnly Property ObjectType() As enumMaskObjectType
        Get
            Return m_eObjType
        End Get
    End Property

    Public Property MaskSearchType() As enumMaskSearchType
        Get
            Return m_eMaskSearchType
        End Get
        Set(ByVal value As enumMaskSearchType)
            m_eMaskSearchType = value
        End Set
    End Property

    Public Property MaskType() As clsTableMask.enumMaskType
        Get
            Return m_eMaskType
        End Get
        Set(ByVal value As clsTableMask.enumMaskType)
            m_eMaskType = value
        End Set
    End Property

    Public Property TableMask() As clsTableMask
        Get
            Return m_objTableMask
        End Get
        Set(ByVal value As clsTableMask)
            m_objTableMask = value
        End Set
    End Property

    ''' <summary>
    ''' Set this if you want to forcefully replace filters on mask collections
    ''' </summary>
    Public Property SearchFilter() As clsSearchFilter
        Get
            Return m_objSearchFilter
        End Get
        Set(ByVal value As clsSearchFilter)
            m_objSearchFilter = value
        End Set
    End Property

    Public Property UIUpdate() As Boolean
        Get
            Return m_blnUIUpdate
        End Get
        Set(ByVal value As Boolean)
            m_blnUIUpdate = value
        End Set
    End Property

    Public MustOverride ReadOnly Property HasTableAccess() As Boolean
#End Region

#Region " Methods "

    Protected MustOverride Sub DisposeObject()
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing

                If m_objTableMask IsNot Nothing Then
                    m_objTableMask.Dispose()
                    m_objTableMask = Nothing
                End If

                DisposeObject()
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
