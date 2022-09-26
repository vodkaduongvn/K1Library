#Region " File Information "

'==============================================================================
' This class is a fieldlink/UI Control combination for use on the Mask UI Page
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       09/03/2007  Implemented.
'==============================================================================

#End Region

#End Region

Public Class clsMaskFieldLink
    Inherits clsMaskBase

#Region " Members "

    Private m_objFieldLink As clsFieldLink

    'Used In Add, Modify, and Searches
    Private m_colIDs As Hashtable
    Private m_colNewIDs As Hashtable
    Private m_colRemoveIDs As Hashtable

    Private m_blnExpanded As Boolean = False
    Private m_blnValuesLoaded As Boolean = False
    Private m_blnUpdated As Boolean = False
    Private m_blnIsDirty As Boolean = False
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, ByVal objFieldLink As clsFieldLink, _
    ByVal eObjType As enumMaskObjectType, ByVal eMaskType As clsTableMask.enumMaskType)
        MyBase.New(objDB, eObjType, eMaskType)
        m_objFieldLink = objFieldLink
    End Sub

    Protected Sub New(ByVal objDB As clsDB, _
    ByVal objFieldLink As clsFieldLink, ByVal strCaption As String, _
    ByVal blnIsVisible As Boolean, ByVal eObjType As enumMaskObjectType, _
    ByVal eMaskType As clsTableMask.enumMaskType)
        MyBase.New(objDB, eObjType, eMaskType)
        m_objFieldLink = objFieldLink
        m_strCaption = strCaption
        m_blnIsVisible = blnIsVisible
    End Sub
#End Region

#Region " Properties "

    Public Property FieldLink() As clsFieldLink
        Get
            Return m_objFieldLink
        End Get
        Set(ByVal Value As clsFieldLink)
            m_objFieldLink = Value
        End Set
    End Property

    Public Property IDCollection() As Hashtable
        Get
            If m_colIDs Is Nothing Then
                m_colIDs = New Hashtable
            End If
            Return m_colIDs
        End Get
        Set(ByVal Value As Hashtable)
            m_colIDs = Value
            m_blnIsDirty = True
            m_blnUIUpdate = True
        End Set
    End Property

    Public Property NewIDCollection() As Hashtable
        Get
            If m_colNewIDs Is Nothing Then
                m_colNewIDs = New Hashtable
            End If
            Return m_colNewIDs
        End Get
        Set(ByVal Value As Hashtable)
            m_colNewIDs = Value
        End Set
    End Property

    Public Property RemoveIDCollection() As Hashtable
        Get
            If m_colRemoveIDs Is Nothing Then
                m_colRemoveIDs = New Hashtable
            End If
            Return m_colRemoveIDs
        End Get
        Set(ByVal Value As Hashtable)
            m_colRemoveIDs = Value
        End Set
    End Property

    Public Overrides ReadOnly Property HasTableAccess() As Boolean
        Get
            If m_eObjType = enumMaskObjectType.ONETOMANY Then
                Return (m_objDB.Profile.HasAccess( _
                    m_objFieldLink.IdentityTable.SecurityID) AndAlso _
                    m_objDB.Profile.LinkTables( _
                    CType(m_objFieldLink.IdentityTable.ID, String)) IsNot Nothing)
            Else
                Return (m_objDB.Profile.HasAccess( _
                    m_objFieldLink.LinkedTable.SecurityID) AndAlso _
                    m_objDB.Profile.LinkTables( _
                    CType(m_objFieldLink.LinkedTable.ID, String)) IsNot Nothing)
            End If
        End Get
    End Property

    Public Property Expanded() As Boolean
        Get
            Return m_blnExpanded
        End Get
        Set(ByVal value As Boolean)
            m_blnExpanded = value
        End Set
    End Property

    Public Property Updated() As Boolean
        Get
            Return m_blnUpdated
        End Get
        Set(ByVal value As Boolean)
            m_blnUpdated = value
        End Set
    End Property

    Public Property ValuesLoaded() As Boolean
        Get
            Return m_blnValuesLoaded
        End Get
        Set(ByVal value As Boolean)
            m_blnValuesLoaded = value
        End Set
    End Property

    Public ReadOnly Property HasForeignKeyTableAccess() As Boolean
        Get
            '2016/07/29 -- James -- Bug ID: 1600003145. Added in access checks for LinkedTables
            If m_objFieldLink.ForeignKeyTable.IsLinkTable Then
                If m_objDB.Profile.HasAccess(m_objFieldLink.LinkedTable.SecurityID) AndAlso m_objDB.Profile.LinkTables(CType(m_objFieldLink.LinkedTable.ID, String)) IsNot Nothing Then
                    Return True
                Else
                    Return False
                End If
            Else
                If (m_objDB.Profile.HasAccess(m_objFieldLink.ForeignKeyTable.SecurityID) AndAlso m_objDB.Profile.LinkTables(CType(m_objFieldLink.ForeignKeyTable.ID, String)) IsNot Nothing) Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Get
    End Property

    Public Property IsDirty() As Boolean
        Get
            Return m_blnIsDirty
        End Get
        Set(ByVal value As Boolean)
            m_blnIsDirty = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " CreateMaskCollection "

    ''' <summary>
    ''' Retrieves a collection of field link mask objects (specific to the Mask Object Type)
    ''' </summary>
    Public Shared Function CreateMaskCollection( _
    ByVal objTable As clsTable, ByVal eObjType As enumMaskObjectType, _
    ByVal eMaskType As clsTableMask.enumMaskType, _
    Optional ByVal intTypeID As Integer = clsDBConstants.cintNULL) As clsMaskFieldLinkDictionary
        Try
            Dim objProfile As clsUserProfile = objTable.Database.Profile
            Dim colMasks As New clsMaskFieldLinkDictionary
            Dim colFieldLinks As FrameworkCollections.K1Dictionary(Of clsFieldLink)

            If eObjType = enumMaskObjectType.ONETOMANY Then
                colFieldLinks = objTable.OneToManyLinks
            Else
                colFieldLinks = objTable.ManyToManyLinks
            End If

            'Create our collection of maskobjects if new instance or we have found a record
            For Each objFieldLink As clsFieldLink In colFieldLinks.Values
                Dim objTFLI As clsTypeFieldLink = Nothing
                Dim strCaption As String = Nothing
                Dim blnVisible As Boolean = objFieldLink.IsVisible
                Dim intSecurityID As Integer = objFieldLink.SecurityID

                If objTable.TypeDependent AndAlso Not intTypeID = clsDBConstants.cintNULL Then
                    objTFLI = objFieldLink.TypeFieldLinkInfos(CStr(intTypeID))
                End If

                If objTable.TypeDependent AndAlso Not objTFLI Is Nothing Then
                    strCaption = objTFLI.CaptionText
                    blnVisible = objTFLI.IsVisible
                    intSecurityID = objTFLI.SecurityID
                End If

                Dim blnIsVisible As Boolean = ( _
                    blnVisible AndAlso _
                    objProfile.HasAccess(intSecurityID))

                If blnIsVisible Then
                    If strCaption Is Nothing Then strCaption = objFieldLink.CaptionText
                End If

                colMasks.Add(New clsMaskFieldLink(objTable.Database, _
                    objFieldLink, strCaption, blnIsVisible, eObjType, eMaskType))
            Next

            Return colMasks
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

    Public Sub LoadValues(ByVal intRecordID As Integer)
        If Not m_blnValuesLoaded Then
            m_colIDs = New Hashtable

            If Not m_objFieldLink.ForeignKeyTable Is Nothing Then
                Dim objSE As New clsSearchElement(clsSearchFilter.enumOperatorType.NONE, _
                    m_objFieldLink.LinkedTable.DatabaseName & ".*" & _
                    m_objFieldLink.ForeignKeyTable.DatabaseName & "." & _
                    m_objFieldLink.LinkTableOppositeFieldLink.ForeignKeyField.DatabaseName & "." & _
                    m_objFieldLink.ForeignKeyField.DatabaseName, _
                    clsSearchFilter.enumComparisonType.EQUAL, intRecordID)

                Dim objSelectInfo As New clsSelectInfo(m_objFieldLink.LinkedTable, _
                    Nothing, Nothing, objSE)

                Dim objDT As DataTable = objSelectInfo.DataTable

                For Each objRow As DataRow In objDT.Rows
                    m_colIDs.Add(CStr(objRow(0)), CInt(objRow(0)))
                Next

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If
            End If

            m_blnValuesLoaded = True
        End If
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeObject()
        m_objFieldLink = Nothing

        If m_colIDs IsNot Nothing Then
            m_colIDs.Clear()
            m_colIDs = Nothing
        End If

        If m_colNewIDs IsNot Nothing Then
            m_colNewIDs.Clear()
            m_colNewIDs = Nothing
        End If

        If m_colRemoveIDs IsNot Nothing Then
            m_colRemoveIDs.Clear()
            m_colRemoveIDs = Nothing
        End If

        m_objTableMask = Nothing
    End Sub
#End Region

End Class
