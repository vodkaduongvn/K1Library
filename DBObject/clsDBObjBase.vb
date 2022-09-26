#Region " File Information "

'=====================================================================
' This is the base class for all database objects
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date        Description
'---------------------------------------------------------------------
' KSD       11/06/2004  Implemented.
'=====================================================================

#End Region

#End Region

Public MustInherit Class clsDBObjBase
    Implements IDisposable

#Region " Member Variables "

    Protected m_intID As Integer
    Protected m_strExternalID As String
    Protected m_intSecurityID As Integer
    Protected m_objSecurity As clsSecurity
    Protected m_intTypeID As Integer
    Protected m_objDB As clsDB
    Protected m_blnDisposedValue As Boolean = False
#End Region

#Region " Constructors "

    'TODO: Delete this New
    Public Sub New()
        MyBase.New()
        m_intID = clsDBConstants.cintNULL
        m_intSecurityID = clsDBConstants.cintNULL
        m_intTypeID = clsDBConstants.cintNULL
        m_strExternalID = clsDBConstants.cstrNULL
    End Sub

    Public Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        Me.New()
        m_intID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cID, clsDBConstants.cintNULL), Integer)
        m_strExternalID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cEXTERNALID, clsDBConstants.cstrNULL), String)
        m_intSecurityID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cSECURITYID, clsDBConstants.cintNULL), Integer)
        m_intTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cTYPEID, clsDBConstants.cintNULL), Integer)
        m_objDB = objDB
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal intID As Integer, _
    ByVal strExternalID As String, ByVal intSecurityID As Integer, ByVal intTypeID As Integer)
        m_objDB = objDB
        m_intID = intID
        m_strExternalID = strExternalID
        m_intSecurityID = intSecurityID
        m_intTypeID = intTypeID
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property ID() As Integer
        Get
            Return m_intID
        End Get
    End Property

    Public ReadOnly Property ExternalID() As String
        Get
            Return m_strExternalID
        End Get
    End Property

    Public ReadOnly Property Database() As clsDB
        Get
            Return m_objDB
        End Get
    End Property

    Public ReadOnly Property SecurityID() As Integer
        Get
            Return m_intSecurityID
        End Get
    End Property

    Public ReadOnly Property TypeID() As Integer
        Get
            Return m_intTypeID
        End Get
    End Property

    ''' <summary>
    ''' This is the ID field converted to a string (used a lot in dictionaries indexes)
    ''' </summary>
    Public ReadOnly Property KeyID() As String
        Get
            Return CStr(m_intID)
        End Get
    End Property
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub DisposeDBObject()
    End Sub

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
                m_objSecurity = Nothing

                DisposeDBObject()
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