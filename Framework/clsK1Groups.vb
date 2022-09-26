Public Class clsK1Groups
    Implements IDisposable

#Region " Members "

    Private m_colDefaultTypes As FrameworkCollections.K1Dictionary(Of Integer)
    Private m_colIDValues As FrameworkCollections.K1Dictionary(Of Hashtable)
    Private m_colTypeGroups As FrameworkCollections.K1Dictionary(Of clsDBConstants.enumMDPTypeCodes)

    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB)

        m_colDefaultTypes = New FrameworkCollections.K1Dictionary(Of Integer)
        m_colIDValues = New FrameworkCollections.K1Dictionary(Of Hashtable)
        m_colTypeGroups = New FrameworkCollections.K1Dictionary(Of clsDBConstants.enumMDPTypeCodes)

        Dim objDT As DataTable
        If objDB.SysInfo.K1Configuration.DbVersion >= 11.04 Then
            objDT = objDB.GetDataTableBySQL( _
                "SELECT * FROM [" & clsDBConstants.Tables.cK1GROUPS & "]")

            For Each objRow As DataRow In objDT.Rows
                Dim eTypeCode As clsDBConstants.enumMDPTypeCodes = _
                    CType(objRow(clsDBConstants.Fields.K1Groups.cTYPECODE),  _
                    clsDBConstants.enumMDPTypeCodes)
                Dim intDefaultTypeID As Integer = CInt(clsDB.NullValue(objRow(clsDBConstants.Fields.K1Groups.cDEFAULTTYPEID), clsDBConstants.cintNULL))

                m_colDefaultTypes.Add(CStr(eTypeCode), intDefaultTypeID)
            Next
        End If

        objDT = objDB.GetDataTableBySQL( _
            "SELECT * FROM [" & clsDBConstants.Tables.cK1MDPTYPES & "] ORDER BY " & _
                "[" & clsDBConstants.Fields.K1MDPTypes.cTYPECODE & "]")

        Dim intPrevType As Integer = -1
        Dim colIDs As Hashtable = Nothing
        For Each objRow As DataRow In objDT.Rows
            Dim eTypeCode As clsDBConstants.enumMDPTypeCodes = _
                CType(objRow(clsDBConstants.Fields.K1Groups.cTYPECODE),  _
                clsDBConstants.enumMDPTypeCodes)
            Dim intTypeID As Integer = CInt(objRow(clsDBConstants.Fields.cTYPEID))

            If Not intPrevType = eTypeCode Then
                colIDs = New Hashtable

                m_colIDValues.Add(CStr(eTypeCode), colIDs)

                intPrevType = eTypeCode
            End If

            If Not colIDs.ContainsKey(CStr(intTypeID)) Then
                colIDs.Add(CStr(intTypeID), intTypeID)
            End If

            If Not m_colTypeGroups.ContainsKey(CStr(intTypeID)) Then
                m_colTypeGroups.Add(CStr(intTypeID), eTypeCode)
            End If
        Next

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If
    End Sub
#End Region

#Region " Properties "

    Public Property TypeGroups() As FrameworkCollections.K1Dictionary(Of clsDBConstants.enumMDPTypeCodes)
        Get
            Return m_colTypeGroups
        End Get
        Set(ByVal value As FrameworkCollections.K1Dictionary(Of clsDBConstants.enumMDPTypeCodes))
            m_colTypeGroups = value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Function GetDefaultType(ByVal eMDPTypeCode As clsDBConstants.enumMDPTypeCodes) As Integer
        If m_colDefaultTypes.ContainsKey(CType(eMDPTypeCode, String)) Then
            Return m_colDefaultTypes(CType(eMDPTypeCode, String))
        Else
            Return clsDBConstants.cintNULL
        End If
    End Function

    Public Function GetTypeIDCollection(ByVal eMDPTypeCode As clsDBConstants.enumMDPTypeCodes) As Hashtable
        If m_colIDValues.ContainsKey(CType(eMDPTypeCode, String)) Then
            Return m_colIDValues(CType(eMDPTypeCode, String))
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                If Not m_colDefaultTypes Is Nothing Then
                    m_colDefaultTypes.Clear()
                    m_colDefaultTypes = Nothing
                End If

                If Not m_colIDValues Is Nothing Then
                    For Each colIDs As Hashtable In m_colIDValues.Values
                        colIDs.Clear()
                    Next
                    m_colIDValues.Clear()
                    m_colIDValues = Nothing
                End If

                If Not m_colTypeGroups Is Nothing Then
                    m_colTypeGroups.Clear()
                    m_colTypeGroups = Nothing
                End If
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
