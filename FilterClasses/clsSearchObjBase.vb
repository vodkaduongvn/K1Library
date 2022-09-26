#Region " File Information "

'==============================================================================
' This is the base class for a search group or search element
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       27/03/2007  Implemented.
'==============================================================================

#End Region

#End Region

<Serializable()> Public MustInherit Class clsSearchObjBase
    Implements IDisposable

#Region " Members "

    Private m_eOpType As clsSearchFilter.enumOperatorType
    Private m_blnDisposedValue As Boolean
#End Region

#Region " Constructors "

    Public Sub New()
    End Sub

    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType)
        m_eOpType = eOpType
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' AND, OR, ANDNOT, ORNOT, [Nothing]
    ''' </summary>
    Public Property OperatorType() As clsSearchFilter.enumOperatorType
        Get
            Return m_eOpType
        End Get
        Set(ByVal value As clsSearchFilter.enumOperatorType)
            m_eOpType = value
        End Set
    End Property
#End Region

#Region " IDisposable Support "

    Protected MustOverride Sub DisposeObject()

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
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
