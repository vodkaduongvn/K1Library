#Region " File Information "

'=====================================================================
' This class represents the table AutoFill in the Database.
' It is used to automatically assign values to certain mask fields.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Kevin      11/10/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsAutoFillInfo
    Implements IDisposable

#Region " Members "

    Private m_eFillType As clsDBConstants.enumAutoFillTypes
    Private m_strFillValue As String
    Private m_intAutoNumberFormatID As Integer
    Private m_objAutoNumberFormat As clsAutoNumberFormat
    Private m_objDB As clsDB
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, ByVal eFillType As clsDBConstants.enumAutoFillTypes, _
    ByVal strFillValue As String, ByVal intAutoNumberFormatID As Integer)
        m_objDB = objDB
        m_eFillType = eFillType
        m_strFillValue = strFillValue
        m_intAutoNumberFormatID = intAutoNumberFormatID
    End Sub
#End Region

#Region " Properties "

    Public Property FillType() As clsDBConstants.enumAutoFillTypes
        Get
            Return m_eFillType
        End Get
        Set(ByVal Value As clsDBConstants.enumAutoFillTypes)
            m_eFillType = Value
        End Set
    End Property

    Public Property AutoNumberFormat() As clsAutoNumberFormat
        Get
            If m_objAutoNumberFormat Is Nothing Then
                If Not m_intAutoNumberFormatID = clsDBConstants.cintNULL Then
                    m_objAutoNumberFormat = clsAutoNumberFormat.GetItem(m_intAutoNumberFormatID, m_objDB)
                End If
            End If
            Return m_objAutoNumberFormat
        End Get
        Set(ByVal Value As clsAutoNumberFormat)
            m_objAutoNumberFormat = Value
        End Set
    End Property

    Public ReadOnly Property FillValue() As String
        Get
            Return m_strFillValue
        End Get
    End Property

    Public ReadOnly Property AutoNumberFormatID() As Integer
        Get
            Return m_intAutoNumberFormatID
        End Get
    End Property
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing

                If Not m_objAutoNumberFormat Is Nothing Then
                    m_objAutoNumberFormat.Dispose()
                    m_objAutoNumberFormat = Nothing
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
