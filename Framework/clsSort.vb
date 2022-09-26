#Region " File Information "

'=====================================================================
' This class is used to implement sorting for the search screen
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date        Description
'---------------------------------------------------------------------
' KSD       ??/08/2004  Implemented.
'=====================================================================

#End Region

#End Region

Imports System.ComponentModel

Public Class clsSort
    Implements IDisposable

#Region " Members "

    Private m_intSortOrder As Integer
    Private m_objField As clsField
    Private m_blnAscending As Boolean
    Private m_blnDisposedValue As Boolean = False        ' To detect redundant calls
#End Region

#Region " Constructors "

    Public Sub New(ByVal objField As clsField, ByVal blnAscending As Boolean)
        m_objField = objField
        m_blnAscending = blnAscending
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' Used to sort fields (starts from lowest to highest)
    ''' </summary>
    Public Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
        Set(ByVal Value As Integer)
            m_intSortOrder = Value
        End Set
    End Property

    ''' <summary>
    ''' The field to sort on
    ''' </summary>
    Public Property Field() As clsField
        Get
            Return m_objField
        End Get
        Set(ByVal Value As clsField)
            m_objField = Value
        End Set
    End Property

    ''' <summary>
    ''' The direction with which to sort the field
    ''' </summary>
    Public Property Ascending() As Boolean
        Get
            Return m_blnAscending
        End Get
        Set(ByVal Value As Boolean)
            m_blnAscending = Value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Shared Function GetDefaultSortCollection(ByVal objTable As clsTable) As clsSortCollection
        Dim colSorts As New clsSortCollection
        Dim objField As clsField
        Dim blnIncludeIDField As Boolean = True

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@ID", objTable.ID))
        Dim strSQL As String = "SELECT DefaultSort.* FROM DefaultSort " & _
            "INNER JOIN Field ON DefaultSort.FieldID = Field.[ID] " & _
            "WHERE Field.TableID = @ID"

        Dim objDT As DataTable = objTable.Database.GetDataTableBySQL(strSQL, colParams)
        colParams.Dispose()

        If objDT.Rows.Count = 0 Then
            objField = objTable.Database.SysInfo.Fields( _
                objTable.ID & "_" & clsDBConstants.Fields.cEXTERNALID)
            colSorts.Add(New clsSort(objField, True))
        Else
            objDT.DefaultView.Sort = "SortOrder"

            For intLoop As Integer = 0 To objDT.DefaultView.Count - 1
                objField = objTable.Database.SysInfo.Fields( _
                    CInt(objDT.DefaultView(intLoop)("FieldID")))
                colSorts.Add(New clsSort(objField, CBool(objDT.DefaultView(intLoop)("IsAscending"))))
                If objField.DatabaseName.ToUpper = clsDBConstants.Fields.cID.ToUpper Then
                    blnIncludeIDField = False
                End If
            Next
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If

        If blnIncludeIDField Then
            objField = objTable.Database.SysInfo.Fields( _
                objTable.ID & "_" & clsDBConstants.Fields.cID)
            colSorts.Add(New clsSort(objField, True))
        End If

        Return colSorts
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objField = Nothing
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
