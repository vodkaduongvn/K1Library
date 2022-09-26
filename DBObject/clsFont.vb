#Region " File Information "

'=====================================================================
' This class represents the table Font in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsFont
	Inherits clsDBObjBase

#Region " Members "

    Private m_intSize As Integer
    Private m_blnIsItalic As Boolean
    Private m_strColor As String
    Private m_intColorID As Integer
    Private m_blnIsBold As Boolean
    Private m_strFontName As String
    Private m_blnHasDropShadow As Boolean
    Private m_strDropShadowColor As String
    Private m_intDropShadowColorID As Integer
    Private m_intDropShadowPixelOffset As Integer
#End Region

#Region " Constructors "

	Public Sub New()
		MyBase.New()
        m_intSize = clsDBConstants.cintNULL
        m_blnIsItalic = False
        m_intColorID = clsDBConstants.cintNULL
        m_blnIsBold = False
        m_strFontName = clsDBConstants.cstrNULL
        m_blnHasDropShadow = False
        m_intDropShadowPixelOffset = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intSize = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cSIZE, clsDBConstants.cintNULL), Integer)
        m_blnIsItalic = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cISITALIC, False), Boolean)
        m_intColorID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cCOLORID, clsDBConstants.cintNULL), Integer)
        m_blnIsBold = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cISBOLD, False), Boolean)
        m_strFontName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cFONTNAME, clsDBConstants.cstrNULL), String)
        m_blnHasDropShadow = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cHASDROPSHADOW, False), Boolean)
        m_intDropShadowColorID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cDROPSHADOWCOLORID, clsDBConstants.cintNULL), Integer)
        m_intDropShadowPixelOffset = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Font.cDROPSHADOWPIXELOFFSET, clsDBConstants.cintNULL), Integer)

        If m_intSize < 6 Then
            m_intSize = 6
        End If

        If m_intSize > 36 Then
            m_intSize = 36
        End If
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property Size() As Integer
        Get
            Return m_intSize
        End Get
    End Property

    Public ReadOnly Property IsItalic() As Boolean
        Get
            Return m_blnIsItalic
        End Get
    End Property

    Public ReadOnly Property Color() As String
        Get
            If m_strColor Is Nothing Then
                If Not m_intColorID = clsDBConstants.cintNULL Then
                    Dim objDT As DataTable = m_objDB.GetItem(clsDBConstants.Tables.cCOLOR, m_intColorID)
                    Dim objDB As New clsDBObject(objDT.Rows(0), m_objDB)
                    m_strColor = objDB.ExternalID
                    objDB.Dispose()
                Else
                    m_strColor = "White"
                End If
            End If
            Return m_strColor
        End Get
    End Property

    Public ReadOnly Property IsBold() As Boolean
        Get
            Return m_blnIsBold
        End Get
    End Property

    Public ReadOnly Property FontName() As String
        Get
            Return m_strFontName
        End Get
    End Property

    Public ReadOnly Property HasDropShadow() As Boolean
        Get
            Return m_blnHasDropShadow
        End Get
    End Property

    Public ReadOnly Property DropShadowColor() As String
        Get
            If m_strDropShadowColor Is Nothing Then
                If Not m_intDropShadowColorID = clsDBConstants.cintNULL Then
                    Dim objDT As DataTable = m_objDB.GetItem(clsDBConstants.Tables.cCOLOR, _
                        m_intDropShadowColorID)
                    Dim objDB As New clsDBObject(objDT.Rows(0), m_objDB)
                    m_strDropShadowColor = objDB.ExternalID
                    objDB.Dispose()
                End If
            End If
            Return m_strDropShadowColor
        End Get
    End Property

    Public ReadOnly Property DropShadowPixelOffset() As Integer
        Get
            Return m_intDropShadowPixelOffset
        End Get
    End Property
#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsFont
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cFONT, intID)

            Return New clsFont(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
