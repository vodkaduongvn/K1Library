#Region " File Information "

'=====================================================================
' This class represents the table AutoNumberFormatMultipleSequence in
' the Database. It is used to maintain multiple sequence values.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' M.O.      03/02/2005    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsMultipleSequence
    Inherits clsDBObjBase

#Region " Members "

    Private m_intAutoNumberFormatID As Integer
    Private m_intMultipleSequenceLastValue As Integer
    Private m_intTitle1ID As Integer
    Private m_intTitle2ID As Integer
    Private m_intTitle3ID As Integer
    Private m_intTitle4ID As Integer
#End Region

#Region " Constructors "

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intAutoNumberFormatID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cAUTONUMBERFORMATID, clsDBConstants.cintNULL), Integer)
        m_intMultipleSequenceLastValue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cMULTIPLESEQUENCELASTVALUE, 0), Integer)
        m_intTitle1ID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cTITLE1, 0), Integer)
        m_intTitle2ID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cTITLE2, 0), Integer)
        m_intTitle3ID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cTITLE3, 0), Integer)
        m_intTitle4ID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cTITLE4, 0), Integer)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property AutoNumberFormatID() As Integer
        Get
            Return m_intAutoNumberFormatID
        End Get
    End Property

    Public ReadOnly Property MultipleSequenceLastValue() As Integer
        Get
            Return m_intMultipleSequenceLastValue
        End Get
    End Property

    Public ReadOnly Property Title1ID() As Integer
        Get
            Return m_intTitle1ID
        End Get
    End Property

    Public ReadOnly Property Title2ID() As Integer
        Get
            Return m_intTitle2ID
        End Get
    End Property

    Public ReadOnly Property Title3ID() As Integer
        Get
            Return m_intTitle3ID
        End Get
    End Property

    Public ReadOnly Property Title4ID() As Integer
        Get
            Return m_intTitle4ID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsMultipleSequence
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE, intID)

            Return New clsMultipleSequence(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objAutoNumberFormat As clsAutoNumberFormat, _
    ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsMultipleSequence)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsMultipleSequence)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE, _
                clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cAUTONUMBERFORMATID, objAutoNumberFormat.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsMultipleSequence)
            For Each objDataRow As DataRow In objDT.Rows
                Dim objItem As New clsMultipleSequence(objDataRow, objDB)
                colObjects.Add(objItem.ExternalID, objItem)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
