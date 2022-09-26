#Region " File Information "

'=====================================================================
' This class represents the table TypeFieldInfo in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date            Description
'---------------------------------------------------------------------
' KD        16/05/2005      Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsTypeField
    Inherits clsDBObjBase

#Region " Members "

    Private m_intFieldID As Integer
    Private m_blnIsVisible As Boolean
    Private m_blnMandatory As Boolean
    Private m_blnReadOnly As Boolean
    Private m_intCaptionID As Integer
    Private m_objCaption As clsCaption
    Private m_intAppliesToTypeID As Integer
    Private m_objAutoFillInfo As clsAutoFillInfo
    Private m_intSortOrder As Integer
    Private m_blnIsMultipleSequenceField As Boolean
    Private m_blnDeterminesMultipleSequence As Boolean
    Private m_blnAllowFreeTextEntry As Boolean
#End Region

#Region " Constructors "

    Private Sub New()
        MyBase.New()
        m_blnIsVisible = True
        m_blnReadOnly = False
        m_intCaptionID = clsDBConstants.cintNULL
        m_intAppliesToTypeID = clsDBConstants.cintNULL
        m_intSortOrder = 5
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_blnIsVisible = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cISVISIBLE, True), Boolean)
        m_intCaptionID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cCAPTIONID, clsDBConstants.cintNULL), Integer)
        m_intAppliesToTypeID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cAPPLIESTOTYPEID, clsDBConstants.cintNULL), Integer)
        Dim eAutoFillType As clsDBConstants.enumAutoFillTypes
        Dim strAutoFillValue As String
        Dim intAutoNumberFormatID As Integer
        eAutoFillType = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cAUTOFILLTYPE, clsDBConstants.cintNULL), clsDBConstants.enumAutoFillTypes)
        strAutoFillValue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cAUTOFILLVALUE, clsDBConstants.cstrNULL), String)
        intAutoNumberFormatID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cAUTONUMBERFORMATID, clsDBConstants.cintNULL), Integer)
        m_objAutoFillInfo = New clsAutoFillInfo(objDB, eAutoFillType, strAutoFillValue, intAutoNumberFormatID)
        m_intSortOrder = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cSORTORDER, clsDBConstants.cintNULL), Integer)
        m_blnIsMultipleSequenceField = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cISMULTIPLESEQUENCEFIELD, False), Boolean)
        m_blnDeterminesMultipleSequence = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cDETERMINESMULTIPLESEQUENCE, False), Boolean)
        m_blnAllowFreeTextEntry = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cALLOWFREETEXTENTRY, False), Boolean)
        m_blnMandatory = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cISMANDATORY, False), Boolean)
        m_blnReadOnly = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cISREADONLY, False), Boolean)
        m_intFieldID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.TypeFieldInfo.cFIELDID, clsDBConstants.cintNULL), Integer)
    End Sub

    Public Sub New(ByVal objField As clsField, ByVal intID As Integer, ByVal strExternalID As String, _
    ByVal intSecurityID As Integer, ByVal intAppliesToTypeID As Integer, ByVal blnIsVisible As Boolean, _
    ByVal intCaptionID As Integer, ByVal intSortOrder As Integer, ByVal objAutoFill As clsAutoFillInfo, _
    ByVal blnIsMultipleSequenceField As Boolean, ByVal blnDeterminesMultipleSequence As Boolean, _
    ByVal blnAllowFreeTextEntry As Boolean, ByVal blnMandatory As Boolean, ByVal blnReadOnly As Boolean)
        MyBase.New(objField.Database, intID, strExternalID, intSecurityID, clsDBConstants.cintNULL)
        m_intAppliesToTypeID = intAppliesToTypeID
        m_intCaptionID = intCaptionID
        m_objAutoFillInfo = objAutoFill
        m_intFieldID = objField.ID
        m_intSortOrder = intSortOrder
        m_blnIsVisible = blnIsVisible
        m_blnIsMultipleSequenceField = blnIsMultipleSequenceField
        m_blnDeterminesMultipleSequence = blnDeterminesMultipleSequence
        m_blnAllowFreeTextEntry = blnAllowFreeTextEntry
        m_blnMandatory = blnMandatory
        m_blnReadOnly = blnReadOnly
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property CaptionText() As String
        Get
            Return m_objDB.SysInfo.GetCaptionString(m_intCaptionID)
        End Get
    End Property

    Public ReadOnly Property Caption() As clsCaption
        Get
            If m_objCaption Is Nothing Then
                If Not m_intCaptionID = clsDBConstants.cintNULL Then
                    m_objCaption = clsCaption.GetItem(m_intCaptionID, Me.Database)
                End If
            End If
            Return m_objCaption
        End Get
    End Property

    Public ReadOnly Property IsVisible() As Boolean
        Get
            Return m_blnIsVisible
        End Get
    End Property

    Public ReadOnly Property IsMandatory() As Boolean
        Get
            Return m_blnMandatory
        End Get
    End Property

    Public ReadOnly Property IsReadOnly() As Boolean
        Get
            Return m_blnReadOnly
        End Get
    End Property

    Public ReadOnly Property AutoFillInfo() As clsAutoFillInfo
        Get
            Return m_objAutoFillInfo
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            Return m_intAppliesToTypeID
        End Get
    End Property

    Public ReadOnly Property SortOrder() As Integer
        Get
            Return m_intSortOrder
        End Get
    End Property

    Public ReadOnly Property IsMultipleSequenceField() As Boolean
        Get
            Return m_blnIsMultipleSequenceField
        End Get
    End Property

    Public ReadOnly Property DeterminesMultipleSequence() As Boolean
        Get
            Return m_blnDeterminesMultipleSequence
        End Get
    End Property

    Public ReadOnly Property AllowFreeTextEntry() As Boolean
        Get
            Return m_blnAllowFreeTextEntry
        End Get
    End Property

    Public ReadOnly Property FieldID() As Integer
        Get
            Return m_intFieldID
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsTypeField
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cTYPEFIELDINFO, intID)

            Return New clsTypeField(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTypeField)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTypeField)
        Dim strSP As String
        Try
            strSP = clsDBConstants.Tables.cTYPEFIELDINFO & clsDBConstants.StoredProcedures.cGETLIST
            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTypeField)
            For Each objDR As DataRow In objDT.Rows
                Dim objTypeFieldInfo As New clsTypeField(objDR, objDB)
                colObjects.Add(objTypeFieldInfo.FieldID & "_" & CType(objTypeFieldInfo.m_intAppliesToTypeID, String), objTypeFieldInfo)
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

    Public Shared Function GetList(ByVal objField As clsField, _
    ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsTypeField)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsTypeField)

        Try
            Dim objDT As DataTable = objDB.GetList(clsDBConstants.Tables.cTYPEFIELDINFO, _
                clsDBConstants.Fields.TypeFieldInfo.cFIELDID, objField.ID)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsTypeField)
            For Each objDR As DataRow In objDT.Rows
                Dim objTypeFieldInfo As New clsTypeField(objDR, objDB)
                colObjects.Add(CType(objTypeFieldInfo.m_intAppliesToTypeID, String), objTypeFieldInfo)
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

#Region " InsertUpdate "

    Public Sub InsertUpdate()
        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection( _
            m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTYPEFIELDINFO), m_intID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cAPPLIESTOTYPEID, m_intAppliesToTypeID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cFIELDID, m_intFieldID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cISVISIBLE, m_blnIsVisible)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cISMANDATORY, m_blnMandatory)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cISREADONLY, m_blnReadOnly)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cSORTORDER, m_intSortOrder)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cCAPTIONID, m_intCaptionID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cISMULTIPLESEQUENCEFIELD, m_blnIsMultipleSequenceField)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cDETERMINESMULTIPLESEQUENCE, m_blnDeterminesMultipleSequence)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cALLOWFREETEXTENTRY, m_blnAllowFreeTextEntry)

        If Not m_objAutoFillInfo Is Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cAUTOFILLTYPE, m_objAutoFillInfo.FillType)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cAUTOFILLVALUE, m_objAutoFillInfo.FillValue)
            If Not m_objAutoFillInfo.AutoNumberFormat Is Nothing Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.TypeFieldInfo.cAUTONUMBERFORMATID, _
                    m_objAutoFillInfo.AutoNumberFormat.ID)
            End If
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLTYPE, Nothing)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTOFILLVALUE, Nothing)
            colMasks.UpdateMaskObj(clsDBConstants.Fields.Field.cAUTONUMBERFORMATID, Nothing)
        End If

        If m_intID = clsDBConstants.cintNULL Then
            m_intID = colMasks.Insert(m_objDB)
        Else
            colMasks.Update(m_objDB)
        End If
    End Sub
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_objCaption Is Nothing Then
                m_objCaption.Dispose()
                m_objCaption = Nothing
            End If

            If Not m_objAutoFillInfo Is Nothing Then
                m_objAutoFillInfo.Dispose()
                m_objAutoFillInfo = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
