Public Class clsDRMTrigger
    Inherits clsDRMBase

#Region " Members "

    Dim m_strTableName As String
    Dim m_strDatabaseName As String
    Dim m_strBody As String
    Dim m_strTriggerAction As String
    Dim m_blnOnInsert As Boolean
    Dim m_blnOnUpdate As Boolean
    Dim m_blnOnDelete As Boolean

#End Region

#Region " Constructors "

#Region " New "

    ''' <summary>
    ''' Create a new trigger
    ''' </summary>
    ''' <param name="objDB"></param>
    ''' <param name="strTableName"></param>
    ''' <param name="strTriggerName"></param>
    ''' <param name="strBody"></param>
    ''' <param name="strTriggerAction"></param>
    ''' <param name="blnOnInsert"></param>
    ''' <param name="blnOnUpdate"></param>
    ''' <param name="blnOnDelete"></param>
    ''' <param name="strExternalID"></param>
    ''' <param name="intSecurityID"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objDB As clsDB, ByVal strTableName As String, ByVal strTriggerName As String,
    ByVal strBody As String, ByVal strTriggerAction As String, ByVal blnOnInsert As Boolean,
    ByVal blnOnUpdate As Boolean, ByVal blnOnDelete As Boolean, ByVal strExternalID As String,
    ByVal intSecurityID As Integer)
        MyBase.New(objDB, strExternalID, intSecurityID, clsDBConstants.cintNULL)

        m_strDatabaseName = strTriggerName
        m_strTableName = strTableName
        m_strBody = strBody
        m_strTriggerAction = strTriggerAction
        m_blnOnInsert = blnOnInsert
        m_blnOnUpdate = blnOnUpdate
        m_blnOnDelete = blnOnDelete
    End Sub

#End Region

#Region " Existing "

    ''' <summary>
    ''' Get an existing Trigger
    ''' </summary>
    ''' <param name="objDB"></param>
    ''' <param name="strTriggerName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal objDB As clsDB, ByVal strTriggerName As String)
        MyBase.New(objDB, clsDBConstants.cstrNULL, clsDBConstants.cintNULL, clsDBConstants.cintNULL)

        Dim objTable As clsTable
        Dim dicMaskField As clsMaskFieldDictionary

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@Name", strTriggerName))
        Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT [ID] FROM [Trigger] " &
            "WHERE [DatabaseName] = @Name", colParams)
        colParams.Dispose()

        If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
            objTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cTRIGGER)
            dicMaskField = CType(clsMaskField.CreateMaskCollection(objTable, CInt(objDT.Rows(0)(0))), clsMaskFieldDictionary)

            m_intID = CInt(dicMaskField.GetMaskValue(clsDBConstants.Fields.cID, clsDBConstants.cintNULL))
            m_strExternalID = CType(dicMaskField.GetMaskValue(clsDBConstants.Fields.cEXTERNALID, strTriggerName), String)
            m_intTypeID = CInt(dicMaskField.GetMaskValue(clsDBConstants.Fields.cTYPEID, clsDBConstants.cintNULL))
            m_intSecurityID = CInt(dicMaskField.GetMaskValue(clsDBConstants.Fields.cSECURITYID, clsDBConstants.cintNULL))
            m_strDatabaseName = CType(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cDATABASENAME, strTriggerName), String)
            m_strBody = CType(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cSQL, clsDBConstants.cstrNULL), String)
            m_blnOnInsert = CBool(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cONINSERT, False))
            m_blnOnUpdate = CBool(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cONUPDATE, False))
            m_blnOnDelete = CBool(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cONDELETE, False))
            m_strTriggerAction = CType(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cTRIGGERACTION, clsDBConstants.cstrNULL), String)

            objTable = CType(objDB.SysInfo.Tables(CInt(dicMaskField.GetMaskValue(clsDBConstants.Fields.Trigger.cTABLEID))), clsTable)
            m_strTableName = objTable.DatabaseName
        Else
            Throw New Exception("Trigger does not exist.")
        End If

        '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
        If objDT IsNot Nothing Then
            objDT.Dispose()
            objDT = Nothing
        End If
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal objTrigger As clsTrigger)
        MyBase.New(objDB, objTrigger)

        m_strDatabaseName = objTrigger.DatabaseName
        m_strBody = objTrigger.SQL
        m_blnOnInsert = objTrigger.OnInsert
        m_blnOnUpdate = objTrigger.OnUpdate
        m_blnOnDelete = objTrigger.OnDelete
        If objTrigger.Action = clsTrigger.enumTriggerAction.INSTEAD_OF Then
            m_strTriggerAction = "I"
        Else
            m_strTriggerAction = "F"
        End If
        m_strTableName = objTrigger.Table.DatabaseName
    End Sub
#End Region

#End Region

#Region " Properties "

    Public Property Trigger() As clsTrigger
        Get
            Return CType(m_objDBObj, clsTrigger)
        End Get
        Set(ByVal value As clsTrigger)
            m_objDBObj = value
        End Set
    End Property

    Public Property TriggerAction() As String
        Get
            Return m_strTriggerAction
        End Get
        Set(ByVal value As String)
            m_strTriggerAction = value
        End Set
    End Property

    Public Property OnInsert() As Boolean
        Get
            Return m_blnOnInsert
        End Get
        Set(ByVal value As Boolean)
            m_blnOnInsert = value
        End Set
    End Property

    Public Property OnUpdate() As Boolean
        Get
            Return m_blnOnUpdate
        End Get
        Set(ByVal value As Boolean)
            m_blnOnUpdate = value
        End Set
    End Property

    Public Property OnDelete() As Boolean
        Get
            Return m_blnOnDelete
        End Get
        Set(ByVal value As Boolean)
            m_blnOnDelete = value
        End Set
    End Property

    Public Property Body() As String
        Get
            Return m_strBody
        End Get
        Set(ByVal value As String)
            m_strBody = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Insert/Update Trigger "

    ''' <summary>
    ''' Inserts or updates a trigger in K1 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InsertUpdate()
        Dim blnCreatedTransaction As Boolean = False
        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            Dim objTable As clsTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTRIGGER)
            Dim colMaskObjs As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(objTable, m_intID)

            m_objDB.DropTrigger(m_strDatabaseName)

            UpdateMaskCollection(colMaskObjs)

            If Trigger Is Nothing AndAlso m_intID = clsDBConstants.cintNULL Then
                '-- New trigger so insert
                m_intID = colMaskObjs.Insert(m_objDB)
                Trigger = clsTrigger.GetItem(m_intID, m_objDB)
            Else
                '-- Update existing trigger
                colMaskObjs.Update(m_objDB)
            End If

            SystemDB.CreateTrigger(colMaskObjs)

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

#End Region

#Region " Delete Trigger "

    ''' <summary>
    ''' Deletes a trigger from K1 and the database
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Delete()
        Dim blnCreatedTransaction As Boolean = False

        Try
            If Not m_objDB.HasTransaction Then
                blnCreatedTransaction = True
                m_objDB.BeginTransaction()
            End If

            If Not m_intID = clsDBConstants.cintNULL Then
                clsDRMBase.RecurseDeleteRelatedRecords(m_objDB, clsDBConstants.Tables.cTRIGGER, m_intID)

                SystemDB.DropTrigger(m_strDatabaseName)
            Else
                Throw New Exception("Trigger does not exist in the database, could not delete trigger '" & m_strDatabaseName & "'")
            End If

            If blnCreatedTransaction Then m_objDB.EndTransaction(True)
        Catch ex As Exception
            If blnCreatedTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

#End Region

#Region " Rename Trigger "

    ''' <summary>
    ''' Renames a trigger in K1 and the database
    ''' </summary>
    ''' <param name="strNewName">New name of the trigger</param>
    ''' <param name="blnCreateTransaction">Will we create a transaction to rename the trigger</param>
    ''' <remarks>
    ''' sp_rename does not work with triggers (as it does not update syscomments) 
    ''' so they must be dropped and recreated when renaming.
    ''' </remarks>
    Public Sub Rename(ByVal strNewName As String, Optional ByVal blnCreateTransaction As Boolean = False)
        Try
            If blnCreateTransaction Then
                m_objDB.BeginTransaction()
            End If

            Dim objTable As clsTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cTRIGGER)
            Dim colMaskObjs As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(
                objTable, clsDBConstants.cintNULL)

            m_strDatabaseName = strNewName

            UpdateMaskCollection(colMaskObjs)

            If Not m_intID = clsDBConstants.cintNULL Then
                '-- Update existing trigger
                colMaskObjs.Update(m_objDB)
            Else
                Throw New Exception("Can only rename triggers that already exist in the database.")
            End If

            SystemDB.CreateTrigger(colMaskObjs)

            If blnCreateTransaction Then
                m_objDB.EndTransaction(True)
            End If
        Catch ex As Exception
            If blnCreateTransaction Then m_objDB.EndTransaction(False)

            Throw
        End Try
    End Sub

#End Region

#Region " Trigger Exists "

    ''' <summary>
    ''' Checks if a trigger exists in the trigger table of K1
    ''' </summary>
    ''' <param name="objDB"></param>
    ''' <param name="strTriggerName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Exists(ByVal objDB As clsDB, ByVal strTriggerName As String) As Boolean
        Try
            Dim objExist As Object = objDB.ExecuteScalar("SELECT 1 FROM [" & clsDBConstants.Tables.cTRIGGER & "] " &
                "WHERE [DatabaseName] = '" & strTriggerName & "'")

            If objExist Is Nothing Then
                Return False
            Else
                Return CBool(objExist)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Update Mask Collection "

    ''' <summary>
    ''' Updates a given clsMaskFieldDictionary with the members values of the trigger
    ''' </summary>
    ''' <param name="dicMaskField">clsMaskFieldDictionary that we want to update</param>
    ''' <remarks></remarks>
    Private Sub UpdateMaskCollection(ByVal dicMaskField As clsMaskFieldDictionary)
        '-- Get the tables id
        Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strTableName)
        Dim intTableID As Integer = objTable.ID

        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.cID, m_intID)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strExternalID)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intTypeID)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cDATABASENAME, m_strDatabaseName)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cSQL, m_strBody)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cONINSERT, m_blnOnInsert)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cONUPDATE, m_blnOnUpdate)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cONDELETE, m_blnOnDelete)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cTRIGGERACTION, m_strTriggerAction)
        dicMaskField.UpdateMaskObj(clsDBConstants.Fields.Trigger.cTABLEID, intTableID)
    End Sub

#End Region

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
    End Sub

#End Region

End Class
