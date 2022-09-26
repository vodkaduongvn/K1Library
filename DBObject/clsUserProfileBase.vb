Public Class clsUserProfileBase
    Inherits clsDBObjBase

#Region " Members "

    Protected m_intLanguageID As Integer
    Protected m_objLanguage As clsLanguage
    Protected m_intBackgroundID As Integer
    Protected m_objBackground As clsBackground
    Protected m_intFontID As Integer
    Protected m_objFont As clsFont
    Protected m_intHeadingFontID As Integer
    Protected m_objHeadingFont As clsFont
    Protected m_intButtonFontID As Integer
    Protected m_objButtonFont As clsFont
    Protected m_blnEnableScan As Boolean = True
    Protected m_intDefaultLanguageID As Integer
    Protected m_blnDisplayTableList As Boolean
    Protected m_intPersonID As Integer
    Protected m_objPerson As clsPerson
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_intPersonID = clsDBConstants.cintNULL
        m_intLanguageID = clsDBConstants.cintNULL
        m_intBackgroundID = clsDBConstants.cintNULL
        m_intFontID = clsDBConstants.cintNULL
        m_blnEnableScan = True
        m_intDefaultLanguageID = clsDBConstants.cintNULL
        m_blnDisplayTableList = False
    End Sub

    Protected Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intLanguageID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cLANGUAGEID, clsDBConstants.cintNULL), Integer)
        m_intBackgroundID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cBACKGROUNDID, clsDBConstants.cintNULL), Integer)
        m_intFontID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cFONTID, clsDBConstants.cintNULL), Integer)
        m_intHeadingFontID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cHEADINGFONTID, clsDBConstants.cintNULL), Integer)
        m_intButtonFontID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cBUTTONFONTID, clsDBConstants.cintNULL), Integer)
        m_blnDisplayTableList = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cDISPLAYTABLELIST, False), Boolean)
        m_intPersonID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.UserProfile.cPERSONID, clsDBConstants.cintNULL), Integer)
        If objDB.SysInfo IsNot Nothing AndAlso _
        objDB.SysInfo.K1Configuration IsNot Nothing AndAlso _
        objDB.SysInfo.K1Configuration.IsDefaultProfileLoaded Then
            m_intDefaultLanguageID = objDB.SysInfo.K1Configuration.DefaultProfile.LanguageID
        Else
            m_intDefaultLanguageID = m_intLanguageID
        End If
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property PersonID() As Integer
        Get
            Return m_intPersonID
        End Get
    End Property

    Public ReadOnly Property Person() As clsPerson
        Get
            If m_objPerson Is Nothing Then
                If Not m_intPersonID = clsDBConstants.cintNULL Then
                    m_objPerson = clsPerson.GetItem(m_intPersonID, Me.Database)
                End If
            End If
            Return m_objPerson
        End Get
    End Property

    Public ReadOnly Property LanguageID() As Integer
        Get
            Return m_intLanguageID
        End Get
    End Property

    Public ReadOnly Property Language() As clsLanguage
        Get
            If m_objLanguage Is Nothing Then
                If Not m_intLanguageID = clsDBConstants.cintNULL Then
                    m_objLanguage = clsLanguage.GetItem(m_intLanguageID, Me.Database)
                End If
            End If
            Return m_objLanguage
        End Get
    End Property

    Public ReadOnly Property Background() As clsBackground
        Get
            If m_objBackground Is Nothing Then
                If Not m_intBackgroundID = clsDBConstants.cintNULL Then
                    m_objBackground = clsBackground.GetItem(m_intBackgroundID, Me.Database)
                End If
            End If
            Return m_objBackground
        End Get
    End Property

    Public ReadOnly Property Font() As clsFont
        Get
            If m_objFont Is Nothing Then
                If Not m_intFontID = clsDBConstants.cintNULL Then
                    m_objFont = clsFont.GetItem(m_intFontID, Me.Database)
                End If
            End If
            Return m_objFont
        End Get
    End Property

    Public ReadOnly Property HeadingFont() As clsFont
        Get
            If m_objHeadingFont Is Nothing Then
                If Not m_intHeadingFontID = clsDBConstants.cintNULL Then
                    m_objHeadingFont = clsFont.GetItem(m_intHeadingFontID, Me.Database)
                End If
            End If
            Return m_objHeadingFont
        End Get
    End Property

    Public ReadOnly Property ButtonFont() As clsFont
        Get
            If m_objButtonFont Is Nothing Then
                If Not m_intButtonFontID = clsDBConstants.cintNULL Then
                    m_objButtonFont = clsFont.GetItem(m_intButtonFontID, Me.Database)
                End If
            End If
            Return m_objButtonFont
        End Get
    End Property

    Public ReadOnly Property EnableScan() As Boolean
        Get
            Return m_blnEnableScan
        End Get
    End Property

    Public ReadOnly Property DefaultLanguageID() As Integer
        Get
            Return m_intDefaultLanguageID
        End Get
    End Property

    Public ReadOnly Property DisplayTableList() As Boolean
        Get
            Return m_blnDisplayTableList
        End Get
    End Property
#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsUserProfileBase
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter(clsDB.ParamName(clsDBConstants.Fields.cID), intID))

            Dim objDT As DataTable = objDB.GetDataTable( _
                clsDBConstants.Tables.cUSERPROFILE & clsDBConstants.StoredProcedures.cGETITEM, colParams)

            Return New clsUserProfileBase(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_objLanguage Is Nothing Then
                m_objLanguage.Dispose()
                m_objLanguage = Nothing
            End If

            If Not m_objBackground Is Nothing Then
                m_objBackground.Dispose()
                m_objBackground = Nothing
            End If

            If Not m_objFont Is Nothing Then
                m_objFont.Dispose()
                m_objFont = Nothing
            End If

            If Not m_objHeadingFont Is Nothing Then
                m_objHeadingFont.Dispose()
                m_objHeadingFont = Nothing
            End If

            If Not m_objButtonFont Is Nothing Then
                m_objButtonFont.Dispose()
                m_objButtonFont = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
