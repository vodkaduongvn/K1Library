Public Class clsCheckIn

#Region " Members "

    Private m_objDB As clsDB
    Private m_objTable As clsTable
    Private m_strFileName As String
    Private m_intType As Integer
    Private m_intSecurity As Integer
    Private m_intSelectedMDP As Integer
    Private m_enumCheckInType As enumCheckInType
    Private m_strCustomText As String
    Private m_blnDelete As Boolean
    Private m_intID As Integer
    Private m_enumApplicationType As clsDBConstants.enumApplicationType
    Private m_strIPAddress As String
    Private m_strAuthor As String
    Private m_strAbstract As String

#End Region

#Region " Properties "
    Public Property CheckInType() As enumCheckInType
        Get
            Return m_enumCheckInType
        End Get
        Set(ByVal value As enumCheckInType)
            m_enumCheckInType = value
        End Set
    End Property

    Public Property CustomText() As String
        Get
            Return m_strCustomText
        End Get
        Set(ByVal value As String)
            m_strCustomText = value
        End Set
    End Property

    Public Property Security() As Integer
        Get
            Return m_intSecurity
        End Get
        Set(ByVal value As Integer)
            m_intSecurity = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return m_intType
        End Get
        Set(ByVal value As Integer)
            m_intType = value
        End Set
    End Property

    Public Property ID() As Integer
        Get
            Return m_intID
        End Get
        Set(ByVal value As Integer)
            m_intID = value
        End Set
    End Property

    Public Property IPAddress() As String
        Get
            Return m_strIPAddress
        End Get
        Set(ByVal value As String)
            m_strIPAddress = value
        End Set
    End Property
#End Region

#Region " Enumerators "
    Public Enum enumCheckInType
        CHECKIN = 0
        NEWEDOC = 1
        NEWVERSION = 2
    End Enum
#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, ByVal strFileName As String, ByVal intSecurity As Integer, ByVal intType As Integer, ByVal intSelectedMDP As Integer, ByVal blnDelete As Boolean, ByVal enumApplicationType As clsDBConstants.enumApplicationType)
        m_objDB = objDB
        m_objTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cEDOC)
        m_strFileName = strFileName
        m_intType = intType
        m_intSecurity = intSecurity
        m_intSelectedMDP = intSelectedMDP
        m_blnDelete = blnDelete
        m_strCustomText = Nothing
        m_enumCheckInType = Nothing
        m_intID = 0
        m_enumApplicationType = enumApplicationType
        m_strIPAddress = Nothing
    End Sub

    Public Sub New(ByVal objDB As clsDB, ByVal strFileName As String, ByVal enumApplicationType As clsDBConstants.enumApplicationType)
        m_objDB = objDB
        m_objTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cEDOC)
        m_strFileName = strFileName
        m_intType = clsDBConstants.cintNULL
        m_intSecurity = clsDBConstants.cintNULL
        m_intSelectedMDP = clsDBConstants.cintNULL
        m_blnDelete = False
        m_strCustomText = Nothing
        m_enumCheckInType = Nothing
        m_intID = 0
        m_enumApplicationType = enumApplicationType
        m_strIPAddress = Nothing
    End Sub
#End Region

#Region " Methods "
    Public Sub CheckIn()
        If m_enumCheckInType = Nothing Then
            m_enumCheckInType = GetCheckInType()
        End If

        Dim colMasks As clsMaskFieldDictionary
        Select Case m_enumCheckInType
            Case enumCheckInType.CHECKIN
                colMasks = clsMaskField.CreateMaskCollection(m_objTable, GetInternalID(m_strFileName), True)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCHECKEDOUT, False)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCHECKEDINPERSONID, m_objDB.Profile.PersonID)
                colMasks.Update(m_objDB)
            Case enumCheckInType.NEWEDOC
                addNewEDOC()
            Case enumCheckInType.NEWVERSION
                Dim intPreviousID As Integer = GetInternalID(m_strFileName)

                colMasks = clsMaskField.CreateMaskCollection(m_objTable, clsTableMask.enumMaskType.MODIFY, intPreviousID)

                m_intSecurity = CInt(colMasks.GetMaskValue(clsDBConstants.Fields.cSECURITYID, m_intSecurity))
                m_intType = CInt(colMasks.GetMaskValue(clsDBConstants.Fields.cTYPEID, m_intType))

                '2016/09/29 -- James -- Fix for #1600003210 to include Author and Abstract from Previous EDOC
                m_strAuthor = CStr(colMasks.GetMaskValue(clsDBConstants.Fields.EDOC.cAUTHOR, m_strAuthor))
                m_strAbstract = CStr(colMasks.GetMaskValue(clsDBConstants.Fields.EDOC.cABSTRACT, m_strAuthor))

                addNewEDOC(intPreviousID)

                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCHECKEDOUT, False)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCHECKEDINPERSONID, m_objDB.Profile.PersonID)
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cISLATESTVERSION, False)
                colMasks.Update(m_objDB)
        End Select

        If m_blnDelete Then
            File.Delete(m_strFileName)
        End If
    End Sub

    Private Sub addNewEDOC(Optional ByVal intPreviousID As Integer = clsDBConstants.cintNULL)
        Dim strNewFileName As String
        'Rename file to remove ID and create a copy
        If m_intID <> 0 Then
            strNewFileName = RemoveInternalID(m_strFileName)
            File.Copy(m_strFileName, strNewFileName, True)
        Else
            strNewFileName = m_strFileName
        End If

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(m_objTable, clsTableMask.enumMaskType.ADD, clsDBConstants.cintNULL, m_intType)

        clsEDOC.UpdateEDOCFields(m_objDB, colMasks, strNewFileName, False, intPreviousID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurity)

        'CreateMaskCollection does not add the Type when not Type Dependent, so Type is set manually
        If Not m_intType = clsDBConstants.cintNULL Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cTYPEID, m_intType)
        End If

        '2016/09/29 -- James -- Fix for #1600003210 to include Author and Abstract from Previous EDOC
        If Not m_strAuthor = String.Empty Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cAUTHOR, m_strAuthor)
        End If

        If Not m_strAbstract = String.Empty Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cABSTRACT, m_strAbstract)
        End If

        If intPreviousID = clsDBConstants.cintNULL AndAlso m_intSelectedMDP <> clsDBConstants.cintNULL Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cMETADATAPROFILEID, m_intSelectedMDP)
        End If

        If m_strCustomText Is Nothing Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, IO.Path.GetFileName(strNewFileName))
        Else
            colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, m_strCustomText)
        End If

        colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cIMAGE, New Byte() {0})

        If m_enumApplicationType = clsDBConstants.enumApplicationType.RecFind Then
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCREATEDBYAPPLICATION, My.Application.Info.Title & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor)
        ElseIf m_enumApplicationType = clsDBConstants.enumApplicationType.WebClient Then
            If m_strIPAddress IsNot Nothing Then
                colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cORIGINALPATH, "[" & m_strIPAddress & "]")
            End If
            colMasks.UpdateMaskObj(clsDBConstants.Fields.EDOC.cCREATEDBYAPPLICATION, "Recfind 6 Web Client")
        End If

        Dim newID As Integer = colMasks.Insert(m_objDB)

        Dim objMaskField = m_objTable.Fields.Values.FirstOrDefault(Function(f) f.DatabaseName = clsDBConstants.Fields.EDOC.cIMAGE)
        m_objDB.WriteBLOB(m_objTable, objMaskField, newID, strNewFileName, False)

        'If the file originally included an ID, then delete the new file that was created
        If (m_enumApplicationType = clsDBConstants.enumApplicationType.RecFind OrElse m_enumApplicationType = clsDBConstants.enumApplicationType.WebClient) _
        AndAlso m_intID <> 0 Then
            File.Delete(strNewFileName)
        End If
    End Sub

    Public Function GetCheckInType() As enumCheckInType
        Dim objFileInfo As New FileInfo(m_strFileName)
        m_intID = GetInternalID(m_strFileName)
        If m_intID = 0 Then
            Return enumCheckInType.NEWEDOC
        End If

        Dim strFileName As String = IO.Path.GetFileName(RemoveInternalID(m_strFileName))
        Dim strSize As String = CStr(objFileInfo.Length)
        Dim strSuffix As String = Path.GetExtension(strFileName).Replace(".", "")
        'Changed Published Date to be Modified Date of file instead of Creation Date
        Dim dtPublishedDate As DateTime = CDate(objFileInfo.LastWriteTime)
        'Dim dtPublishedDate As DateTime = CDate(objFileInfo.CreationTime)

        Dim colMasks As clsMaskFieldDictionary
        Try
            colMasks = clsMaskField.CreateMaskCollection(m_objTable, m_intID, True)
        Catch ex As Exception
            Return enumCheckInType.NEWEDOC
        End Try

        If Not CBool(colMasks(clsDBConstants.Fields.EDOC.cCHECKEDOUT).Value1.Value) Then
            Return enumCheckInType.NEWEDOC
        End If

        If Not (strSuffix = CStr(colMasks(clsDBConstants.Fields.EDOC.cSUFFIX).Value1.Value)) Then
            Return enumCheckInType.NEWEDOC
        End If

        If Not (strSize = CStr(colMasks(clsDBConstants.Fields.EDOC.cSIZE).Value1.Value)) OrElse (Not (dtPublishedDate = CDate(colMasks(clsDBConstants.Fields.EDOC.cPUBLISHEDDATE).Value1.Value)) And (m_enumApplicationType = clsDBConstants.enumApplicationType.RecFind)) Then
            Return enumCheckInType.NEWVERSION
        End If

        If strFileName = CStr(colMasks(clsDBConstants.Fields.EDOC.cFILENAME).Value1.Value) Then
            Return enumCheckInType.CHECKIN
        Else
            Return enumCheckInType.NEWEDOC
        End If

    End Function

    Private Function GetInternalID(ByVal strFileName As String) As Integer
        Dim intID As Integer = 0

        Dim strFile As String = Path.GetFileNameWithoutExtension(strFileName)
        Dim blnTest As Boolean = strFile Like "*{*}"
        If blnTest Then
            Try
                intID = CInt(strFile.Substring(strFile.LastIndexOf("{") + 1, (strFile.Length - 1) - (strFile.LastIndexOf("{") + 1)))
            Catch
            End Try
        End If

        Return intID
    End Function

    Public Function RemoveInternalID(ByVal strFileName As String) As String
        Dim strExt As String = Path.GetExtension(strFileName)
        strFileName = strFileName.Substring(0, strFileName.LastIndexOf("{")) & strExt
        Return strFileName
    End Function

#End Region

End Class
