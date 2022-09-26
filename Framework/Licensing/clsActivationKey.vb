Imports System.Xml.Linq

Namespace Licensing

    Public Class clsActivationKey

#Region " Members "

        Private m_intFileVersion As Integer
        Private m_strProductSerialNumber As String = ""
        Private m_strProductName As String = ""
        Private m_strCustomerNumber As String = ""
        Private m_strComputerName As String = ""
        Private m_strServerName As String = ""
        Private m_strDatabaseName As String = ""
        Private m_dtActivatedDate As Date
        Private m_strAdminEmail As String = ""
        Private m_strCustomerName As String = ""
        Private m_blnExists As Boolean = False
        Private m_eLicenceType As clsDBConstants.enumApplicationType

#End Region

#Region " Constructors "

        Public Sub New(ByVal objDB As clsDB, ByVal eLicenceType As clsDBConstants.enumApplicationType)
            Try
                Select Case eLicenceType
                    Case clsDBConstants.enumApplicationType.RecFindActivationKey
                        m_strProductName = "RecFind 6"

                    Case clsDBConstants.enumApplicationType.K1ActivationKey
                        m_strProductName = "K1"
                End Select

                Dim objLicenceFile As New clsLicenseFile(objDB, eLicenceType)
                Dim strKey As String = objLicenceFile.DecryptLicenseFile
                If objLicenceFile.Exists AndAlso Not String.IsNullOrEmpty(strKey) Then
                    LoadValues(strKey)
                ElseIf objLicenceFile.Exists Then
                    Throw New clsK1Exception("Activation key does not exist or is invalid.", True)
                End If
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Public Sub New(ByVal strFile As String)
            Try
                Dim strKey As String = Licensing.clsLicenseFile.DecryptLicenseFile(strFile)
                If Not String.IsNullOrEmpty(strKey) Then
                    LoadValues(strKey)
                End If
            Catch ex As Exception
                Throw
            End Try
        End Sub

#End Region

#Region " Properties "

        Public ReadOnly Property FileVersion() As Integer
            Get
                Return m_intFileVersion
            End Get
        End Property

        Public ReadOnly Property ProductSerialNumber() As String
            Get
                Return m_strProductSerialNumber
            End Get
        End Property

        Public ReadOnly Property ProductName() As String
            Get
                Return m_strProductName
            End Get
        End Property

        Public ReadOnly Property CustomerNumber() As String
            Get
                Return m_strCustomerNumber
            End Get
        End Property

        Public ReadOnly Property ComputerName() As String
            Get
                Return m_strComputerName
            End Get
        End Property

        Public ReadOnly Property ServerName() As String
            Get
                Return m_strServerName
            End Get
        End Property

        Public ReadOnly Property DatabaseName() As String
            Get
                Return m_strDatabaseName
            End Get
        End Property

        Public ReadOnly Property ActivatedDate() As Date
            Get
                Return m_dtActivatedDate
            End Get
        End Property

        Public ReadOnly Property AdminEmail() As String
            Get
                Return m_strAdminEmail
            End Get
        End Property

        Public ReadOnly Property CustomerName() As String
            Get
                Return m_strCustomerName
            End Get
        End Property

        Public ReadOnly Property Exists() As Boolean
            Get
                Return m_blnExists
            End Get
        End Property

        Public ReadOnly Property LicenceType() As clsDBConstants.enumApplicationType
            Get
                Return m_eLicenceType
            End Get
        End Property



#End Region

        Private Sub LoadValues(ByVal strPlainKey As String)
            Dim xmlFile As XDocument = XDocument.Parse(strPlainKey)

            Dim objActivationInfo As XElement = xmlFile.<ActivationKey>.First

            If objActivationInfo Is Nothing Then
                Throw New clsK1Exception("Invalid Activation File.", True)
            End If

            If ((objActivationInfo.Nodes.Count = 8) Or (objActivationInfo.Nodes.Count = 9)) Then
                m_intFileVersion = CType(objActivationInfo.@Version, Integer)
                m_strProductSerialNumber = objActivationInfo.<ProductSerialNumber>.Value
                m_strProductName = objActivationInfo.<ProductName>.Value
                m_strCustomerNumber = objActivationInfo.<CustomerNumber>.Value
                m_strComputerName = objActivationInfo.<ComputerName>.Value
                m_strServerName = objActivationInfo.<DBServerName>.Value
                m_strDatabaseName = objActivationInfo.<DatabaseName>.Value.Trim({""""c})
                m_strAdminEmail = objActivationInfo.<EmailAddress>.Value
                m_strCustomerName = objActivationInfo.<CompanyName>.Value
                Date.TryParse(objActivationInfo.<ActivatedDate>.Value, m_dtActivatedDate)

                Select Case GetApplicationType(m_strProductName)
                    Case clsDBConstants.enumApplicationType.K1
                        m_eLicenceType = clsDBConstants.enumApplicationType.K1ActivationKey

                    Case clsDBConstants.enumApplicationType.RecFind
                        m_eLicenceType = clsDBConstants.enumApplicationType.RecFindActivationKey

                    Case Else
                        Throw New clsK1Exception("Invalid Activation File.", True)
                End Select

                m_blnExists = True
            End If
        End Sub

        Public Function IsValid(ByVal objLicenceFileInfo As clsLicenceFileInfo,
                                ByVal dtSys As Date, ByVal objDB As clsDB) As Boolean
            Try
                Dim strServerName As String = Nothing
                Dim strDatabaseName As String = Nothing
                Dim strGroup As String = Nothing
                Dim strGroupOrServerName As String = Nothing
                Dim strInstanceName As String = Nothing
                strGroup = objDB.GetAvailabilityGroupName()

                objDB.GetDatabaseInfo(strServerName, strDatabaseName, "")

                strServerName = strServerName.Replace(" (Using Web Services)", "")
                strInstanceName = (strServerName & "\").Split("\"c)(1)
                strServerName = strServerName.Split("\"c)(0) '-- strip out instance part of servername

                If strServerName.Contains(",") Then strServerName = strServerName.Substring(0, strServerName.IndexOf(","))
                If strServerName.Contains(".") Then strServerName = strServerName.Substring(0, strServerName.IndexOf("."))

                'if the DB is in an SQL Availability group then use that name     
                If String.IsNullOrEmpty(strGroup) Then
                    strGroupOrServerName = objDB.CheckSQLServerName()
                Else
                    strGroupOrServerName = If(String.IsNullOrEmpty(strInstanceName), strGroup, strGroup & "\" & strInstanceName)
                End If
#If TRACE Then
                Trace.WriteVerbose($"objLicenceFileInfo.MustActivate: {objLicenceFileInfo.MustActivate} 
ProductSerialNumber: {ProductSerialNumber} = {objLicenceFileInfo.ProductSerialNumber}
CustomerNumber: {CustomerNumber} = {objLicenceFileInfo.CustomerNumber} 
HostName: {ServerName} = {objDB.CheckHostName(strServerName)}
HostName: {ServerName} = {strGroupOrServerName}
GroupName: {strGroup} 
InstanceName: {strInstanceName} 
GroupOrServerName: {strGroupOrServerName} 
SQLServerName: {Me.ServerName} = {objDB.CheckSQLServerName()}
DatabaseName: {Me.DatabaseName} = {strDatabaseName}
", "")
#End If

                If (Not objLicenceFileInfo.MustActivate _
                OrElse String.Compare(Me.ProductSerialNumber, objLicenceFileInfo.ProductSerialNumber, True) = 0) _
                AndAlso String.Compare(Me.CustomerNumber, objLicenceFileInfo.CustomerNumber, True) = 0 _
                AndAlso ServerCompareName(Me.ServerName, objDB.CheckHostName(strServerName), strInstanceName, strGroupOrServerName) _
                AndAlso String.Compare(Me.DatabaseName, strDatabaseName, True) = 0 Then
                    '*************************************************
                    'JAMES - REMOVED DATE CHECKING BECAUSE ONE IS US FORMAT AND OTHER IS AU FORMAT. CAUSING ISSUES WHEN COMPARING
                    '*************************************************
                    'AndAlso Me.ActivatedDate = dtSys Then
                    'AndAlso String.Compare(Me.CustomerName, objLicenceFileInfo.CustomerName, True) = 0 _
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Compare a server name, use a group name if one exists
        ''' </summary>
        ''' <param name="strServerName"></param>
        ''' <param name="strServerCheckName"></param>
        ''' <param name="strGroupOrServerName"></param>
        ''' <returns></returns>
        ''' 
        Function ServerCompareName(strServerName As String, strServerCheckName As String, strInstance As String, strGroupOrServerName As String) As Boolean
            If String.Compare(strServerName, strServerCheckName, True) = 0 Then
                Return True
            End If
            If String.Compare(strServerName, strGroupOrServerName, True) = 0 Then
                Return True
            End If

            Dim s1 As String = strServerName
            Dim s2 As String = strServerCheckName
            Dim inst1 As String = If(s1.Contains("\"), s1.Substring(s1.IndexOf("\") + 1), "")

            Dim chs As String() = {",", "."}
            For Each ch As Char In chs
                s1 = StringUpto(s1, ch)
                s2 = StringUpto(s2, ch)
            Next
            s1 = String.Join("\", {s1, inst1})
            s2 = String.Join("\", {s2, strInstance})
            If String.Compare(s1, s2, True) = 0 Then
                Return True
            End If


            Return False
        End Function

        ''' <summary>
        ''' Return that portion of the string that precedes a character 
        ''' </summary>
        ''' <param name="strS">String to compare</param>
        ''' <param name="chC">Character that delimits the string you want</param>
        ''' <returns></returns>
        Function StringUpto(strS As String, chC As Char) As String
            If strS.Contains(chC) Then Return strS.Substring(0, strS.IndexOf(chC))
            Return strS
        End Function

    End Class


End Namespace