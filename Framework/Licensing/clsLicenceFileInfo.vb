Imports K1Library.clsDBConstants
Imports System.Xml.Linq

Namespace Licensing

    Public Class clsLicenceFileInfo

#Region " Members "

        Private m_intID As Integer
        Private m_eApplicationType As enumApplicationType
        Private m_blnMustActivate As Boolean

        Private m_strProductSerialNumber As String
        Private m_dtDateCreated As DateTime = DateTime.MinValue
        Private m_intNumberOfUsers As Integer
        Private m_intNumberOfMailboxes As Integer
        Private m_intTotalNumberOfDocuments As Integer
        Private m_intNumberOfRecords As Integer
        Private m_dtInvoiceDate As Date = Date.MinValue
        Private m_dtDropDeadDate As Date = Date.MinValue
        Private m_strComments As String
        Private m_strCustomerName As String
        Private m_strCustomerNumber As String
        Private m_intProductID As Integer
        Private m_strProductName As String
        Private m_blnIsDemo As Boolean = True
        Private m_blnIsValid As Boolean = True
        Private m_blnLicenceExists As Boolean = True
        Private m_blnTestLicence As Boolean = True
        Private m_strError As String
        Private m_colDocumentLincenses As New FrameworkCollections.K1Dictionary(Of clsDocumentLicence)
#End Region

#Region " Constants "

        Private Const cEncryptedCleanInstall As String = "HrjbVwO3FD5Um0SAa0kKvqtP0JkHj2YMiA8aKqOyDdw="
        'HrjbVwO3FD5Um0SAa0kKvqtP0JkHj2YMiA8aKqOyDdw=
#End Region

#Region " Constructors "

        Public Sub New(ByVal objDB As clsDB, ByVal eApplicationType As enumApplicationType)
            m_eApplicationType = eApplicationType

            If eApplicationType = enumApplicationType.K1 _
            OrElse eApplicationType = enumApplicationType.RecFind Then
                m_blnMustActivate = True
            End If

            RePopulate(objDB)

            If Not (eApplicationType = enumApplicationType.K1 _
            OrElse eApplicationType = enumApplicationType.RecFind) Then
                Dim objLic As New clsLicenceFileInfo(objDB, enumApplicationType.RecFind)

                If objLic IsNot Nothing Then
                    m_blnTestLicence = objLic.IsTestLicence
                End If
            End If

            If eApplicationType = enumApplicationType.Mini_API AndAlso
                objDB.Profile IsNot Nothing AndAlso
                Not objDB.Profile.IsSharePointUser Then

                Dim objLic As New clsLicenceFileInfo(objDB, enumApplicationType.SharePoint)
                If objLic IsNot Nothing Then
                    m_intNumberOfUsers = m_intNumberOfUsers - objLic.NumberOfUsers
                    If m_intNumberOfUsers < 0 Then
                        m_intNumberOfUsers = 0
                    End If
                End If

            End If

        End Sub

#End Region

#Region " Properties "

        Public ReadOnly Property ID() As Integer
            Get
                Return m_intID
            End Get
        End Property

        Public ReadOnly Property ApplicationType() As enumApplicationType
            Get
                Return m_eApplicationType
            End Get
        End Property

        Public ReadOnly Property MustActivate() As Boolean
            Get
                Return m_blnMustActivate
            End Get
        End Property

        Public ReadOnly Property ProductSerialNumber() As String
            Get
                Return m_strProductSerialNumber
            End Get
        End Property

        Public ReadOnly Property DateCreated() As DateTime
            Get
                Return m_dtDateCreated
            End Get
        End Property

        Public ReadOnly Property NumberOfUsers() As Integer
            Get
                Return m_intNumberOfUsers
            End Get
        End Property

        Public ReadOnly Property NumberOfMailboxes() As Integer
            Get
                Return m_intNumberOfMailboxes
            End Get
        End Property

        Public ReadOnly Property TotalNumberOfDocuments() As Integer
            Get
                Return m_intTotalNumberOfDocuments
            End Get
        End Property

        Public ReadOnly Property NumberOfRecords() As Integer
            Get
                Return m_intNumberOfRecords
            End Get
        End Property

        Public ReadOnly Property InvoiceDate() As Date
            Get
                Return m_dtInvoiceDate
            End Get
        End Property

        Public ReadOnly Property DropDeadDate() As Date
            Get
                Return m_dtDropDeadDate
            End Get
        End Property

        Public ReadOnly Property Comments() As String
            Get
                Return m_strComments
            End Get
        End Property

        Public ReadOnly Property CustomerName() As String
            Get
                Return m_strCustomerName
            End Get
        End Property

        Public ReadOnly Property CustomerNumber() As String
            Get
                Return m_strCustomerNumber
            End Get
        End Property

        Public ReadOnly Property ProductID() As Integer
            Get
                Return m_intProductID
            End Get
        End Property

        Public ReadOnly Property ProductName() As String
            Get
                Return m_strProductName
            End Get
        End Property

        Public ReadOnly Property IsTrainingVersion() As Boolean
            Get
                Return m_blnIsDemo
            End Get
        End Property

        Public ReadOnly Property HasExpired() As Boolean
            Get
                If Not m_dtDropDeadDate = Date.MinValue AndAlso m_dtDropDeadDate < Now Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Public ReadOnly Property IsValid() As Boolean
            Get
                Return m_blnIsValid
            End Get
        End Property

        Public ReadOnly Property LicenceExists() As Boolean
            Get
                Return m_blnLicenceExists
            End Get
        End Property

        Public ReadOnly Property [Error]() As String
            Get
                Return m_strError
            End Get
        End Property

        Public ReadOnly Property IsTestLicence() As Boolean
            Get
                Return m_blnTestLicence
            End Get
        End Property

        Public ReadOnly Property DocumentLicences() As FrameworkCollections.K1Dictionary(Of Licensing.clsDocumentLicence)
            Get
                Return m_colDocumentLincenses
            End Get
        End Property

#End Region

#Region " Methods "

        Private Sub RePopulate(ByVal objDB As clsDB)
            Try
                Dim objLicense As New clsLicenseFile(objDB, m_eApplicationType)

                If Not objLicense.Exists Then
                    m_blnLicenceExists = False
                    m_blnIsValid = False
                    m_strError = "This product is currently unlicensed."
                    Return
                End If

                Dim xmlFile As XDocument = objLicense.LoadXMLFromLicenseFile

                Dim objLicenceInfo As XElement = (From objNodes In xmlFile.Descendants("Table") _
                                                 Where objNodes.Parent IsNot Nothing _
                                                 Select objNodes).First

                If objLicenceInfo Is Nothing Then
                    m_blnIsValid = False
                    Throw New clsK1Exception(modErrors.ErrorNumber.No_Licence, _
                                             "Invalid Licence File.", True)
                End If

                m_intID = objLicense.ID
                m_strProductSerialNumber = objLicenceInfo.<ProductSerialNumber>.Value
                m_strProductName = objLicenceInfo.<ProductName>.Value
                m_blnLicenceExists = True

                If Not GetProductName(m_eApplicationType).ToLower = m_strProductName.ToLower Then
                    m_blnIsValid = False
                    m_strError = "Invalid " & m_strProductName & " Licence file."
                ElseIf Date.TryParse(objLicenceInfo.<DropDeadDate>.Value, m_dtDropDeadDate) AndAlso Me.HasExpired() Then
                    '-- Check if Licence has expired
                    m_blnIsValid = False
                    m_strError = m_strProductName & " Licence file has expired."
                End If

                Date.TryParse(objLicenceInfo.<DateCreated>.Value, m_dtDateCreated)
                Date.TryParse(objLicenceInfo.<InvoiceDate>.Value, m_dtInvoiceDate)
                m_strComments = objLicenceInfo.<Comments>.Value
                m_strCustomerName = objLicenceInfo.<CustomerName>.Value
                m_strCustomerNumber = objLicenceInfo.<CustomerNumber>.Value
                m_intProductID = CInt(objLicenceInfo.<ProductID>.Value)

                If Not m_eApplicationType = clsDBConstants.enumApplicationType.RecFind Then
                    Dim objLF As New clsLicenceFileInfo(objDB, clsDBConstants.enumApplicationType.RecFind)
                    If objLF Is Nothing OrElse Not objLF.CustomerNumber = m_strCustomerNumber Then
                        m_blnIsValid = False
                        m_strError = m_strProductName & " - customer number does not match RecFind 6."
                    End If
                End If

                Try
                    Dim strTestLicence As String = objLicenceInfo.<TestLicence>.Value
                    If strTestLicence IsNot Nothing AndAlso Not strTestLicence.ToUpper = "TRUE" Then
                        m_blnTestLicence = False
                    End If
                Catch ex As Exception

                End Try

                If CChar(objLicenceInfo.<LiveOrDemo>.Value) = "L"c Then
                    m_blnIsDemo = False
                Else
                    m_intNumberOfRecords = CType(objLicenceInfo.<NumberOfRecords>.Value, Integer)
                End If

                Select Case m_eApplicationType
                    Case enumApplicationType.GEM
                        m_intNumberOfMailboxes = CType(objLicenceInfo.<NumberOfMailboxes>.Value, Integer)

                    Case enumApplicationType.RecCapture
                        Dim dtNow As Date = objDB.GetCurrentTime

                        For Each objLicenceNode As XElement In objLicenceInfo.<Licences>.Descendants("Licence")
                            Dim intSerialNumber As Integer = CType(objLicenceNode.<SerialNumber>.Value, Integer)
                            Dim intNumberOfDocuments As Integer = CType(objLicenceNode.<NumDocuments>.Value, Integer)
                            Dim dtDateStart As Date = CType(objLicenceNode.<StartDate>.Value, Date)
                            Dim val = CInt(objLicenceNode.<Type>.Value)
                            Dim eDocLicType As clsDocumentLicence.enumDocLicType = CType(val, clsDocumentLicence.enumDocLicType)

                            If intSerialNumber > 0 AndAlso
                                (intNumberOfDocuments > 0 OrElse intNumberOfDocuments = Integer.MinValue) AndAlso
                                (Not eDocLicType = clsDocumentLicence.enumDocLicType.VALID_YEAR OrElse
                                 (dtNow < dtDateStart.AddYears(1))) AndAlso
                             Not m_colDocumentLincenses.ContainsKey(intSerialNumber.ToString()) Then

                                m_colDocumentLincenses.Add(intSerialNumber.ToString(),
                                                           New clsDocumentLicence(intSerialNumber,
                                                                                  dtDateStart,
                                                                                  intNumberOfDocuments,
                                                                                  eDocLicType))

                                If m_intTotalNumberOfDocuments = Integer.MinValue Then
                                    Continue For
                                End If

                                If intNumberOfDocuments = Integer.MinValue Then
                                    m_intTotalNumberOfDocuments = Integer.MinValue
                                Else
                                    m_intTotalNumberOfDocuments += intNumberOfDocuments
                                End If
                            End If
                        Next

                    Case Else
                        m_intNumberOfUsers = CType(objLicenceInfo.<NumberOfUsers>.Value, Integer)
                End Select
            Catch ex As Exception
                m_blnIsValid = False
                m_strError = ex.Message
            End Try
        End Sub

        Public Shared Function GetInstallationDate(ByVal objDB As clsDB) As Date
            Try
                Dim objEncryption As New clsEncryption(True, True)
                Dim objDT As DataTable = objDB.GetDataTable(clsDBConstants.StoredProcedures.cDRM_SYSTEMCHECK)
                Dim strSQL As String = String.Empty

                If Not objDT.Rows.Count = 1 Then
                    Throw New clsK1Exception("Installation validation has failed.", True)
                End If

                Dim strEncrypted As String = CType(objDT.Rows(0)("Value"), String)

                '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
                If objDT IsNot Nothing Then
                    objDT.Dispose()
                    objDT = Nothing
                End If

                If strEncrypted = cEncryptedCleanInstall Then
                    strEncrypted = objEncryption.Encrypt(Now.ToString("MMM dd, yyyy HH:mm:ss"))

                    objDB.BeginTransaction()

                    objDB.ExecuteSQL("If OBJECT_ID('dbo.drmSystemCheck') IS NOT NULL DROP PROCEDURE dbo.drmSystemCheck")

                    strSQL = "/*" & vbCrLf
                    strSQL &= "=====================================================================" & vbCrLf
                    strSQL &= "Description" & vbCrLf
                    strSQL &= "------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                    strSQL &= "This stored procedure is essential to RecFind 6." & vbCrLf
                    strSQL &= "DO NOT DELETE UNDER ANY CIRCUMSTANCES!" & vbCrLf
                    strSQL &= "=====================================================================*/" & vbCrLf & vbCrLf
                    strSQL &= "CREATE PROCEDURE dbo.drmSystemCheck AS" & vbCrLf & vbCrLf
                    strSQL &= "SELECT '" & strEncrypted & "' AS [Value]"

                    objDB.ExecuteSQL(strSQL)

                    objDB.EndTransaction(True)

                    'Throw New clsK1Exception("The database has not been registered and cannot run. " & _
                    '                         "Please register using the Registration Wizard.", True)
                End If

                Try
                    Return CType(objEncryption.Decrypt(strEncrypted), Date)
                Catch ex As Exception
                    Throw New clsK1Exception("Registration validation has failed.", True)
                End Try
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Shared Function CheckProperInstallation(ByRef objDB As clsDB, _
                                                       ByVal eApplicationType As enumApplicationType, _
                                                       ByRef strMessage As String) As Boolean
            Try
                Dim objLicenceFileInfo As New clsLicenceFileInfo(objDB, eApplicationType)

                If Not objLicenceFileInfo.IsValid Then
                    Throw New clsK1Exception(objLicenceFileInfo.Error, True)
                End If

                Dim dtSys As Date = GetInstallationDate(objDB)
                Dim objActivationKey As clsActivationKey
                If objLicenceFileInfo.MustActivate Then
                    objActivationKey = New clsActivationKey(objDB, GetActivationType(eApplicationType))
                Else
                    objActivationKey = New clsActivationKey(objDB, enumApplicationType.RecFindActivationKey)
                End If

                If objLicenceFileInfo.IsTestLicence Then
                    strMessage = "You are currently using a test/training licence. " & _
                        "This licence is not to be used for live production systems."
                End If

                If objActivationKey.Exists AndAlso objActivationKey.IsValid(objLicenceFileInfo, dtSys, objDB) Then
                    'K1 Activated
                    If objLicenceFileInfo.IsTrainingVersion Then
                        objDB.RecordLimit = objLicenceFileInfo.NumberOfRecords
                    End If
                    Return True
                End If

                Try
                    Dim intTotalDays As Integer = objDB.GetCurrentTime.Subtract(dtSys).Days
                    If intTotalDays > 45 Then
                        Throw New clsK1Exception("The trial period for " & objActivationKey.ProductName & " has expired. To continue using " & objLicenceFileInfo.ProductName & _
                                                 " activate " & objActivationKey.ProductName & " using a valid Activation Key in the DRM.", True)
                    ElseIf intTotalDays < 0 Then
                        Throw New clsK1Exception("A valid Activation Key not found, cannot run.", True)
                    Else
                        If Not String.IsNullOrEmpty(strMessage) Then
                            strMessage &= vbCrLf & vbCrLf
                        Else
                            strMessage = ""
                        End If

                        strMessage &= "The database has not been activated. It will work for 45 days and then will cease to operate unless it is activated." & vbCrLf & _
                            "There are currently " & (45 - intTotalDays) & " days left to activate."
                    End If
                Catch ex As clsK1Exception
                    Throw
                Catch ex As Exception
                    Throw New clsK1Exception("Installation validation has failed, cannot run.", True)
                End Try

                Return True
            Catch ex As clsK1Exception
                Throw
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function GetAllLicences(ByVal objDB As clsDB) As FrameworkCollections.K1Collection(Of clsLicenceFileInfo)
            Try
                Dim colLicences As New FrameworkCollections.K1Collection(Of clsLicenceFileInfo)

                'Get All LicenceFile Records
                Dim arrApplicationTypes() As Integer = CType([Enum].GetValues(GetType(enumApplicationType)), Integer())

                For intIndex As Integer = 0 To arrApplicationTypes.GetUpperBound(0)
                    If enumApplicationType.K1ActivationKey = arrApplicationTypes(intIndex) OrElse
                        enumApplicationType.RecFindActivationKey = arrApplicationTypes(intIndex) Then
                        Continue For
                    End If

                    Try
                        Dim objLicenceInfo As New clsLicenceFileInfo(objDB, CType(arrApplicationTypes(intIndex), enumApplicationType))
                        If objLicenceInfo IsNot Nothing AndAlso objLicenceInfo.LicenceExists Then
                            colLicences.Add(objLicenceInfo)
                        End If
                    Catch ex As clsK1Exception
                    Catch ex As Exception
                        Throw
                    End Try
                Next

                Return colLicences
            Catch ex As Exception
                Throw
            End Try
        End Function

        Public Shared Sub RemoveLicence(ByVal objDB As clsDB_Direct, _
                                        ByVal eApptype As clsDBConstants.enumApplicationType)
            Try
                clsLicenseFile.Remove(objDB, eApptype)
            Catch ex As Exception
                Throw
            End Try
        End Sub

#End Region

    End Class

End Namespace