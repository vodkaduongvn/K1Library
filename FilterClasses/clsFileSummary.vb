Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports System.Xml.Linq

Public Class clsFileSummary

#Region "Fields"

    Private m_strAuthor As String
    Private m_strTitle As String
    Private m_strSubject As String
    Private m_strKeyWords As String
    Private m_strComments As String
    Private m_strCategory As String
    Private dso As DSOFile.OleDocumentPropertiesClass
    Private m_StrErrorMessage As String
#End Region

#Region "Properties"

    Public Property Author() As String
        Get
            Return m_strAuthor
        End Get
        Set(ByVal Value As String)
            m_strAuthor = Value
        End Set
    End Property

    Public ReadOnly Property Title() As String
        Get
            Return m_strTitle
        End Get
    End Property

    Public Property Subject() As String
        Get
            Return m_strSubject
        End Get
        Set(ByVal Value As String)
            m_strSubject = Value
        End Set
    End Property

    Public ReadOnly Property KeyWords() As String
        Get
            Return m_strKeyWords
        End Get
    End Property

    Public ReadOnly Property Comments() As String
        Get
            Return m_strComments
        End Get
    End Property

    Public ReadOnly Property Category() As String
        Get
            Return m_strCategory
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return m_StrErrorMessage
        End Get
    End Property
#End Region

#Region "Constructors"

    Public Sub New(ByVal strFileName As String)

        Dim strExtension = IO.Path.GetExtension(strFileName)

        Select Case strExtension
            Case ".docx", ".docm", ".dotx", ".docm"
                ExtractWordProperties(strFileName)
            Case ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam"
                ExtractExcelProperties(strFileName)
            Case ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx"
                ExtractPowerPointProperties(strFileName)
            Case Else
                Initiliaze(strFileName)
        End Select

    End Sub

    ''' <summary>
    '''[Naing] Do not use this. It is for unit testing purposes for now.
    ''' </summary>
    ''' <param name="strFileName"></param>
    ''' <param name="blnUnsafe"></param>
    ''' <remarks>Do not use for production code.</remarks>
    Public Sub New(ByVal strFileName As String, blnUnsafe As Boolean)

        If (blnUnsafe) Then
            InitiliazeUnsafe(strFileName)
        Else
            Initiliaze(strFileName)
        End If
    End Sub

#End Region

#Region "Members"

    Private Sub InitiliazeUnsafe(ByVal strFileName As String)
        dso = New DSOFile.OleDocumentPropertiesClass
        dso.Open(strFileName.Trim, True, DSOFile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

        m_strAuthor = CStr(IIf(dso.SummaryProperties.Author Is Nothing, "", dso.SummaryProperties.Author))
        m_strSubject = CStr(IIf(dso.SummaryProperties.Subject Is Nothing, "", dso.SummaryProperties.Subject))
        m_strTitle = CStr(IIf(dso.SummaryProperties.Title Is Nothing, "", dso.SummaryProperties.Title))
        m_strKeyWords = CStr(IIf(dso.SummaryProperties.Keywords Is Nothing, "", dso.SummaryProperties.Keywords))
        m_strComments = CStr(IIf(dso.SummaryProperties.Comments Is Nothing, "", dso.SummaryProperties.Comments))
        m_strCategory = CStr(IIf(dso.SummaryProperties.Category Is Nothing, "", dso.SummaryProperties.Category))

        dso.Close()
    End Sub

    ''-- 14/12/2015 -- Neelam -- This is used for DSOfile it will be still in use for document formats before Office 2007
    '' for dso to work we need to install at least office 2000 & then with the office compatibility pack we don't need to install 
    '' office 2007 onwards versions
    Private Sub Initiliaze(ByVal strFileName As String)
        Try
            dso = New DSOFile.OleDocumentPropertiesClass
            dso.Open(strFileName.Trim, True, DSOFile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

            m_strAuthor = CStr(IIf(dso.SummaryProperties.Author Is Nothing, "", dso.SummaryProperties.Author))
            m_strSubject = CStr(IIf(dso.SummaryProperties.Subject Is Nothing, "", dso.SummaryProperties.Subject))
            m_strTitle = CStr(IIf(dso.SummaryProperties.Title Is Nothing, "", dso.SummaryProperties.Title))
            m_strKeyWords = CStr(IIf(dso.SummaryProperties.Keywords Is Nothing, "", dso.SummaryProperties.Keywords))
            m_strComments = CStr(IIf(dso.SummaryProperties.Comments Is Nothing, "", dso.SummaryProperties.Comments))
            m_strCategory = CStr(IIf(dso.SummaryProperties.Category Is Nothing, "", dso.SummaryProperties.Category))

            dso.Close()
        Catch ex As Exception
        End Try
    End Sub

    ''-- 14/12/2015 -- Neelam -- Here we need to get the coreProperties node & it's child elements
    '' this is not the 100% perfect code but something which is working to extract the document properties
    '' feel free to change something more efficient
    '' In OpenXml document properties extraction I wasn't able to extract the property "Category"
    Private Sub ExtractWordProperties(ByVal strFileName As String)

        Dim coreFileProperties As CoreFilePropertiesPart
        Dim stream As Stream
        Dim xdoc As XDocument
        Dim document As WordprocessingDocument

        Try
            document = WordprocessingDocument.Open(strFileName, True)

            Using document

                coreFileProperties = document.CoreFilePropertiesPart

                If Not IsNothing(coreFileProperties) Then
                    stream = coreFileProperties.GetStream()

                    xdoc = XDocument.Load(coreFileProperties.GetStream())

                    Dim mainElements As IEnumerable(Of XElement) = From el In xdoc.Descendants()
                                                                 Where el.Name.[Namespace] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                                                                 Select el

                    For Each el As XElement In mainElements

                        Select Case el.Name.LocalName
                            Case "coreProperties"

                                Dim childElements As IEnumerable(Of XElement) = From cl In el.Descendants()
                                                                              Select cl

                                For Each cl As XElement In childElements

                                    Select Case cl.Name.LocalName

                                        Case "title"
                                            m_strTitle = cl.Value

                                        Case "subject"
                                            m_strSubject = cl.Value

                                        Case "creator"
                                            m_strAuthor = cl.Value

                                        Case "keywords"
                                            m_strKeyWords = cl.Value

                                        Case "description"
                                            m_strComments = cl.Value
                                    End Select
                                Next

                        End Select
                    Next

                End If

            End Using

        Catch ex As Exception
            m_strTitle = Nothing
            m_strSubject = Nothing
            m_strAuthor = Nothing
            m_strKeyWords = Nothing
            m_strComments = Nothing
            m_StrErrorMessage = ex.Message
        Finally
            coreFileProperties = Nothing
            stream = Nothing
            xdoc = Nothing
            document = Nothing
        End Try
    End Sub

    ''-- 14/12/2015 -- Neelam -- Here we need to get the coreProperties node & it's child elements
    '' this is not the 100% perfect code but something which is working to extract the document properties
    '' feel free to change something more efficient
    '' In OpenXml document properties extraction I wasn't able to extract the property "Category"
    Private Sub ExtractExcelProperties(ByVal strFileName As String)

        Dim coreFileProperties As CoreFilePropertiesPart
        Dim stream As Stream
        Dim xdoc As XDocument
        Dim document As SpreadsheetDocument
        Dim strFileNameFix As String = ""

        Try

            document = SpreadsheetDocument.Open(strFileName, False)

        Catch ex As OpenXmlPackageException When ex.InnerException.GetType() Is GetType(UriFormatException)
            '' Emmanuel Cardakaris - 2200003756
            '' This is to handle an exception where Excel incorrectly stores a string as a URI, and OpenXML errors because the URI is malformed
            '' The fix is to find the malformed URI and rewrite some dummy text that is correctly formed
            '' Discovered in a spreadsheet that had an '@' in the cell with other text
            strFileNameFix = strFileName & "_fix"

            FileSystem.FileCopy(strFileName, strFileNameFix)
            Using fs As FileStream = New FileStream(strFileNameFix, FileMode.Open, FileAccess.ReadWrite)
                FixInvalidUri(fs, Function(brokenUri As String) New Uri("http://dummy-link/"))
            End Using
            document = SpreadsheetDocument.Open(strFileNameFix, False)
        End Try

        Try
            Using document

                coreFileProperties = document.CoreFilePropertiesPart

                If Not IsNothing(coreFileProperties) Then
                    stream = coreFileProperties.GetStream()

                    xdoc = XDocument.Load(coreFileProperties.GetStream())

                    Dim mainElements As IEnumerable(Of XElement) = From el In xdoc.Descendants()
                                                                   Where el.Name.[Namespace] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                                                                   Select el

                    For Each el As XElement In mainElements

                        Select Case el.Name.LocalName
                            Case "coreProperties"

                                Dim childElements As IEnumerable(Of XElement) = From cl In el.Descendants()
                                                                                Select cl

                                For Each cl As XElement In childElements

                                    Select Case cl.Name.LocalName

                                        Case "title"
                                            m_strTitle = cl.Value

                                        Case "subject"
                                            m_strSubject = cl.Value

                                        Case "creator"
                                            m_strAuthor = cl.Value

                                        Case "keywords"
                                            m_strKeyWords = cl.Value

                                        Case "description"
                                            m_strComments = cl.Value
                                    End Select
                                Next

                        End Select
                    Next

                End If

            End Using

        Catch ex As Exception
            m_strTitle = Nothing
            m_strSubject = Nothing
            m_strAuthor = Nothing
            m_strKeyWords = Nothing
            m_strComments = Nothing
            m_StrErrorMessage = ex.Message
        Finally
            coreFileProperties = Nothing
            stream = Nothing
            xdoc = Nothing
            document = Nothing

            If Not String.IsNullOrEmpty(strFileNameFix) Then
                FileSystem.Kill(strFileNameFix)
            End If
        End Try
    End Sub

    ''-- 14/12/2015 -- Neelam -- Here we need to get the coreProperties node & it's child elements
    '' this is not the 100% perfect code but something which is working to extract the document properties
    '' feel free to change something more efficient
    '' In OpenXml document properties extraction I wasn't able to extract the property "Category"
    Private Sub ExtractPowerPointProperties(ByVal strFileName As String)

        Dim coreFileProperties As CoreFilePropertiesPart
        Dim stream As Stream
        Dim xdoc As XDocument
        Dim document As PresentationDocument

        document = PresentationDocument.Open(strFileName, False)

        Try

            Using document

                coreFileProperties = document.CoreFilePropertiesPart

                If Not IsNothing(coreFileProperties) Then
                    stream = coreFileProperties.GetStream()

                    xdoc = XDocument.Load(coreFileProperties.GetStream())

                    Dim mainElements As IEnumerable(Of XElement) = From el In xdoc.Descendants()
                                                                 Where el.Name.[Namespace] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                                                                 Select el

                    For Each el As XElement In mainElements

                        Select Case el.Name.LocalName
                            Case "coreProperties"

                                Dim childElements As IEnumerable(Of XElement) = From cl In el.Descendants()
                                                                              Select cl

                                For Each cl As XElement In childElements

                                    Select Case cl.Name.LocalName

                                        Case "title"
                                            m_strTitle = cl.Value

                                        Case "subject"
                                            m_strSubject = cl.Value

                                        Case "creator"
                                            m_strAuthor = cl.Value

                                        Case "keywords"
                                            m_strKeyWords = cl.Value

                                        Case "description"
                                            m_strComments = cl.Value
                                    End Select
                                Next

                        End Select
                    Next

                End If

            End Using

        Catch ex As Exception
            m_strTitle = Nothing
            m_strSubject = Nothing
            m_strAuthor = Nothing
            m_strKeyWords = Nothing
            m_strComments = Nothing
            m_StrErrorMessage = ex.Message
        Finally
            coreFileProperties = Nothing
            stream = Nothing
            xdoc = Nothing
            document = Nothing
        End Try
    End Sub

#End Region

#Region "Commented Code"

    '*******************************************************************************************
    ''-- 14/12/2015 -- Neelam -- These are the core propeties I found for extraction
    '*******************************************************************************************
    '<dc:title>title</dc:title>
    '<dc:subject>teesstt</dc:subject>
    '<dc:creator>Neelam Patil;testUser</dc:creator>
    '<cp:keywords>tag</cp:keywords>
    '<dc:description>comment</dc:description>
    '<cp:lastModifiedBy>Neelam Patil</cp:lastModifiedBy>
    '<cp:revision>4</cp:revision>
    '<dcterms:created xsi:type="dcterms:W3CDTF">2015-10-15T00:28:00Z</dcterms:created>
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2015-12-13T22:14:00Z</dcterms:modified>
    '********************************************************************************************

    'Public Sub New(ByVal strFileName As String)

    '    Dim strExtension = IO.Path.GetExtension(strFileName)

    '    Select Case strExtension
    '        Case ".docx", ".docm", ".dotx", ".docm",
    '             ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam",
    '             ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx"
    '            ExtractFileProperties(strFileName)
    '        Case Else
    '            Initiliaze(strFileName)
    '    End Select

    'End Sub

    ' ''-- 14/12/2015 -- Neelam -- Here we need to get the coreProperties node & it's child elements
    ' '' this is not the 100% perfect code but something which is working to get the document properties
    ' '' feel free to change something more efficient
    'Private Sub ExtractFileProperties(ByVal strFileName As String)

    '    Dim coreFileProperties As CoreFilePropertiesPart
    '    Dim stream As Stream
    '    Dim xdoc As XDocument
    '    Dim strDocType As String


    '    Select Case (IO.Path.GetExtension(strFileName))
    '        Case ".docx", ".docm", ".dotx", ".docm"
    '            strDocType = "Word"
    '        Case ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam"
    '            strDocType = "Excel"
    '        Case ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx"
    '            strDocType = "PowerPoint"
    '    End Select

    '    If strDocType = "Word" Then
    '        Dim document As WordprocessingDocument = WordprocessingDocument.Open(strFileName, False)
    '    ElseIf strDocType = "Excel" Then
    '        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(strFileName, False)
    '    ElseIf strDocType = "PowerPoint" Then
    '        Dim document As PresentationDocument = PresentationDocument.Open(strFileName, False)
    '    End If

    '    'Dim document As WordprocessingDocument = WordprocessingDocument.Open(strFileName, False)

    '    Using document

    '        coreFileProperties = document.CoreFilePropertiesPart

    '        If Not IsNothing(coreFileProperties) Then
    '            stream = coreFileProperties.GetStream()

    '            xdoc = XDocument.Load(coreFileProperties.GetStream())

    '            Dim mainElements As IEnumerable(Of XElement) = From el In xdoc.Descendants()
    '                                                         Where el.Name.[Namespace] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    '                                                         Select el

    '            For Each el As XElement In mainElements

    '                Select Case el.Name.LocalName
    '                    Case "coreProperties"

    '                        Dim childElements As IEnumerable(Of XElement) = From cl In el.Descendants()
    '                                                                      Select cl

    '                        For Each cl As XElement In childElements

    '                            Select Case cl.Name.LocalName

    '                                Case "title"
    '                                    m_strTitle = cl.Value

    '                                Case "subject"
    '                                    m_strSubject = cl.Value

    '                                Case "creator"
    '                                    m_strAuthor = cl.Value

    '                                Case "keywords"
    '                                    m_strKeyWords = cl.Value

    '                                Case "description"
    '                                    m_strComments = cl.Value
    '                            End Select
    '                        Next

    '                End Select
    '            Next

    '        End If

    '    End Using
    'End Sub

#End Region

End Class





