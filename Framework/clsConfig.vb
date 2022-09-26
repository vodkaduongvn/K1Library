Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Drawing

Public Class clsConfig

#Region " Members "

    Private m_strTColor1 As Integer
    Private m_strTColor2 As Integer
    Private m_strTColor3 As Integer
    Private m_strFColor1 As Integer
    Private m_strFColor2 As Integer
    Private m_strFColor3 As Integer
    Private m_strRColor1 As Integer
    Private m_strRColor2 As Integer
    Private m_objFont1 As clsConfigFont
    Private m_objFont2 As clsConfigFont
    Private m_objFont3 As clsConfigFont
    Private m_eEmailType As clsEmailInfo.enumMailType
    Private m_strSMTPServer As String
    Private m_blnNewWindowFullScreen As Boolean
    Private m_eScanType As clsDBConstants.enumImageTypes
    Private m_blnUseLowGraphics As Boolean
    '2016-12-07 -- Peter Melisi -- Changes for adding an option for switching to Client Side OCR
    Private m_blnPerformOCRClientSide As Boolean
    '201904-11 -- Emmanuel -- defaults for bulk check in
    Private m_BulkCheckInDfltFolder As String
    Private m_BulkCheckInDfltSecurityCode As String
    Private m_BulkCheckInDfltDeleteFlag As Boolean
    Private m_BulkCheckInDfltEDOCType As String
    Private m_BulkCheckInDfltMetadataProfile As String
    Private m_BulkCheckInDfltMetadataProfileId As String
    Private m_BulkCheckInDfltCustomText As String
    Private m_BulkCheckInDfltFileNameFlag As Boolean
#End Region

#Region " Constructors "

    Public Sub New()
        m_strTColor1 = Color.WhiteSmoke.ToArgb
        m_strTColor2 = Color.LightGray.ToArgb
        m_strTColor3 = Color.DarkGray.ToArgb
        m_strFColor1 = Color.AliceBlue.ToArgb
        m_strFColor2 = Color.CornflowerBlue.ToArgb
        m_strFColor3 = Color.Gray.ToArgb
        m_strRColor1 = Color.White.ToArgb
        m_strRColor2 = Color.AliceBlue.ToArgb

        m_eEmailType = clsEmailInfo.enumMailType.MAPI
        m_objFont1 = New clsConfigFont("Tahoma", 8.25, FontStyle.Regular, Color.Black.ToArgb)
        m_objFont2 = New clsConfigFont("Tahoma", 8.25, FontStyle.Regular, Color.Black.ToArgb)
        m_objFont3 = New clsConfigFont("Tahoma", 8.25, FontStyle.Regular, Color.Black.ToArgb)
        m_blnNewWindowFullScreen = True
        m_eScanType = clsDBConstants.enumImageTypes.Tif
        m_blnUseLowGraphics = False
        '2016-12-07 -- Peter Melisi -- Changes for adding an option for switching to Client Side OCR
        m_blnPerformOCRClientSide = False
    End Sub
#End Region

#Region " Properties "

    Public Property TColor1() As Integer
        Get
            Return m_strTColor1
        End Get
        Set(ByVal value As Integer)
            m_strTColor1 = value
        End Set
    End Property

    Public Property TColor2() As Integer
        Get
            Return m_strTColor2
        End Get
        Set(ByVal value As Integer)
            m_strTColor2 = value
        End Set
    End Property

    Public Property TColor3() As Integer
        Get
            Return m_strTColor3
        End Get
        Set(ByVal value As Integer)
            m_strTColor3 = value
        End Set
    End Property

    Public Property FColor1() As Integer
        Get
            Return m_strFColor1
        End Get
        Set(ByVal value As Integer)
            m_strFColor1 = value
        End Set
    End Property

    Public Property FColor2() As Integer
        Get
            Return m_strFColor2
        End Get
        Set(ByVal value As Integer)
            m_strFColor2 = value
        End Set
    End Property

    Public Property FColor3() As Integer
        Get
            Return m_strFColor3
        End Get
        Set(ByVal value As Integer)
            m_strFColor3 = value
        End Set
    End Property

    Public Property RColor1() As Integer
        Get
            Return m_strRColor1
        End Get
        Set(ByVal value As Integer)
            m_strRColor1 = value
        End Set
    End Property

    Public Property RColor2() As Integer
        Get
            Return m_strRColor2
        End Get
        Set(ByVal value As Integer)
            m_strRColor2 = value
        End Set
    End Property

    Public Property Font1() As clsConfigFont
        Get
            Return m_objFont1
        End Get
        Set(ByVal value As clsConfigFont)
            m_objFont1 = value
        End Set
    End Property

    Public Property Font2() As clsConfigFont
        Get
            Return m_objFont2
        End Get
        Set(ByVal value As clsConfigFont)
            m_objFont2 = value
        End Set
    End Property

    Public Property Font3() As clsConfigFont
        Get
            Return m_objFont3
        End Get
        Set(ByVal value As clsConfigFont)
            m_objFont3 = value
        End Set
    End Property

    Public Property ScanType() As clsDBConstants.enumImageTypes
        Get
            Return m_eScanType
        End Get
        Set(ByVal value As clsDBConstants.enumImageTypes)
            m_eScanType = value
        End Set
    End Property

    Public ReadOnly Property ActualTColor1() As Color
        Get
            Return Color.FromArgb(m_strTColor1)
        End Get
    End Property

    Public ReadOnly Property ActualTColor2() As Color
        Get
            Return Color.FromArgb(m_strTColor2)
        End Get
    End Property

    Public ReadOnly Property ActualTColor3() As Color
        Get
            Return Color.FromArgb(m_strTColor3)
        End Get
    End Property

    Public ReadOnly Property ActualFColor1() As Color
        Get
            Return Color.FromArgb(m_strFColor1)
        End Get
    End Property

    Public ReadOnly Property ActualFColor2() As Color
        Get
            Return Color.FromArgb(m_strFColor2)
        End Get
    End Property

    Public ReadOnly Property ActualFColor3() As Color
        Get
            Return Color.FromArgb(m_strFColor3)
        End Get
    End Property

    Public ReadOnly Property ActualRColor1() As Color
        Get
            Return Color.FromArgb(m_strRColor1)
        End Get
    End Property

    Public ReadOnly Property ActualRColor2() As Color
        Get
            Return Color.FromArgb(m_strRColor2)
        End Get
    End Property

    Public ReadOnly Property ActualFont1() As Font
        Get
            Return New Font(m_objFont1.Name, m_objFont1.Size, CType(m_objFont1.FontStyle, FontStyle))
        End Get
    End Property

    Public ReadOnly Property ActualFont2() As Font
        Get
            Return New Font(m_objFont2.Name, m_objFont2.Size, CType(m_objFont2.FontStyle, FontStyle))
        End Get
    End Property

    'Public ReadOnly Property ActualFont3() As Font
    '    Get
    '        Return New Font(m_objFont3.Name, m_objFont3.Size, CType(m_objFont3.FontStyle, FontStyle))
    '    End Get
    'End Property

    Public ReadOnly Property ActualFontColor1() As Color
        Get
            Return Color.FromArgb(m_objFont1.Color)
        End Get
    End Property

    Public ReadOnly Property ActualFontColor2() As Color
        Get
            Return Color.FromArgb(m_objFont2.Color)
        End Get
    End Property

    'Public ReadOnly Property ActualFontColor3() As Color
    '    Get
    '        Return Color.FromArgb(m_objFont3.Color)
    '    End Get
    'End Property

    Public Property NewWindowWidthFullScreen() As Boolean
        Get
            Return m_blnNewWindowFullScreen
        End Get
        Set(ByVal value As Boolean)
            m_blnNewWindowFullScreen = value
        End Set
    End Property

    Public Property EmailType() As clsEmailInfo.enumMailType
        Get
            Return m_eEmailType
        End Get
        Set(ByVal value As clsEmailInfo.enumMailType)
            m_eEmailType = value
        End Set
    End Property

    Public Property SMTPServer() As String
        Get
            Return m_strSMTPServer
        End Get
        Set(ByVal value As String)
            m_strSMTPServer = value
        End Set
    End Property

    Public Property UseLowGraphics() As Boolean
        Get
            Return m_blnUseLowGraphics
        End Get
        Set(ByVal value As Boolean)
            m_blnUseLowGraphics = value
        End Set
    End Property

    '2016-12-07 -- Peter Melisi -- Changes for adding an option for switching to Client Side OCR
    Public Property PerformOCRClientSide() As Boolean
        Get
            Return m_blnPerformOCRClientSide
        End Get
        Set(ByVal value As Boolean)
            m_blnPerformOCRClientSide = value
        End Set
    End Property
    
    '201904-11 -- Emmanuel -- defaults for bulk check in
    Public Property BulkCheckInDfltFolder As String
        Get
            Return m_BulkCheckInDfltFolder
        End Get
        Set(value As String)
            m_BulkCheckInDfltFolder = value
        End Set
    End Property
    Public Property BulkCheckInDfltSecurityCode As String
        Get
            Return m_BulkCheckInDfltSecurityCode
        End Get
        Set(value As String)
            m_BulkCheckInDfltSecurityCode = value
        End Set
    End Property
    Public Property BulkCheckInDfltEDOCType As String
        Get
            Return m_BulkCheckInDfltEDOCType
        End Get
        Set(value As String)
            m_BulkCheckInDfltEDOCType = value
        End Set
    End Property
    Public Property BulkCheckInDfltMetadataProfile As String
        Get
            Return m_BulkCheckInDfltMetadataProfile
        End Get
        Set(value As String)
            m_BulkCheckInDfltMetadataProfile = value
        End Set
    End Property
    Public Property BulkCheckInDfltMetadataProfileId As String
        Get
            Return m_BulkCheckInDfltMetadataProfileId
        End Get
        Set(value As String)
            m_BulkCheckInDfltMetadataProfileId = value
        End Set
    End Property
    Public Property BulkCheckInDfltCustomText As String
        Get
            Return m_BulkCheckInDfltCustomText
        End Get
        Set(value As String)
            m_BulkCheckInDfltCustomText = value
        End Set
    End Property

    Public Property BulkCheckInDfltFileNameFlag As Boolean
        Get
            Return m_BulkCheckInDfltFileNameFlag 
        End Get
        Set(value As Boolean)
            m_BulkCheckInDfltFileNameFlag = value
        End Set
    End Property

    Public Property BulkCheckInDfltDeleteFlag As Boolean
        Get
            Return m_BulkCheckInDfltDeleteFlag 
        End Get
        Set(value As Boolean)
            m_BulkCheckInDfltDeleteFlag = value
        End Set
    End Property

#End Region

#Region " Methods "

    Public Shared Function GetXMLFile() As String
        Return ProperPath(System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) & "RecFind6\Client\Config.xml"
    End Function

    Public Shared Function Deserialize(ByVal objDB As clsDB) As clsConfig
        Dim objConfig As clsConfig
        Dim strXML As String = objDB.Profile.GetSettingsXML

        If strXML Is Nothing Then
            objConfig = New clsConfig
        Else
            Dim objSerializer As New XmlSerializer(GetType(clsConfig))
            Dim objTextReader As New StringReader(strXML)
            Dim objXmlReader As New XmlTextReader(objTextReader)

            If objSerializer.CanDeserialize(objXmlReader) Then
                objConfig = CType(objSerializer.Deserialize(objXmlReader), clsConfig)
            Else
                objConfig = New clsConfig
            End If

            objXmlReader.Close()
            objTextReader.Close()
            objTextReader.Dispose()
        End If

        If String.IsNullOrEmpty(objConfig.SMTPServer) Then
            '-- Default to SMTP Server set in registration
            Dim objK1Config As clsK1Configuration = clsK1Configuration.GetDefault(objDB)
            If Not String.IsNullOrEmpty(objK1Config.SMTPServer) Then
                objConfig.SMTPServer = objK1Config.SMTPServer
            End If
        End If

        Return objConfig
    End Function

    Public Function Serialize() As String
        Dim objSB As New System.Text.StringBuilder
        Dim objWriter As New StringWriter(objSB)
        Dim objXmlSerializer As New XmlSerializer(Me.GetType())

        objXmlSerializer.Serialize(objWriter, Me)

        objWriter.Close()

        Return objSB.ToString
    End Function

    Public Function Serialize(ByVal objDB As clsDB) As String
        Dim strXML As String = Me.Serialize

        Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cUSERPROFILE)

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(
                    objTable, objDB.Profile.ID)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.UserProfile.cSETTINGS, strXML)

        colMasks.Update(objDB)

        Return strXML
    End Function
#End Region

End Class
