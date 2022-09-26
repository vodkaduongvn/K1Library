Imports System.Configuration

Namespace Configuration

    ''' <summary>
    ''' Base Class for creating custom configuration sections in the application configuration file
    ''' </summary>
    ''' <remarks>
    ''' You must add the following tag to configuration/configSections:
    '''    <section name="[SectionName]" type="[StrongName], [Assembly]" />
    ''' 
    ''' Example
    '''   [SectionName] = BroadcastingSettings
    '''   [StrongName]  = K1.DRM.BroadcastingSettingsSection
    '''   [Assembly]    = DRM
    ''' </remarks>
    Public MustInherit Class SettingsSectionBase
        Inherits ConfigurationSection

        Protected Shared m_objConfiguration As System.Configuration.Configuration
        Protected Shared ReadOnly m_objPropertySettings As ConfigurationProperty = New ConfigurationProperty(Nothing,
                                                                                                             GetType(SettingElementCollection),
                                                                                                             Nothing,
                                                                                                             ConfigurationPropertyOptions.IsDefaultCollection)
        <ConfigurationProperty("", IsDefaultCollection:=True)> _
        Public ReadOnly Property Settings() As SettingElementCollection
            Get
                Return DirectCast(MyBase.Item(m_objPropertySettings), SettingElementCollection)
            End Get
        End Property

        Public Shared Function GetCurrentSection(Of T As ConfigurationSection)(ByVal strSectionName As String,
                                                                                   Optional configfilePath As String = "") As T
            If m_objConfiguration Is Nothing Then
                If (String.IsNullOrEmpty(configfilePath)) Then
                    m_objConfiguration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                Else
                    ' Map the new configuration file. 
                    Dim configFileMap As New ExeConfigurationFileMap()
                    configFileMap.ExeConfigFilename = configfilePath
                    m_objConfiguration = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None)
                End If
            End If
            Dim section = m_objConfiguration.GetSection(strSectionName)

            Return CType(section, T)

        End Function


        Public Shared Function MapAppConfigFile(ByVal appName As String, ByVal folderPath As String) As String

            ' Get the application configuration file. 
            Dim config As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            ' Create a new configuration file by saving  
            ' the application configuration to a new file. 

            Dim configFile As String = String.Concat(appName, ".exe.config")

            Dim filePath = Path.Combine(folderPath, configFile)

            If (Not File.Exists(filePath)) Then
                config.SaveAs(filePath, ConfigurationSaveMode.Full)
            End If

            Return filePath

        End Function

        ''' <summary>
        ''' Protects the section so it cannot manually be editted.
        ''' </summary>
        Public Sub Encrypt()
            Me.SectionInformation.ProtectSection("DataProtectionConfigurationProvider")
            Me.Save()
        End Sub

        ''' <summary>
        ''' Unprotects the section so it can manually be editted.
        ''' </summary>
        Public Sub Decrypt()
            Me.SectionInformation.UnprotectSection()
            Me.Save()
        End Sub

        ''' <summary>
        ''' Saves the section to the application configuration file
        ''' </summary>
        Public Sub Save()
            Me.SectionInformation.ForceSave = True
            m_objConfiguration.Save(ConfigurationSaveMode.Full)
        End Sub

        ''' <summary>
        ''' Returns the innerText for the specified SettingElement.
        ''' </summary>
        ''' <param name="strSettingName">Name of the SettingElement that we want to return the value for.</param>
        Protected Function GetSetting(ByVal strSettingName As String) As String
            Try
                Dim objSetting As SettingElement = Me.Settings.Get(strSettingName)

                If objSetting Is Nothing Then
                    Return Nothing
                End If

                Return objSetting.Value.ValueXml.InnerText
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Sets the innerText for the specified SettingElement.
        ''' </summary>
        ''' <param name="strSettingName">Name of the SettingElement that we want to set value for.</param>
        ''' <param name="strValue">Value to set.</param>
        Protected Sub SetSetting(ByVal strSettingName As String, ByVal strValue As String)
            Try
                Dim objSetting As SettingElement = Me.Settings.Get(strSettingName)

                If objSetting Is Nothing Then
                    objSetting = New SettingElement(strSettingName, SettingsSerializeAs.String)
                    Dim objDoc As New System.Xml.XmlDocument

                    objSetting.Value.ValueXml = objDoc.CreateElement("value")
                    Me.Settings.Add(objSetting)
                End If

                objSetting.Value.ValueXml.InnerText = strValue
            Catch ex As Exception
                Throw
            End Try
        End Sub

    End Class

End Namespace