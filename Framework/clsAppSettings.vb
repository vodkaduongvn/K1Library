Imports System
Imports System.Xml
Imports System.Configuration
Imports System.Reflection
Imports System.Diagnostics

Public Class clsAppSettings
    Private m_strDocName As String
    Private m_objNode As XmlNode
    Private m_eConfigType As enumConfigFileType

    Public Enum enumConfigFileType
        WEB_CONFIG = 0
        APP_CONFIG = 1
    End Enum

    Public Enum enumSettingType
        APPLICATION = 0
        USER = 1
    End Enum

    Public Sub New(ByVal eConfigType As enumConfigFileType)
        m_eConfigType = eConfigType
    End Sub

    Private Function LoadConfigDoc(ByVal objCfgDoc As XmlDocument) As XmlDocument
        'load the config file 
        If m_eConfigType = enumConfigFileType.APP_CONFIG Then
            m_strDocName = Assembly.GetEntryAssembly().GetName().Name & ".exe.config"
        Else
            m_strDocName = System.Web.HttpContext.Current.Server.MapPath("web.config")
        End If

        objCfgDoc.Load(m_strDocName)

        Return objCfgDoc
    End Function

    Public Function SetValue(ByVal strKey As String, ByVal strValue As String, ByVal eType As enumSettingType) As Boolean
        Try
            Dim objCfgDoc As New XmlDocument
            objCfgDoc = LoadConfigDoc(objCfgDoc)

            Dim strRoot As String
            If eType = enumSettingType.USER Then
                strRoot = "userSettings"
            Else
                strRoot = "applicationSettings"
            End If

            Dim strName As String = IO.Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().GetName().Name)

            m_objNode = objCfgDoc.SelectSingleNode("configuration/" & strRoot & "/" & strName & ".My.MySettings")
            If (m_objNode Is Nothing) Then
                Dim objRoot As XmlNode = objCfgDoc.DocumentElement
                Dim objNode1 As XmlElement = objCfgDoc.CreateElement(strRoot)
                objRoot.AppendChild(objNode1)
                m_objNode = objCfgDoc.CreateElement(strName & ".My.MySettings")
                objNode1.AppendChild(m_objNode)
            End If

            Dim objElem As XmlElement = objCfgDoc.CreateElement("setting")
            objElem.SetAttribute("name", strKey)
            objElem.SetAttribute("serializeAs", "String")
            m_objNode.AppendChild(objElem)

            Dim objVal As XmlElement = objCfgDoc.CreateElement("value")
            objVal.InnerText = strValue
            objElem.AppendChild(objVal)


            SaveConfigDoc(objCfgDoc, m_strDocName)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub SaveConfigDoc(ByVal objCfgDoc As XmlDocument, ByVal strCfgDocPath As String)
        Dim objWriter As New XmlTextWriter(strCfgDocPath, Nothing)
        objWriter.Formatting = Formatting.Indented
        objCfgDoc.WriteTo(objWriter)
        objWriter.Flush()
        objWriter.Close()
    End Sub

End Class