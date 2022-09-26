Imports System.Management
'Imports System.Xml.Serialization
'Imports System.ComponentModel

Public Class WMILogicalDisk

#Region " Members "

    Private m_strName As String
    Private m_strDescription As String
    Private m_strVolumeName As String
    Private m_strFileSystem As String
    Private m_strSize As String
    Private m_strFreeSpace As String
    Private m_strVolumeSerialNumber As String

#End Region

#Region " Constructor "

    Public Sub New()
    End Sub

    Public Sub New(ByVal strName As String, ByVal strDescription As String, ByVal strVolumeName As String, _
                   ByVal strFileSystem As String, ByVal strSize As String, ByVal strFreeSpace As String, _
                   ByVal strVolumeSerialNumber As String)
        m_strName = strName
        m_strDescription = strDescription
        m_strVolumeName = strVolumeName
        m_strFileSystem = strFileSystem
        m_strSize = strSize
        m_strFreeSpace = strFreeSpace
        m_strVolumeSerialNumber = strVolumeSerialNumber
    End Sub

    Public Sub New(ByVal instance As ManagementObject)
        m_strName = GetValue(instance.Properties("Name").Value)
        m_strDescription = GetValue(instance.Properties("Description").Value)
        m_strVolumeName = GetValue(instance.Properties("VolumeName").Value)
        m_strFileSystem = GetValue(instance.Properties("FileSystem").Value)
        m_strSize = GetValue(instance.Properties("Size").Value)
        m_strFreeSpace = GetValue(instance.Properties("FreeSpace").Value)
        m_strVolumeSerialNumber = GetValue(instance.Properties("VolumeSerialNumber").Value)
    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property Name() As String
        Get
            Return m_strName
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return m_strDescription
        End Get
    End Property

    Public ReadOnly Property VolumeName() As String
        Get
            Return m_strVolumeName
        End Get
    End Property

    Public ReadOnly Property FileSystem() As String
        Get
            Return m_strFileSystem
        End Get
    End Property

    Public ReadOnly Property Size() As String
        Get
            Return m_strSize
        End Get
    End Property
 _
    Public ReadOnly Property FreeSpace() As String
        Get
            Return m_strFreeSpace
        End Get
    End Property

    Public ReadOnly Property VolumeSerialNumber() As String
        Get
            Return m_strVolumeSerialNumber
        End Get
    End Property

#End Region

    Public Shared Function GetLogicalDisks() As Generic.Dictionary(Of String, WMILogicalDisk)
        Return GetLogicalDisks(".", Nothing, Nothing)
    End Function

    Public Shared Function GetLogicalDisks(ByVal strMachineName As String, ByVal strUserName As String, _
                                           ByVal strPassword As String) As Generic.Dictionary(Of String, WMILogicalDisk)
        Dim strPath As String = "\\" + strMachineName + "\root\cimv2"
        Dim mScope As ManagementScope = InitializeScope(strPath, strUserName, strPassword)
        Dim managementClass As ManagementClass = New ManagementClass(strPath + ":Win32_LogicalDisk")

        If mScope IsNot Nothing Then
            managementClass.Scope = mScope
        End If

        Dim instances As ManagementObjectCollection = managementClass.GetInstances()
        If instances Is Nothing Then
            Return Nothing
        End If

        Dim logicalDisks As New Generic.Dictionary(Of String, WMILogicalDisk)(instances.Count)
        Dim enumerator As ManagementObjectCollection.ManagementObjectEnumerator = instances.GetEnumerator()
        Do While enumerator.MoveNext()
            Dim objDisk As New WMILogicalDisk(CType(enumerator.Current, ManagementObject))
            logicalDisks.Add(objDisk.Name, objDisk)
        Loop

        Return logicalDisks
    End Function

    Public Shared Function GetLogicalDisksArray() As WMILogicalDisk()
        Return GetLogicalDisksArray(".", Nothing, Nothing)
    End Function

    Public Shared Function GetLogicalDisksArray(ByVal strMachineName As String, ByVal strUserName As String, _
                                           ByVal strPassword As String) As WMILogicalDisk()
        Dim strPath As String = "\\" + strMachineName + "\root\cimv2"
        Dim mScope As ManagementScope = InitializeScope(strPath, strUserName, strPassword)
        Dim managementClass As ManagementClass = New ManagementClass(strPath + ":Win32_LogicalDisk")

        If mScope IsNot Nothing Then
            managementClass.Scope = mScope
        End If

        Dim instances As ManagementObjectCollection = managementClass.GetInstances()
        If instances Is Nothing Then
            Return Nothing
        End If

        Dim logicalDisks(instances.Count) As WMILogicalDisk
        Dim index As Integer = 0
        Dim enumerator As ManagementObjectCollection.ManagementObjectEnumerator = instances.GetEnumerator()
        Do While enumerator.MoveNext()
            Dim objDisk As New WMILogicalDisk(CType(enumerator.Current, ManagementObject))
            logicalDisks(index) = objDisk
            index += 1
        Loop

        Return logicalDisks
    End Function

    Public Shared Function InitializeScope(ByVal strPath As String, ByVal strUserName As String, _
                                           ByVal strPassword As String) As ManagementScope

        If Not String.IsNullOrEmpty(strUserName) AndAlso Not String.IsNullOrEmpty(strPassword) Then
            Dim objOptions As ConnectionOptions = New ConnectionOptions()
            objOptions.Username = strUserName
            objOptions.Password = strPassword

            Dim objManScope As New ManagementScope()
            objManScope.Options = objOptions
            objManScope.Path = New ManagementPath(strPath)

            Return objManScope
        End If

        Return Nothing
    End Function

    Private Function GetValue(ByVal objValue As Object) As String
        Return GetValue(objValue, String.Empty)
    End Function

    Private Function GetValue(ByVal objValue As Object, ByVal strDefault As String) As String
        Try
            If objValue Is Nothing Then
                Return strDefault
            Else
                Return CStr(objValue).Trim()
            End If
        Catch ex As Exception
            Return strDefault
        End Try
    End Function

End Class

