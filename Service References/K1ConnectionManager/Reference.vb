﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.Serialization

Namespace K1ConnectionManager
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0"),  _
     System.Runtime.Serialization.DataContractAttribute(Name:="IConnectionManager.SystemType", [Namespace]:="http://schemas.datacontract.org/2004/07/K1ConnectionManager")>  _
    Public Enum IConnectionManagerSystemType As Integer
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        K1 = 1
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        RecFind = 2
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        All = 3
    End Enum
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0"),  _
     System.Runtime.Serialization.DataContractAttribute(Name:="IConnectionManager.AccessType", [Namespace]:="http://schemas.datacontract.org/2004/07/K1ConnectionManager")>  _
    Public Enum IConnectionManagerAccessType As Integer
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        Direct = 1
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        WebService = 2
        
        <System.Runtime.Serialization.EnumMemberAttribute()>  _
        All = 3
    End Enum
    
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0"),  _
     System.Runtime.Serialization.DataContractAttribute(Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault"),  _
     System.SerializableAttribute()>  _
    Partial Public Class K1ServiceFault
        Inherits Object
        Implements System.Runtime.Serialization.IExtensibleDataObject, System.ComponentModel.INotifyPropertyChanged
        
        <System.NonSerializedAttribute()>  _
        Private extensionDataField As System.Runtime.Serialization.ExtensionDataObject
        
        Private ErrorNumberField As Integer
        
        Private MessageField As String
        
        Private TypeField As String
        
        <Global.System.ComponentModel.BrowsableAttribute(false)>  _
        Public Property ExtensionData() As System.Runtime.Serialization.ExtensionDataObject Implements System.Runtime.Serialization.IExtensibleDataObject.ExtensionData
            Get
                Return Me.extensionDataField
            End Get
            Set
                Me.extensionDataField = value
            End Set
        End Property
        
        <System.Runtime.Serialization.DataMemberAttribute(IsRequired:=true)>  _
        Public Property ErrorNumber() As Integer
            Get
                Return Me.ErrorNumberField
            End Get
            Set
                If (Me.ErrorNumberField.Equals(value) <> true) Then
                    Me.ErrorNumberField = value
                    Me.RaisePropertyChanged("ErrorNumber")
                End If
            End Set
        End Property
        
        <System.Runtime.Serialization.DataMemberAttribute(IsRequired:=true)>  _
        Public Property Message() As String
            Get
                Return Me.MessageField
            End Get
            Set
                If (Object.ReferenceEquals(Me.MessageField, value) <> true) Then
                    Me.MessageField = value
                    Me.RaisePropertyChanged("Message")
                End If
            End Set
        End Property
        
        <System.Runtime.Serialization.DataMemberAttribute(IsRequired:=true)>  _
        Public Property Type() As String
            Get
                Return Me.TypeField
            End Get
            Set
                If (Object.ReferenceEquals(Me.TypeField, value) <> true) Then
                    Me.TypeField = value
                    Me.RaisePropertyChanged("Type")
                End If
            End Set
        End Property
        
        Public Event PropertyChanged As System.ComponentModel.PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
        
        Protected Sub RaisePropertyChanged(ByVal propertyName As String)
            Dim propertyChanged As System.ComponentModel.PropertyChangedEventHandler = Me.PropertyChangedEvent
            If (Not (propertyChanged) Is Nothing) Then
                propertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(propertyName))
            End If
        End Sub
    End Class
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0"),  _
     System.ServiceModel.ServiceContractAttribute([Namespace]:="urn:KnowledgeoneCorp/Contracts/IConnectionManger", ConfigurationName:="K1ConnectionManager.IConnectionManager", SessionMode:=System.ServiceModel.SessionMode.Required)>  _
    Public Interface IConnectionManager
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemName"& _ 
            "s", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemName"& _ 
            "sResponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemName"& _ 
            "sK1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Function GetSystemNames(ByVal eSystemTypeFilter As K1ConnectionManager.IConnectionManagerSystemType, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType) As String()
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/SystemExists", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/SystemExistsR"& _ 
            "esponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/SystemExistsK"& _ 
            "1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Function SystemExists(ByVal strSystemName As String) As Boolean
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemInfo"& _ 
            "", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemInfo"& _ 
            "Response"),  _
         System.ServiceModel.FaultContractAttribute(GetType(System.ServiceModel.FaultException), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemInfo"& _ 
            "FaultExceptionFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="FaultException", [Namespace]:="http://schemas.datacontract.org/2004/07/System.ServiceModel"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/GetSystemInfo"& _ 
            "K1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Function GetSystemInfo(ByVal strSystemName As String, ByRef blnIsLatest As Boolean, ByVal dblVersion As Double, ByVal strAppName As String, ByVal blnSkipCheckVersion As Boolean, ByVal dblMinDatabaseVersion As Double) As String()
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ionString", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ionStringResponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ionStringK1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Sub InsertConnectionString(ByVal strSystemName As String, ByVal strConnectionString As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType)
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ion", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ionResponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/InsertConnect"& _ 
            "ionK1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Sub InsertConnection(ByVal strSystemName As String, ByVal strConnectionValue As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType)
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ionString", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ionStringResponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ionStringK1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Sub UpdateConnectionString(ByVal strSystemName As String, ByVal strConnectionString As String)
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ion", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ionResponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/UpdateConnect"& _ 
            "ionK1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Sub UpdateConnection(ByVal strSystemName As String, ByVal strConnectionValue As String, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType)
        
        <System.ServiceModel.OperationContractAttribute(ProtectionLevel:=System.Net.Security.ProtectionLevel.None, Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/DeleteSystem", ReplyAction:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/DeleteSystemR"& _ 
            "esponse"),  _
         System.ServiceModel.FaultContractAttribute(GetType(K1ConnectionManager.K1ServiceFault), Action:="urn:KnowledgeoneCorp/Contracts/IConnectionManger/IConnectionManager/DeleteSystemK"& _ 
            "1ServiceFaultFault", ProtectionLevel:=System.Net.Security.ProtectionLevel.EncryptAndSign, Name:="K1ServiceFault", [Namespace]:="urn:KnowledgeoneCorp/Fault/K1ServiceFault")>  _
        Sub DeleteSystem(ByVal strSystemName As String)
    End Interface
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
    Public Interface IConnectionManagerChannel
        Inherits K1ConnectionManager.IConnectionManager, System.ServiceModel.IClientChannel
    End Interface
    
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
    Partial Public Class ConnectionManagerClient
        Inherits System.ServiceModel.ClientBase(Of K1ConnectionManager.IConnectionManager)
        Implements K1ConnectionManager.IConnectionManager
        
        Public Sub New()
            MyBase.New
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String)
            MyBase.New(endpointConfigurationName)
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As String)
            MyBase.New(endpointConfigurationName, remoteAddress)
        End Sub
        
        Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
            MyBase.New(endpointConfigurationName, remoteAddress)
        End Sub
        
        Public Sub New(ByVal binding As System.ServiceModel.Channels.Binding, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
            MyBase.New(binding, remoteAddress)
        End Sub
        
        Public Function GetSystemNames(ByVal eSystemTypeFilter As K1ConnectionManager.IConnectionManagerSystemType, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType) As String() Implements K1ConnectionManager.IConnectionManager.GetSystemNames
            Return MyBase.Channel.GetSystemNames(eSystemTypeFilter, eAccessType)
        End Function
        
        Public Function SystemExists(ByVal strSystemName As String) As Boolean Implements K1ConnectionManager.IConnectionManager.SystemExists
            Return MyBase.Channel.SystemExists(strSystemName)
        End Function
        
        Public Function GetSystemInfo(ByVal strSystemName As String, ByRef blnIsLatest As Boolean, ByVal dblVersion As Double, ByVal strAppName As String, ByVal blnSkipCheckVersion As Boolean, ByVal dblMinDatabaseVersion As Double) As String() Implements K1ConnectionManager.IConnectionManager.GetSystemInfo
            Return MyBase.Channel.GetSystemInfo(strSystemName, blnIsLatest, dblVersion, strAppName, blnSkipCheckVersion, dblMinDatabaseVersion)
        End Function
        
        Public Sub InsertConnectionString(ByVal strSystemName As String, ByVal strConnectionString As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType) Implements K1ConnectionManager.IConnectionManager.InsertConnectionString
            MyBase.Channel.InsertConnectionString(strSystemName, strConnectionString, eSystemType)
        End Sub
        
        Public Sub InsertConnection(ByVal strSystemName As String, ByVal strConnectionValue As String, ByVal eSystemType As K1ConnectionManager.IConnectionManagerSystemType, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType) Implements K1ConnectionManager.IConnectionManager.InsertConnection
            MyBase.Channel.InsertConnection(strSystemName, strConnectionValue, eSystemType, eAccessType)
        End Sub
        
        Public Sub UpdateConnectionString(ByVal strSystemName As String, ByVal strConnectionString As String) Implements K1ConnectionManager.IConnectionManager.UpdateConnectionString
            MyBase.Channel.UpdateConnectionString(strSystemName, strConnectionString)
        End Sub
        
        Public Sub UpdateConnection(ByVal strSystemName As String, ByVal strConnectionValue As String, ByVal eAccessType As K1ConnectionManager.IConnectionManagerAccessType) Implements K1ConnectionManager.IConnectionManager.UpdateConnection
            MyBase.Channel.UpdateConnection(strSystemName, strConnectionValue, eAccessType)
        End Sub
        
        Public Sub DeleteSystem(ByVal strSystemName As String) Implements K1ConnectionManager.IConnectionManager.DeleteSystem
            MyBase.Channel.DeleteSystem(strSystemName)
        End Sub
    End Class
End Namespace
