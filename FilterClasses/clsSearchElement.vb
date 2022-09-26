#Region " File Information "

'==============================================================================
' This class represents a search criteria in a complete search filter
' 
' Ex. ((A = 1) or (B = 2)) is the complete search filter,
'   This class represents once section of the filter (ie. A = 1)
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       23/02/2007  Implemented.
'==============================================================================

#End Region

#End Region

Imports System.Xml
Imports System.Xml.Serialization

<Serializable()> Public Class clsSearchElement
    Inherits clsSearchObjBase
    Implements IXmlSerializable

#Region " Members "

    Private m_strFieldRef As String
    Private m_eCompareType As clsSearchFilter.enumComparisonType
    Private m_objValue As Object
    <NonSerialized()> Private m_objField As clsField
    Private m_blnIsVariable As Boolean
    Private m_blnIsConstant As Boolean
    <NonSerialized()> Private m_blnConstantSaved As Boolean
    Private m_strConstantValue As String

#End Region

#Region " Constructors "

    Public Sub New()

        MyBase.New()

    End Sub

    ''' <summary>
    ''' Creates a new filter element
    ''' </summary>
    ''' <param name="eOpType">AND, OR, ANDNOT, ORNOT, [Nothing]</param>
    ''' <param name="strFieldRef">Full Boolean representation of a field (ie. Person.EntityID.ExternalID)</param>
    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType,
                   ByVal strFieldRef As String)

        MyBase.New(eOpType)
        m_strFieldRef = strFieldRef.Trim

    End Sub

    ''' <summary>
    ''' Creates a new filter element
    ''' </summary>
    ''' <param name="eOpType">AND, OR, ANDNOT, ORNOT, [Nothing]</param>
    ''' <param name="strFieldRef">Full Boolean representation of a field (ie. Person.EntityID.ExternalID)</param>
    ''' <param name="eCompareType">The equality operation to use</param>
    ''' <param name="objValue">The value to filter by</param>
    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType,
                   ByVal strFieldRef As String,
                   ByVal eCompareType As clsSearchFilter.enumComparisonType,
                   ByVal objValue As Object)

        MyBase.New(eOpType)
        m_strFieldRef = strFieldRef.Trim
        m_eCompareType = eCompareType
        m_objValue = objValue

    End Sub

    ''' <summary>
    ''' Creates a new filter element
    ''' </summary>
    ''' <param name="eOpType">AND, OR, ANDNOT, ORNOT, [Nothing]</param>
    ''' <param name="strFieldRef">Full Boolean representation of a field (ie. Person.EntityID.ExternalID)</param>
    ''' <param name="eCompareType">The equality operation to use</param>
    ''' <param name="objValue">The value to filter by</param>
    ''' <param name="objField">The actual field object specified in the field ref</param>
    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType,
                   ByVal strFieldRef As String,
                   ByVal eCompareType As clsSearchFilter.enumComparisonType,
                   ByVal objValue As Object,
                   ByVal objField As clsField)

        MyBase.New(eOpType)
        m_strFieldRef = strFieldRef.Trim
        m_eCompareType = eCompareType
        m_objValue = objValue
        m_objField = objField

    End Sub

#End Region

#Region " Properties "

    ''' <summary>
    ''' Full Boolean representation of a field (ie. Person.EntityID.ExternalID)
    ''' </summary>
    Public Property FieldRef() As String
        Get
            Return m_strFieldRef
        End Get
        Set(ByVal value As String)
            m_strFieldRef = value
        End Set
    End Property

    ''' <summary>
    ''' The equality operation to use
    ''' </summary>
    Public Property CompareType() As clsSearchFilter.enumComparisonType
        Get
            Return m_eCompareType
        End Get
        Set(ByVal value As clsSearchFilter.enumComparisonType)
            m_eCompareType = value
        End Set
    End Property

    ''' <summary>
    ''' The value to filter by
    ''' </summary>
    Public Property Value() As Object
        Get
            Return m_objValue
        End Get
        Set(ByVal value As Object)
            m_objValue = value
        End Set
    End Property

    ''' <summary>
    ''' The Constant value if using system variables
    ''' </summary>
    Public Property ConstantValue() As String
        Get
            Return m_strConstantValue
        End Get
        Set(ByVal value As String)
            m_strConstantValue = value
        End Set
    End Property

    ''' <summary>
    ''' The actual field object specified in the field ref
    ''' </summary>
    Public ReadOnly Property Field() As clsField
        Get
            Return m_objField
        End Get
    End Property

    'Property for user variables
    Public Property IsVariable() As Boolean
        Get
            Return m_blnIsVariable
        End Get
        Set(ByVal value As Boolean)
            m_blnIsVariable = value
        End Set
    End Property

    'Property for system variables
    Public Property IsConstant() As Boolean
        Get
            Return m_blnIsConstant
        End Get
        Set(ByVal value As Boolean)
            m_blnIsConstant = value
        End Set
    End Property

    'Property to check if system variable is saved
    Public Property IsConstantSaved() As Boolean
        Get
            Return m_blnConstantSaved
        End Get
        Set(ByVal value As Boolean)
            m_blnConstantSaved = value
        End Set
    End Property

#End Region

#Region " Methods "

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
        Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
        reader.ReadStartElement()

        reader.ReadStartElement("FieldRef")
        Me.FieldRef = reader.ReadContentAsString()
        reader.ReadEndElement()

        reader.ReadStartElement("CompareType")
        Me.CompareType = CType(reader.ReadContentAsInt(), clsSearchFilter.enumComparisonType)
        reader.ReadEndElement()

        reader.ReadStartElement("IsVariable")
        Me.IsVariable = reader.ReadContentAsBoolean()
        reader.ReadEndElement()

        Try
            reader.ReadStartElement("IsConstant")
            Me.IsConstant = reader.ReadContentAsBoolean()
            reader.ReadEndElement()
        Catch ex As Exception
            'This is a new element and may not exist in previous saved searches
        End Try

        Try
            reader.ReadStartElement("ConstantValue")
            Me.ConstantValue = reader.ReadContentAsString()
            reader.ReadEndElement()
        Catch ex As Exception
            'This is also a new element and may not exist in previous saved searches
        End Try

        reader.ReadStartElement("OperatorType")
        Me.OperatorType = CType(reader.ReadContentAsInt(), clsSearchFilter.enumOperatorType)
        reader.ReadEndElement()

        reader.ReadStartElement("ValueType")
        If reader.NodeType = XmlNodeType.Whitespace Then
            reader.ReadStartElement("Value")
        Else
            Dim strType As String = reader.ReadContentAsString()
            reader.ReadEndElement()

            reader.ReadStartElement("Value")
            If strType Is Nothing Then
                Me.Value = reader.ReadContentAsString()
                Me.Value = Nothing
            Else
                Dim objType As Type = Type.GetType(strType)

                If objType.Equals(GetType(Hashtable)) Then
                    Dim strIDs As String = reader.ReadContentAsString()
                    Me.Value = CreateHashtableFromIDString(strIDs)
                Else
                    Me.Value = reader.ReadContentAs(objType, Nothing)
                End If
            End If
            reader.ReadEndElement()
        End If
        reader.ReadEndElement()
    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
        writer.WriteStartElement("FieldRef")
        writer.WriteValue(Me.FieldRef)
        writer.WriteEndElement()

        writer.WriteStartElement("CompareType")
        writer.WriteValue(Me.CompareType)
        writer.WriteEndElement()

        writer.WriteStartElement("IsVariable")
        writer.WriteValue(Me.IsVariable)
        writer.WriteEndElement()

        writer.WriteStartElement("IsConstant")
        writer.WriteValue(Me.IsConstant)
        writer.WriteEndElement()

        Try
            If Me.IsConstant.Equals(Nothing) Then
                Me.ConstantValue = ""
            End If

            writer.WriteStartElement("ConstantValue")
            writer.WriteValue(Me.ConstantValue)
            writer.WriteEndElement()
        Catch
            'This is also a new element and may not exist in previous saved searches
        End Try


        writer.WriteStartElement("OperatorType")
        writer.WriteValue(Me.OperatorType)
        writer.WriteEndElement()

        writer.WriteStartElement("ValueType")
        If Me.Value IsNot Nothing Then
            writer.WriteValue(Me.Value.GetType().FullName)
        End If
        writer.WriteEndElement()

        writer.WriteStartElement("Value")
        If Me.Value IsNot Nothing Then
            If TypeOf Me.Value Is Hashtable Then
                writer.WriteValue(CreateIDStringFromCollection(CType(Me.Value, Hashtable).Values))
            Else
                writer.WriteValue(Me.Value)
            End If
        End If
        writer.WriteEndElement()
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeObject()
        m_objValue = Nothing
        m_objField = Nothing
    End Sub

#End Region

End Class
