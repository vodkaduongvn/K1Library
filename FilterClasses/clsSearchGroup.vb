#Region " File Information "

'==============================================================================
' This class represents a grouping in a search filter (ie. the parentheses)
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
Imports System.Reflection

<Serializable()> Public Class clsSearchGroup
    Inherits clsSearchObjBase
    Implements IXmlSerializable

#Region " Members "

    Private m_colSearchObjs As List(Of clsSearchObjBase)
#End Region

#Region " Constructors "

    Public Sub New()
    End Sub

    ''' <summary>
    ''' Creates a new search group
    ''' </summary>
    ''' <param name="eOpType">AND, OR, ANDNOT, ORNOT, [Nothing]</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType)
        MyBase.New(eOpType)
        m_colSearchObjs = New List(Of clsSearchObjBase)
    End Sub

    ''' <summary>
    ''' Creates a new search group
    ''' </summary>
    ''' <param name="eOpType">AND, OR, ANDNOT, ORNOT, [Nothing]</param>
    ''' <param name="colSearchObjs">Collection of search groups and search elements</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal eOpType As clsSearchFilter.enumOperatorType, _
    ByVal colSearchObjs As List(Of clsSearchObjBase))
        MyBase.New(eOpType)
        m_colSearchObjs = colSearchObjs
    End Sub
#End Region

#Region " Properties "

    ''' <summary>
    ''' Collection of search groups and search elements
    ''' </summary>
    Public Property SearchObjs() As List(Of clsSearchObjBase)
        Get
            Return m_colSearchObjs
        End Get
        Set(ByVal value As List(Of clsSearchObjBase))
            m_colSearchObjs = value
        End Set
    End Property
#End Region

#Region " Methods "

    Public Shared Function CreateRangeGroup(ByVal eOPType As clsSearchFilter.enumOperatorType, _
    ByVal objField As clsField, ByVal objValue1 As Object, ByVal objValue2 As Object, _
    ByVal strParentRef As String) As clsSearchGroup
        Dim objSG As New clsSearchGroup(eOPType)
        Dim colSOs As New List(Of clsSearchObjBase)
        objSG.SearchObjs = colSOs

        If Not objValue1 Is Nothing Then
            colSOs.Add(New clsSearchElement(clsSearchFilter.enumOperatorType.NONE, _
                strParentRef & "." & objField.DatabaseName, _
                clsSearchFilter.enumComparisonType.GREATER_THAN_EQUAL, objValue1))
        End If

        If Not objValue2 Is Nothing Then
            colSOs.Add(New clsSearchElement(clsSearchFilter.enumOperatorType.AND, _
                strParentRef & "." & objField.DatabaseName, _
                clsSearchFilter.enumComparisonType.LESS_THAN_EQUAL, objValue2))
        End If

        If colSOs.Count = 0 Then
            Return Nothing
        Else
            Return objSG
        End If
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeObject()
        If Not m_colSearchObjs Is Nothing Then
            m_colSearchObjs.Clear()
            m_colSearchObjs = Nothing
        End If
    End Sub
#End Region

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
        Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
        reader.ReadStartElement()

        Dim colPropInfo As PropertyInfo() = Me.GetType.GetProperties( _
            BindingFlags.Public Or BindingFlags.Instance)

        While reader.IsStartElement()
            Dim strType As String = reader.Name

            Dim objType As Type = Nothing

            Select Case strType.ToUpper
                Case "CLSSEARCHELEMENT"
                    objType = GetType(clsSearchElement)

                Case "CLSSEARCHGROUP"
                    objType = GetType(clsSearchGroup)

                Case Else
                    Dim objPI As PropertyInfo = Me.GetType.GetProperty(strType)
                    Dim objPropType As Type = objPI.PropertyType

                    If objPropType.BaseType IsNot Nothing AndAlso _
                    objPropType.BaseType.Equals(GetType(System.Enum)) Then
                        objPropType = GetType(System.Int32)
                    End If

                    Dim objValue As Object = reader.ReadElementContentAs(objPropType, Nothing)
                    objPI.SetValue(Me, objValue, Nothing)

            End Select

            If objType IsNot Nothing Then
                Dim objXML As New XmlSerializer(objType)

                Dim obj As Object = objXML.Deserialize(reader)
                If m_colSearchObjs Is Nothing Then
                    m_colSearchObjs = New List(Of clsSearchObjBase)
                End If
                m_colSearchObjs.Add(CType(obj, clsSearchObjBase))
            End If
        End While

        reader.ReadEndElement()
    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
        Dim colPropInfo As PropertyInfo() = Me.GetType.GetProperties( _
            BindingFlags.Public Or BindingFlags.Instance)

        For Each objPropInfo As PropertyInfo In colPropInfo
            If objPropInfo.PropertyType.IsSerializable Then
                Dim strColName As String = objPropInfo.Name

                If Not objPropInfo.PropertyType.Equals(m_colSearchObjs.GetType) Then
                    writer.WriteStartElement(objPropInfo.Name)
                    Dim objValue As Object = Convert.ChangeType(objPropInfo.GetValue(Me, Nothing), _
                        objPropInfo.PropertyType)

                    If objPropInfo.PropertyType.BaseType IsNot Nothing AndAlso _
                    objPropInfo.PropertyType.BaseType.Equals(GetType(System.Enum)) Then
                        objValue = CInt(objValue)
                    End If

                    writer.WriteValue(objValue)
                    writer.WriteEndElement()
                End If
            End If
        Next

        If m_colSearchObjs IsNot Nothing Then
            For Each objSearchObj As clsSearchObjBase In m_colSearchObjs
                Dim objXML As New XmlSerializer(objSearchObj.GetType())
                objXML.Serialize(writer, objSearchObj)
            Next
        End If
    End Sub
End Class
