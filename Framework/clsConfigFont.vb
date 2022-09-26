Imports System.Drawing

<Serializable()> Public Class clsConfigFont

    Private m_strName As String
    Private m_intSize As Single
    Private m_intFontStyle As Integer
    Private m_intColor As Integer

    Public Sub New()
    End Sub

    Public Sub New(ByVal objFont As Font, ByVal intColor As Integer)
        m_strName = objFont.Name
        m_intSize = objFont.Size
        m_intFontStyle = objFont.Style
        m_intColor = intColor
    End Sub

    Public Sub New(ByVal strName As String, ByVal intSize As Single, _
    ByVal eFontStyle As FontStyle, ByVal intColor As Integer)
        m_strName = strName
        m_intSize = intSize
        m_intFontStyle = eFontStyle
        m_intColor = intColor
    End Sub

    Public Property Name() As String
        Get
            Return m_strName
        End Get
        Set(ByVal value As String)
            m_strName = value
        End Set
    End Property

    Public Property Size() As Single
        Get
            Return m_intSize
        End Get
        Set(ByVal value As Single)
            m_intSize = value
        End Set
    End Property

    Public Property FontStyle() As Integer
        Get
            Return m_intFontStyle
        End Get
        Set(ByVal value As Integer)
            m_intFontStyle = value
        End Set
    End Property

    Public Property Color() As Integer
        Get
            Return m_intColor
        End Get
        Set(ByVal value As Integer)
            m_intColor = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return m_strName & ", " & m_intSize.ToString & " pt"
    End Function
End Class
