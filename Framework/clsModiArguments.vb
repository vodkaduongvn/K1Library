Public Class clsModiArguments
    Private blnWithStraightenImage As Boolean = True
    Private blnWithAutoRotation As Boolean = False
    Private blnTextMappedPDF As TextMappedOptions = TextMappedOptions.Mapped
    Private enumOCRLanguageOption As OcrLanguageOptions = OcrLanguageOptions.LangSysdefault

    Public Property Language() As OcrLanguageOptions
        Get
            Return enumOCRLanguageOption
        End Get
        Set(ByVal value As OcrLanguageOptions)
            enumOCRLanguageOption = value
        End Set
    End Property

    Public Property WithAutoRotation() As Boolean
        Get
            Return blnWithAutoRotation
        End Get
        Set(ByVal value As Boolean)
            blnWithAutoRotation = value
        End Set
    End Property

    Public Property WithStraightenImage() As Boolean
        Get
            Return blnWithStraightenImage
        End Get
        Set(ByVal value As Boolean)
            blnWithStraightenImage = value
        End Set
    End Property

    Public Property TextMappedPDF() As TextMappedOptions
        Get
            Return blnTextMappedPDF
        End Get
        Set(ByVal value As TextMappedOptions)
            blnTextMappedPDF = value
        End Set
    End Property
End Class
