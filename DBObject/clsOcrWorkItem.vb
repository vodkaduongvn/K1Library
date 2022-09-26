Namespace DBObject
    Public Class clsOcrWorkItem

        Property Id() As Integer

        Property EdocId() As Integer

        Property IsDone() As Boolean

        Property SubmittedDate As DateTime

        Property ProcessedDate As Nullable(Of DateTime)

        Property FileFormat() As String

        Property TextOption() As String

        Property AutoRotate() As Boolean

        Property LanguageId() As Integer

    End Class
End Namespace