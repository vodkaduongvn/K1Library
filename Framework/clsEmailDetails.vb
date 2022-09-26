Imports Aspose.Email
Public Class clsEmailDetails
    Public ReadOnly DateRecieved As String
    Public ReadOnly Sender As String
    Public ReadOnly Author As String
    Public ReadOnly Recipients As String
    Public ReadOnly Subject As String
    Public ReadOnly Abstract As String
    Public ReadOnly CCList As String
    Public ReadOnly BCCList As String
    Public ReadOnly FilePath As String

    Public Sub New(strFilename As String)
        Dim objLicense As License = New License()

        'Pass only the name of the license file embedded in the assembly
        objLicense.SetLicense("Aspose.Email.lic")

        FilePath = strFilename
        'Load MSG file into MailMessge object
        Using objMessage = MailMessage.Load(strFilename)
            'set published date to date received
            If IsDate(objMessage.Headers("Date")) Then
                DateRecieved = CDate(objMessage.Headers("Date")).ToString("MMM dd, yyyy HH:mm:ss")
            End If

            If objMessage.From IsNot Nothing Then
                Sender = objMessage.From.DisplayName & " <" & objMessage.From.Address & ">"
                Author = objMessage.From.DisplayName
            End If

            'build string of recipients
            Dim strRecipient As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.To
                strRecipient &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strRecipient.Equals(String.Empty) Then
                Recipients = strRecipient
            End If

            If Not objMessage.Subject.Equals(String.Empty) Then
                Subject = objMessage.Subject
                Abstract = objMessage.Subject
            End If

            'build string of CC
            Dim strCC As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.CC
                strCC &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strCC.Equals(String.Empty) Then
                CCList = strCC
            End If

            'build string of BCC
            Dim strBCC As String = String.Empty

            For Each objEmailAddress As MailAddress In objMessage.Bcc
                strBCC &= objEmailAddress.DisplayName & " <" & objEmailAddress.Address & ">; "
            Next

            If Not strBCC.Equals(String.Empty) Then
                BCCList = strBCC
            End If
        End Using
    End Sub
End Class

Public Class clsButtonEmailDetails
    Inherits clsEmailDetails

    Public ReadOnly OriginalPath As String
    Public ReadOnly ProductType As Integer

    Public Sub New(strFilename As String, strOriginalPath As String, intProductType As Integer)
        MyBase.New(strFilename)

        OriginalPath = strOriginalPath
        ProductType = intProductType
    End Sub
End Class
