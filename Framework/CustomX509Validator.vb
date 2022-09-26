Imports System.IdentityModel.Selectors

Public Class CustomX509Validator
    Inherits X509CertificateValidator

    Public Overrides Sub Validate(ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate2)
        Try
            If Not certificate.SubjectName.Name = "CN=K1ServiceCert" _
            OrElse Not certificate.Thumbprint = "7368FB53D522081ACED05ED1214B8A1A09BB209B" _
            OrElse Not certificate.Issuer = "CN=Knowledgeone Corp Root Authority, OU=Knowledgeone Coporation, C=AU" Then
                '-- Validate all certificates except K1 Default Cert
                Dim objValidator As X509CertificateValidator = ChainTrust
                objValidator.Validate(certificate)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class
