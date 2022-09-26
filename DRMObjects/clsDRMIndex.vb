Public Class clsDRMIndex
    Inherits clsDRMBase

    Private m_strTable As String
    Private m_strName As String
    Private m_arrFields As String

    Public Sub New(ByVal objDB As clsDB)
        MyBase.New(objDB, "", clsDBConstants.cintNULL, clsDBConstants.cintNULL)
    End Sub

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDRMObject()
    End Sub

#End Region

End Class
