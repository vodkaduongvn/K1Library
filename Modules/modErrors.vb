Public Module modErrors

    <Flags()> _
    Public Enum ErrorNumber As Integer
        Application = &H1
        Web_Service = &H2
        DRM_Locked = &H4
        Invalid_Schema = &H8
        No_Broker = &H10
        Server_Not_Initialised = &H20
        Session_Expired = &H40
        No_Available_Licence = &H80
        No_Licence = &H100
        Old_Version = &H200
        Proxy_Auth_Required = &H400
        Invalid_Login = &H800
        Data_Not_Found = &H900
    End Enum

End Module
