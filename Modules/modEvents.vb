Public Module modEvents

    Public Event Refresh(ByVal dtRefreshDate As Date)
    Public Event ServerNotInitialized()
    Public Event DRMLocked()
    Public Event ForceLogOut()
    Public Event SessionExpired()

    Public Sub RaiseRefreshEvent(ByVal dtRefreshDate As Date)
        RaiseEvent Refresh(dtRefreshDate)
    End Sub

    Public Sub RaiseServerNotInitialized()
        RaiseEvent ServerNotInitialized()
    End Sub

    Public Sub RaiseDRMLocked()
        RaiseEvent DRMLocked()
    End Sub

    Public Sub RaiseForceLogOut()
        RaiseEvent ForceLogOut()
    End Sub

    Public Sub RaiseSessionExpired()
        RaiseEvent SessionExpired()
    End Sub



End Module
