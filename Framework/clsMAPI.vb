Imports System
Imports System.Collections
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Win32Mapi

#Region " Mapi Class "

    Public Class Mapi
        Implements IDisposable

#Region " Members "

        Private m_arrAttachments As ArrayList
        Private m_intError As Integer
        Private m_strFindSeed As String
        Private m_objLastMsg As MapiMessage
        Private m_objSB As StringBuilder
        Private m_objOrigin As MapiRecipDesc
        Private m_arrRecipients As ArrayList
        Private m_intSession As IntPtr
        Private m_intWinHandle As IntPtr
        Private m_arrErrorStrings As String()
        Private m_blnDisposedValue As Boolean = False
#End Region

#Region " Constants "

        Private Const cMAPI_BCC As Integer = 3
        Private Const cMAPI_BODY_AS_FILE As Integer = &H200
        Private Const MapiCC As Integer = 2
        Private Const MapiEnvOnly As Integer = &H40
        Private Const MapiExtendedUI As Integer = &H20
        Private Const MapiForceDownload As Integer = &H1000
        Private Const MapiGuaranteeFiFo As Integer = &H100
        Private Const MapiLogonUI As Integer = 1
        Private Const MapiLongMsgID As Integer = &H4000
        Private Const MapiNewSession As Integer = 2
        Private Const MapiORIG As Integer = 0
        Private Const MapiPasswordUI As Integer = &H20000
        Private Const MapiPeek As Integer = &H80
        Private Const MapiReceiptReq As Integer = 2
        Private Const MapiSent As Integer = 4
        Private Const MapiSuprAttach As Integer = &H800
        Private Const MapiTO As Integer = 1
        Private Const MapiUnread As Integer = 1
        Private Const MapiUnreadOnly As Integer = &H20
#End Region

#Region " Constructors "

        Public Sub New()
            m_arrAttachments = New ArrayList
            m_intError = 0
            m_objSB = New StringBuilder(600)
            m_objOrigin = New MapiRecipDesc
            m_arrRecipients = New ArrayList
            m_intSession = IntPtr.Zero
            m_intWinHandle = IntPtr.Zero
            m_arrErrorStrings = New String() {"OK [0]", "User abort [1]", "General MAPI failure [2]", _
                "MAPI login failure [3]", "Disk full [4]", "Insufficient memory [5]", "Access denied [6]", _
                "-unknown- [7]", "Too many sessions [8]", "Too many files were specified [9]", _
                "Too many recipients were specified [10]", "A specified attachment was not found [11]", _
                "Attachment open failure [12]", "Attachment write failure [13]", "Unknown recipient [14]", _
                "Bad recipient type [15]", "No messages [16]", "Invalid message [17]", "Text too large [18]", _
                "Invalid session [19]", "Type not supported [20]", "A recipient was specified ambiguously [21]", _
                "Message in use [22]", "Network failure [23]", "Invalid edit fields [24]", _
                "Invalid recipients [25]", "Not supported [26]"}
        End Sub
#End Region

#Region " DLL Imports "

        <DllImport("MAPI32.DLL", CharSet:=CharSet.Ansi)> _
        Private Shared Function MAPIAddress(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal caption As String, ByVal editfld As Integer, ByVal labels As String, ByVal recipcount As Integer, ByVal ptrrecips As IntPtr, ByVal flg As Integer, ByVal rsv As Integer, ByRef newrec As Integer, ByRef ptrnew As IntPtr) As Integer
        End Function

        <DllImport("MAPI32.DLL", CharSet:=CharSet.Ansi)> _
        Private Shared Function MAPIDeleteMail(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal id As String, ByVal flg As Integer, ByVal rsv As Integer) As Integer
        End Function

        <DllImport("MAPI32.DLL", CharSet:=CharSet.Ansi)> _
        Private Shared Function MAPIFindNext(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal typ As String, ByVal seed As String, ByVal flg As Integer, ByVal rsv As Integer, ByVal id As StringBuilder) As Integer
        End Function

        <DllImport("MAPI32.DLL")> _
        Private Shared Function MAPIFreeBuffer(ByVal ptr As IntPtr) As Integer
        End Function

        <DllImport("MAPI32.DLL")> _
        Private Shared Function MAPILogoff(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal flg As Integer, ByVal rsv As Integer) As Integer
        End Function

        <DllImport("MAPI32.DLL", CharSet:=CharSet.Ansi)> _
        Private Shared Function MAPILogon(ByVal hwnd As IntPtr, ByVal prf As String, ByVal pw As String, ByVal flg As Integer, ByVal rsv As Integer, ByRef sess As IntPtr) As Integer
        End Function

        <DllImport("MAPI32.DLL", CharSet:=CharSet.Ansi)> _
        Private Shared Function MAPIReadMail(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal id As String, ByVal flg As Integer, ByVal rsv As Integer, ByRef ptrmsg As IntPtr) As Integer
        End Function

        <DllImport("MAPI32.DLL")> _
        Private Shared Function MAPISendMail(ByVal sess As IntPtr, ByVal hwnd As IntPtr, ByVal message As MapiMessage, ByVal flg As Integer, ByVal rsv As Integer) As Integer
        End Function
#End Region

#Region " Enumerations "

        Public Enum enumRecipType
            [TO] = 1
            CC = 2
            BCC = 3
        End Enum
#End Region

#Region " Methods "

#Region " Allocation Methods "

        Private Function AllocAttachs(ByRef intFileCount As Integer) As IntPtr
            intFileCount = 0

            If (m_arrAttachments Is Nothing) Then
                Return IntPtr.Zero
            End If

            If ((m_arrAttachments.Count <= 0) OrElse (m_arrAttachments.Count > 100)) Then
                Return IntPtr.Zero
            End If

            Dim objType As Type = GetType(MapiFileDesc)
            Dim intASize As Integer = Marshal.SizeOf(objType)
            Dim intPtrA As IntPtr = Marshal.AllocHGlobal(CInt((m_arrAttachments.Count * intASize)))
            Dim objMFD As New MapiFileDesc

            objMFD.intPosition = -1

            Dim intRunPtr As Integer = CInt(intPtrA)
            Dim intLoop As Integer

            For intLoop = 0 To m_arrAttachments.Count - 1
                Dim strPath As String = TryCast(m_arrAttachments.Item(intLoop), String)
                objMFD.strName = IO.Path.GetFileName(strPath)
                objMFD.strPath = strPath
                Marshal.StructureToPtr(objMFD, CType(intRunPtr, IntPtr), False)
                intRunPtr = (intRunPtr + intASize)
            Next intLoop

            intFileCount = m_arrAttachments.Count

            Return intPtrA
        End Function

        Private Function AllocOrigin() As IntPtr
            m_objOrigin.intRecipClass = 0

            Dim objRType As Type = GetType(MapiRecipDesc)
            Dim intPtrO As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(objRType))

            Marshal.StructureToPtr(m_objOrigin, intPtrO, False)

            Return intPtrO
        End Function

        Private Function AllocRecips(ByRef intRecipCount As Integer) As IntPtr
            intRecipCount = 0

            If (m_arrRecipients.Count = 0) Then
                Return IntPtr.Zero
            End If

            Dim objRType As Type = GetType(MapiRecipDesc)
            Dim intRSize As Integer = Marshal.SizeOf(objRType)
            Dim intPtrR As IntPtr = Marshal.AllocHGlobal(CInt((m_arrRecipients.Count * intRSize)))
            Dim intRunPtr As Integer = CInt(intPtrR)
            Dim intLoop As Integer

            For intLoop = 0 To m_arrRecipients.Count - 1
                Marshal.StructureToPtr(TryCast(m_arrRecipients.Item(intLoop), MapiRecipDesc), CType(intRunPtr, IntPtr), False)
                intRunPtr = (intRunPtr + intRSize)
            Next intLoop

            intRecipCount = m_arrRecipients.Count

            Return intPtrR
        End Function

        Private Sub Dealloc()
            Dim intRunPtr As Integer
            Dim intLoop As Integer
            Dim objRType As Type = GetType(MapiRecipDesc)
            Dim intRSize As Integer = Marshal.SizeOf(objRType)

            If (m_objLastMsg.intOriginator <> IntPtr.Zero) Then
                Marshal.DestroyStructure(m_objLastMsg.intOriginator, objRType)
                Marshal.FreeHGlobal(m_objLastMsg.intOriginator)
            End If

            If (m_objLastMsg.intRecips <> IntPtr.Zero) Then
                intRunPtr = CInt(m_objLastMsg.intRecips)

                For intLoop = 0 To m_objLastMsg.intRecipCount - 1
                    Marshal.DestroyStructure(CType(intRunPtr, IntPtr), objRType)
                    intRunPtr = (intRunPtr + intRSize)
                Next intLoop

                Marshal.FreeHGlobal(m_objLastMsg.intRecips)
            End If

            If (m_objLastMsg.intFiles <> IntPtr.Zero) Then
                Dim objFType As Type = GetType(MapiFileDesc)
                Dim intFSize As Integer = Marshal.SizeOf(objFType)

                intRunPtr = CInt(m_objLastMsg.intFiles)

                For intLoop = 0 To m_objLastMsg.inFileCount - 1
                    Marshal.DestroyStructure(CType(intRunPtr, IntPtr), objFType)
                    intRunPtr = (intRunPtr + intFSize)
                Next intLoop

                Marshal.FreeHGlobal(m_objLastMsg.intFiles)
            End If
        End Sub
#End Region

#Region " Inbox Methods "

        Public Function Delete(ByVal strID As String) As Boolean
            m_intError = Mapi.MAPIDeleteMail(m_intSession, m_intWinHandle, strID, 0, 0)
            Return (m_intError = 0)
        End Function

        Private Sub GetAttachNames(ByRef arrAttachments As MailAttach())
            arrAttachments = New MailAttach(m_objLastMsg.inFileCount - 1) {}

            Dim objFDType As Type = GetType(MapiFileDesc)
            Dim intFDSize As Integer = Marshal.SizeOf(objFDType)
            Dim objFDTmp As New MapiFileDesc
            Dim intRunPtr As Integer = CInt(m_objLastMsg.intFiles)
            Dim intLoop As Integer

            For intLoop = 0 To m_objLastMsg.inFileCount - 1
                Marshal.PtrToStructure(CType(intRunPtr, IntPtr), objFDTmp)
                intRunPtr = (intRunPtr + intFDSize)
                arrAttachments(intLoop) = New MailAttach

                If (objFDTmp.intFlags = 0) Then
                    arrAttachments(intLoop).intPosition = objFDTmp.intPosition
                    arrAttachments(intLoop).strName = objFDTmp.strName
                    arrAttachments(intLoop).strPath = objFDTmp.strPath
                End If
            Next intLoop
        End Sub

        Public Function [Next](ByRef objEnv As MailEnvelop) As Boolean
            m_intError = Mapi.MAPIFindNext(m_intSession, m_intWinHandle, Nothing, m_strFindSeed, &H4000, 0, m_objSB)

            If (m_intError <> 0) Then
                Return False
            End If

            m_strFindSeed = m_objSB.ToString
            Dim intPtrMsg As IntPtr = IntPtr.Zero

            m_intError = Mapi.MAPIReadMail(m_intSession, m_intWinHandle, m_strFindSeed, &H8C0, 0, (intPtrMsg))
            If ((m_intError <> 0) OrElse (intPtrMsg = IntPtr.Zero)) Then
                Return False
            End If

            m_objLastMsg = New MapiMessage
            Marshal.PtrToStructure(intPtrMsg, m_objLastMsg)
            Dim intOrig As New MapiRecipDesc
            If (m_objLastMsg.intOriginator <> IntPtr.Zero) Then
                Marshal.PtrToStructure(m_objLastMsg.intOriginator, intOrig)
            End If

            objEnv.strID = m_strFindSeed
            objEnv.dtDate = DateTime.ParseExact(m_objLastMsg.strDateReceived, "yyyy/MM/dd HH:mm", DateTimeFormatInfo.InvariantInfo)
            objEnv.strSubject = m_objLastMsg.strSubject
            objEnv.strFrom = intOrig.strName
            objEnv.blnUnread = ((m_objLastMsg.intFlags And 1) <> 0)
            objEnv.intAtts = m_objLastMsg.inFileCount

            m_intError = Mapi.MAPIFreeBuffer(intPtrMsg)

            Return (m_intError = 0)
        End Function

        Public Function Read(ByVal strID As String, ByRef arrAttachments As MailAttach()) As String
            arrAttachments = Nothing
            Dim intPtrMsg As IntPtr = IntPtr.Zero

            m_intError = Mapi.MAPIReadMail(m_intSession, m_intWinHandle, strID, &H880, 0, (intPtrMsg))
            If ((m_intError <> 0) OrElse (intPtrMsg = IntPtr.Zero)) Then
                Return Nothing
            End If

            m_objLastMsg = New MapiMessage
            Marshal.PtrToStructure(intPtrMsg, m_objLastMsg)
            If (((m_objLastMsg.inFileCount > 0) AndAlso (m_objLastMsg.inFileCount < 100)) AndAlso (m_objLastMsg.intFiles <> IntPtr.Zero)) Then
                GetAttachNames(arrAttachments)
            End If

            Mapi.MAPIFreeBuffer(intPtrMsg)

            Return m_objLastMsg.strNoteText
        End Function

        Private Function SaveAttachByName(ByVal strName As String, ByVal strSavePath As String) As Boolean
            Dim blnF As Boolean = True
            Dim objFDType As Type = GetType(MapiFileDesc)
            Dim intFDSize As Integer = Marshal.SizeOf(objFDType)
            Dim objFDTmp As New MapiFileDesc
            Dim intRunPtr As Integer = CInt(m_objLastMsg.intFiles)
            Dim intLoop As Integer

            For intLoop = 0 To m_objLastMsg.inFileCount - 1
                Marshal.PtrToStructure(CType(intRunPtr, IntPtr), objFDTmp)
                intRunPtr = (intRunPtr + intFDSize)

                If ((objFDTmp.intFlags = 0) AndAlso (Not objFDTmp.strName Is Nothing)) Then
                    Try
                        If (strName = objFDTmp.strName) Then
                            If File.Exists(strSavePath) Then
                                File.Delete(strSavePath)
                            End If
                            File.Move(objFDTmp.strPath, strSavePath)
                        End If
                    Catch exception1 As Exception
                        blnF = False
                        m_intError = 13
                    End Try
                    Try
                        File.Delete(objFDTmp.strPath)
                    Catch exception2 As Exception
                    End Try
                End If
            Next intLoop

            Return blnF
        End Function

        Public Function SaveAttachm(ByVal strID As String, ByVal strName As String, ByVal strSavePath As String) As Boolean
            Dim intPtrMsg As IntPtr = IntPtr.Zero

            m_intError = Mapi.MAPIReadMail(m_intSession, m_intWinHandle, strID, &H80, 0, intPtrMsg)

            If ((m_intError <> 0) OrElse (intPtrMsg = IntPtr.Zero)) Then
                Return False
            End If

            m_objLastMsg = New MapiMessage
            Marshal.PtrToStructure(intPtrMsg, m_objLastMsg)

            Dim blnF As Boolean = False
            If (((m_objLastMsg.inFileCount > 0) AndAlso (m_objLastMsg.inFileCount < 100)) AndAlso (m_objLastMsg.intFiles <> IntPtr.Zero)) Then
                blnF = SaveAttachByName(strName, strSavePath)
            End If

            Mapi.MAPIFreeBuffer(intPtrMsg)

            Return blnF
        End Function
#End Region

#Region " Logon/Logoff "

        Public Function Logon(ByVal hwnd As IntPtr) As Boolean
            m_intWinHandle = hwnd
            m_intError = Mapi.MAPILogon(hwnd, Nothing, Nothing, 0, 0, m_intSession)

            If (m_intError <> 0) Then
                m_intError = Mapi.MAPILogon(hwnd, Nothing, Nothing, 1, 0, m_intSession)
            End If

            Return (m_intError = 0)
        End Function

        Public Sub Logoff()
            If (m_intSession <> IntPtr.Zero) Then
                m_intError = Mapi.MAPILogoff(m_intSession, m_intWinHandle, 0, 0)
                m_intSession = IntPtr.Zero
            End If
        End Sub
#End Region

#Region " Send Email Methods "

        Public Sub AddRecip(ByVal strName As String, ByVal strAddr As String, ByVal eType As enumRecipType)
            Dim objDest As New MapiRecipDesc

            objDest.intRecipClass = eType
            objDest.strName = strName
            objDest.strAddress = strAddr

            m_arrRecipients.Add(objDest)
        End Sub

        Public Sub Attach(ByVal strFilePath As String)
            m_arrAttachments.Add(strFilePath)
        End Sub

        Public Sub Reset()
            m_strFindSeed = Nothing
            m_objOrigin = New MapiRecipDesc
            m_arrRecipients.Clear()
            m_arrAttachments.Clear()
            m_objLastMsg = Nothing
        End Sub

        Public Function Send(ByVal strSubject As String, ByVal strTxt As String) As Boolean
            m_objLastMsg = New MapiMessage
            m_objLastMsg.strSubject = strSubject
            m_objLastMsg.strNoteText = strTxt
            m_objLastMsg.intOriginator = AllocOrigin()
            m_objLastMsg.intRecips = AllocRecips(m_objLastMsg.intRecipCount)
            m_objLastMsg.intFiles = AllocAttachs(m_objLastMsg.inFileCount)

            m_intError = Mapi.MAPISendMail(m_intSession, m_intWinHandle, m_objLastMsg, 0, 0)

            Dealloc()
            Reset()

            Return (m_intError = 0)
        End Function

        Public Sub SetSender(ByVal strSName As String, ByVal strSAddress As String)
            m_objOrigin.strName = strSName
            m_objOrigin.strAddress = strSAddress
        End Sub

        Public Function GetAddresses(ByVal strLabel As String, ByRef strName As String) As Boolean
            Try
                Dim intNewRec As Integer = 0
                Dim intPtrNew As IntPtr = IntPtr.Zero

                m_intError = Mapi.MAPIAddress(m_intSession, m_intWinHandle, Nothing, 1, strLabel, 0, IntPtr.Zero, 0, 0, intNewRec, intPtrNew)
                If ((intNewRec < 1) OrElse (intPtrNew = IntPtr.Zero)) Then
                    Return False
                End If

                Dim objRecip(intNewRec) As MapiRecipDesc

                Dim intCurPtr As IntPtr = intPtrNew
                For intLoop As Integer = 0 To intNewRec - 1
                    objRecip(intLoop) = New MapiRecipDesc

                    Marshal.PtrToStructure(intCurPtr, objRecip(intLoop))

                    If strName Is Nothing OrElse strName.Length = 0 Then
                        strName = objRecip(intLoop).strName
                    Else
                        strName &= "; " & objRecip(intLoop).strName
                    End If

                    intCurPtr = CType(CInt(intCurPtr) + Marshal.SizeOf(objRecip(intLoop)), IntPtr)
                Next

                Mapi.MAPIFreeBuffer(intPtrNew)

                Return True
            Catch ex As Exception
                Throw New Exception("There was a problem using MAPI to select an email address:" & vbCrLf & vbCrLf & ex.Message)
            End Try
        End Function
#End Region

#Region " Interpret Error "

        Public Function [Error]() As String
            If (m_intError <= &H1A) Then
                Return m_arrErrorStrings(m_intError)
            End If

            Return ("?unknown? [" & m_intError.ToString & "]")
        End Function
#End Region

#Region " IDisposable Support "

        Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
            If Not m_blnDisposedValue Then
                If blnDisposing Then
                    If Not m_arrAttachments Is Nothing Then
                        m_arrAttachments.Clear()
                        m_arrAttachments = Nothing
                    End If

                    m_objLastMsg = Nothing
                    m_objOrigin = Nothing
                    m_objSB = Nothing
                    m_arrErrorStrings = Nothing

                    If Not m_arrRecipients Is Nothing Then
                        m_arrRecipients.Clear()
                        m_arrRecipients = Nothing
                    End If
                End If
            End If
            m_blnDisposedValue = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

#End Region

    End Class
#End Region

#Region " MailAttach Class "

    Public Class MailAttach
        Public strName As String
        Public strPath As String
        Public intPosition As Integer
    End Class
#End Region

#Region " MailEnvelop Class "

    Public Class MailEnvelop
        Public intAtts As Integer
        Public dtDate As DateTime
        Public strFrom As String
        Public strID As String
        Public strSubject As String
        Public blnUnread As Boolean
    End Class
#End Region

#Region " MapiFileDesc Class "

    <StructLayout(LayoutKind.Sequential)> _
    Public Class MapiFileDesc
        Public intReserved As Integer
        Public intFlags As Integer
        Public intPosition As Integer
        Public strPath As String
        Public strName As String
        Public intType As IntPtr
    End Class
#End Region

#Region " MapiMessage Class "

    <StructLayout(LayoutKind.Sequential)> _
    Public Class MapiMessage
        Public intReserved As Integer
        Public strSubject As String
        Public strNoteText As String
        Public strMessageType As String
        Public strDateReceived As String
        Public strConversationID As String
        Public intFlags As Integer
        Public intOriginator As IntPtr
        Public intRecipCount As Integer
        Public intRecips As IntPtr
        Public inFileCount As Integer
        Public intFiles As IntPtr
    End Class
#End Region

#Region " MapiRecipDesc Class "

    <StructLayout(LayoutKind.Sequential)> _
    Public Class MapiRecipDesc
        Public intReserved As Integer
        Public intRecipClass As Integer
        Public strName As String
        Public strAddress As String
        Public intEIDSize As Integer
        Public intEntryID As IntPtr
    End Class
#End Region

End Namespace
