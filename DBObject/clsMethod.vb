#Region " File Information "

'=====================================================================
' This class represents the table Method in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsMethod
    Inherits clsDBObjBase

#Region " Members "

    Private m_intButtonID As Integer
    Private m_objButton As clsButton
    Private m_eUIID As enumMethods
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_intButtonID = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)
        m_intButtonID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Method.cBUTTONID, clsDBConstants.cintNULL), Integer)
        m_eUIID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Method.cUIID, clsDBConstants.cintNULL), enumMethods)
    End Sub
#End Region

#Region " Enumerations "

    Public Enum enumMethods
        cSELECT = 1
        cSEARCH = 2
        cADD = 3
        cMODIFY = 4
        cDELETE = 5
        cPROCESS = 6
        cMOVE = 8
        cSORT = 9
        cREQUEST = 10
        cPRINT = 11
        cVIEW = 12
        cCLONE = 13
        cCHECKOUT = 16
        cIMPORT = 17
        cEXTERNALID = 19
        cBARCODE = 20
        cMETADATA = 21
        cRANGE = 22
        cSAVESEARCH = 23
        cREPLAY = 24
        cEXPORT = 25
        cNEWPART = 26
        cBOOLEAN = 27
        cPICKLIST = 28
    End Enum
#End Region

#Region " Properties "

    Public ReadOnly Property Button() As clsButton
        Get
            If m_objButton Is Nothing Then
                If Not m_intButtonID = clsDBConstants.cintNULL Then
                    m_objButton = clsButton.GetItem(m_intButtonID, Me.Database)
                End If
            End If
            Return m_objButton
        End Get
    End Property

    Public ReadOnly Property UIID() As enumMethods
        Get
            Return m_eUIID
        End Get
    End Property

    Public ReadOnly Property MethodText() As String
        Get
            Return m_objDB.SysInfo.GetMethodString(m_intID)
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsMethod
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cMETHOD, intID)

            Return New clsMethod(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1DualKeyDictionary(Of clsMethod, Integer)
        Dim colMethods As FrameworkCollections.K1DualKeyDictionary(Of clsMethod, Integer)
        Dim objMethod As clsMethod

        Try
            Dim strSP As String = clsDBConstants.Tables.cMETHOD & clsDBConstants.StoredProcedures.cGETLIST

            Dim objDT As DataTable = objDB.GetDataTable(strSP)

            colMethods = New FrameworkCollections.K1DualKeyDictionary(Of clsMethod, Integer)
            For Each objDR As DataRow In objDT.Rows
                objMethod = New clsMethod(objDR, objDB)
                colMethods.Add(CStr(objMethod.UIID), objMethod)
            Next

            '2017-07-12 -- James & Peter -- Bug Fix for DataTable not releasing the SQL connection
            If objDT IsNot Nothing Then
                objDT.Dispose()
                objDT = Nothing
            End If

            Return colMethods
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Security "

    Public Function HasAccess(ByVal objTable As clsTable) As Boolean
        Dim objTM As clsTableMethod = objTable.TableMethods(KeyID)

        If objTM IsNot Nothing AndAlso _
        m_objDB.Profile.HasAccess(objTM.SecurityID) AndAlso _
        m_objDB.Profile.LinkMethods(CStr(objTM.ID)) IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_objButton Is Nothing Then
                m_objButton.Dispose()
                m_objButton = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
