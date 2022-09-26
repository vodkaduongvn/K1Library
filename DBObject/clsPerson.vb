#Region " File Information "

'=====================================================================
' This class represents the table Person in the Database.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' Auto      8/07/2004    Implemented.
'=====================================================================
#End Region

#End Region

Public Class clsPerson
	Inherits clsDBObjBase

#Region " Members "

    Private m_strFirstName As String
    Private m_strLastName As String
    Private m_intEntityID As Integer
    Private m_intLocationID As Integer
    Private m_strWorkPhone As String
    Private m_strWorkFax As String
    Private m_strWorkEmail As String
    Private m_blnIsActive As Boolean

#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
        MyBase.New(objDR, objDB)

        m_strFirstName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cFIRSTNAME, ""), String)
        m_strLastName = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cLASTNAME, ""), String)
        m_intEntityID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cENTITYID, 0), Integer)
        m_intLocationID = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cLOCATIONID, 0), Integer)
        m_strWorkPhone = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cWORKPHONE, ""), String)
        m_strWorkFax = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cWORKFAX, ""), String)
        m_strWorkEmail = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cWORKEMAIL, ""), String)
        m_blnIsActive = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.Person.cISACTIVE, True), Boolean)
    End Sub
#End Region

#Region " Properties "

    Public ReadOnly Property FirstName() As String
        Get
            Return m_strFirstName
        End Get
    End Property

    Public ReadOnly Property LastName() As String
        Get
            Return m_strLastName
        End Get
    End Property

    Public ReadOnly Property EntityID() As Integer
        Get
            Return m_intEntityID
        End Get
    End Property

    Public ReadOnly Property LocationID() As Integer
        Get
            Return m_intLocationID
        End Get
    End Property

    Public ReadOnly Property WorkPhone() As String
        Get
            Return m_strWorkPhone
        End Get
    End Property

    Public ReadOnly Property WorkFax() As String
        Get
            Return m_strWorkFax
        End Get
    End Property

    Public ReadOnly Property WorkEmail() As String
        Get
            Return m_strWorkEmail
        End Get
    End Property

    Public ReadOnly Property IsActive() As Boolean
        Get
            Return m_blnIsActive
        End Get
    End Property

#End Region

#Region " GetItem "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsPerson
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cPERSON, intID)

            Return New clsPerson(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
