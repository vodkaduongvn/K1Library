Public Class clsEDOCArchive
    Inherits clsDBObjBase

#Region " Members "

    Private m_strConnection As String
    Private m_strServer As String
    Private m_strDatabase As String
    Private m_blnCurrent As Boolean
    Private m_intRows As Integer
    Private m_strSize As String
    Private m_strConfigServer As String
    Private m_strConfigDatabase As String
    Private m_blnValid As Boolean

#End Region

#Region " Constructors "

    Public Sub New(ByVal objDB As clsDB, ByVal objDR As DataRow)
        MyBase.New(objDB, CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.cID, 0), Integer), clsDBConstants.cstrNULL, clsDBConstants.cintNULL, clsDBConstants.cintNULL)

        m_strConnection = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Archive.cCONNECTIONSTRING, ""), String)
        m_blnCurrent = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.K1Archive.cCURRENT, False), Boolean)
        m_blnValid = False

        m_strServer = String.Empty
        m_strDatabase = String.Empty
        m_strConfigDatabase = String.Empty
        m_strConfigServer = String.Empty
        m_intRows = 0
        m_strSize = "0"

        If Not m_strConnection.Equals(String.Empty) Then
            Dim objEncryption As New clsEncryption(True)
            Dim strConnection As String = objEncryption.Decrypt(m_strConnection)

            Dim objSQL As New SqlClient.SqlConnectionStringBuilder(strConnection)
            m_strServer = objSQL.DataSource
            m_strDatabase = objSQL.InitialCatalog

            Try
                Dim objConnection As New SqlClient.SqlConnection(strConnection)
                objConnection.Open()

                Dim objDT As DataTable = Nothing
                Dim objCommand As SqlClient.SqlCommand = objConnection.CreateCommand()
                objCommand.CommandText = "SELECT * FROM K1Config"
                objCommand.CommandType = CommandType.Text
                Using objAdapter As New SqlClient.SqlDataAdapter() With {.SelectCommand = objCommand}
                    objDT = New DataTable("DT")
                    objAdapter.Fill(objDT)
                End Using

                If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                    m_strConfigServer = CStr(objDT.Rows(0)("OwnerServer"))
                    m_strConfigDatabase = CStr(objDT.Rows(0)("OwnerDB"))

                    Dim strServer As String = String.Empty
                    Dim strDatabase As String = String.Empty
                    Dim strUser As String = String.Empty
                    m_objDB.GetDatabaseInfo(strServer, strDatabase, strUser)

                    If Not strServer.Equals(String.Empty) AndAlso Not strDatabase.Equals(String.Empty) Then
                        If m_strConfigServer.ToUpper = strServer.ToUpper AndAlso m_strConfigDatabase.ToUpper = strDatabase.ToUpper Then
                            m_blnValid = True
                        End If
                    End If
                End If

                objCommand = New SqlClient.SqlCommand("SELECT COUNT(ID) FROM EDOC", objConnection)
                m_intRows = CInt(objCommand.ExecuteScalar())

                objDT = Nothing
                objCommand = objConnection.CreateCommand()
                objCommand.CommandText = "SELECT CAST(FILEPROPERTY(name, 'SpaceUsed')AS int)/128.0 AS SpaceUsedMB from dbo.sysfiles where Name = '" & m_strDatabase & "'"
                objCommand.CommandType = CommandType.Text
                Using objAdapter As New SqlClient.SqlDataAdapter() With {.SelectCommand = objCommand}
                    objDT = New DataTable("DT")
                    objAdapter.Fill(objDT)
                End Using

                If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                    m_strSize = ConvertSizeToText(CStr(objDT.Rows(0)(0)), enumFileSizeUnit.cMB)
                End If

                objConnection.Close()
            Catch ex As SqlClient.SqlException
                'do nothing
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

#End Region

#Region " Properties "



    Public Overloads ReadOnly Property ID() As Integer
        Get
            Return m_intID
        End Get
    End Property

    Public ReadOnly Property ConnectionString() As String
        Get
            Return m_strConnection
        End Get
    End Property

    Public ReadOnly Property Server() As String
        Get
            Return m_strServer
        End Get
    End Property

    Public ReadOnly Property DB() As String
        Get
            Return m_strDatabase
        End Get
    End Property

    Public Property ConfigServer() As String
        Get
            Return m_strConfigServer
        End Get
        Set(value As String)
            m_strConfigServer = value
        End Set
    End Property

    Public Property ConfigDB() As String
        Get
            Return m_strConfigDatabase
        End Get
        Set(value As String)
            m_strConfigDatabase = value
        End Set
    End Property

    Public Overloads ReadOnly Property Rows() As Integer
        Get
            Return m_intRows
        End Get
    End Property

    Public ReadOnly Property Size() As String
        Get
            Return m_strSize
        End Get
    End Property

    Public ReadOnly Property Current() As Boolean
        Get
            Return m_blnCurrent
        End Get
    End Property

    Public ReadOnly Property IsValid() As Boolean
        Get
            Return m_blnValid
        End Get
    End Property
#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsEDOCArchive
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cK1ARCHIVE, intID)

            Return New clsEDOCArchive(objDB, objDT.Rows(0))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsEDOCArchive)
        Dim colObjects As FrameworkCollections.K1Dictionary(Of clsEDOCArchive)

        Try
            Dim objTable As clsTable = objDB.SysInfo.Tables(clsDBConstants.Tables.cK1ARCHIVE)
            Dim objDT As DataTable = objDB.GetDataTable(objTable)

            colObjects = New FrameworkCollections.K1Dictionary(Of clsEDOCArchive)
            For Each objDR As DataRow In objDT.Rows
                Dim objItem As New clsEDOCArchive(objDB, objDR)
                colObjects.Add(CStr(objItem.ID), objItem)
            Next

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
