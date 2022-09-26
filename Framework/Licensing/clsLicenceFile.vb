Imports K1Library.clsDBConstants
Imports System.Xml.Linq

Namespace Licensing

    Public Class clsLicenseFile

#Region " Members "

        Private m_objDB As clsDB
        Private m_intID As Integer = cintNULL
        Private m_eApplicationType As enumApplicationType
        Private m_blnExist As Boolean

#End Region

#Region " Constructors "

        Public Sub New(ByVal objDB As clsDB, ByVal eApplicationType As enumApplicationType)
            m_objDB = objDB
            m_eApplicationType = eApplicationType
            m_intID = GetLicenseFileID(objDB, m_eApplicationType)
            m_blnExist = (m_intID > 0)
        End Sub

        Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)
            m_objDB = objDB
            m_intID = CInt(clsDB_Direct.DataRowValue(objDR, Fields.cID, cintNULL))
            m_eApplicationType = CType(clsDB_Direct.DataRowValue(objDR,
                                                                 Fields.K1LicenceFile.cAPPLICATION_TYPE,
                                                                 cintNULL),
                                       enumApplicationType)
            m_blnExist = True
        End Sub

#End Region

        Public ReadOnly Property ID() As Integer
            Get
                Return m_intID
            End Get
        End Property

        Public ReadOnly Property Exists() As Boolean
            Get
                Return m_blnExist
            End Get
        End Property

        Public Sub InsertUpdate(ByVal strFile As String)
            If m_intID = clsDBConstants.cintNULL Then
                m_intID = GetLicenseFileID(m_objDB, m_eApplicationType)
            End If


            If m_intID = 0 Then
                Insert(strFile)
            Else
                Update(strFile)
            End If
        End Sub

        Private Sub Insert(ByVal strFile As String)
            Dim blnCreateTransaction As Boolean = False
            Try
                If Not m_objDB.HasTransaction Then
                    m_objDB.BeginTransaction()
                    blnCreateTransaction = True
                End If


                Dim strInsertSQL As String = "INSERT INTO [{0}] ([{1}], [{2}]) VALUES (@AppType, @File)"

                strInsertSQL = String.Format(strInsertSQL, Tables.cK1LICENCEFILE,
                                             Fields.K1LicenceFile.cAPPLICATION_TYPE,
                                             Fields.K1LicenceFile.CLICENCEFILE)

                Dim colParams As New clsDBParameterDictionary
                colParams.Add(New clsDBParameter("@AppType", m_eApplicationType, ParameterDirection.Input, SqlDbType.Int))
                colParams.Add(New clsDBParameter("@File", New Byte() {0}, ParameterDirection.Input, SqlDbType.Image))

                m_objDB.ExecuteSQL(strInsertSQL, colParams)

                m_intID = GetLicenseFileID(m_objDB, m_eApplicationType)

                Update(strFile)

                If blnCreateTransaction Then m_objDB.EndTransaction(True)
            Catch ex As Exception
                If blnCreateTransaction Then m_objDB.EndTransaction(False)
                Throw
            End Try
        End Sub


        Private Sub Update(ByVal strFile As String)
            Try
                m_objDB.WriteBLOB(clsDBConstants.Tables.cK1LICENCEFILE, Fields.K1LicenceFile.CLICENCEFILE,
                                 SqlDbType.Image, cintNULL, m_intID, strFile)
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Private Shared Function GetLicenseFileID(ByVal objDB As clsDB, ByVal eApplicationType As enumApplicationType) As Integer
            Try
                Dim strSelectSQL As String = "SELECT [{0}] FROM [{1}] WHERE [{2}] = {3}"

                strSelectSQL = String.Format(strSelectSQL, Fields.cID,
                                             Tables.cK1LICENCEFILE,
                                             Fields.K1LicenceFile.cAPPLICATION_TYPE,
                                             CInt(eApplicationType))

                Return objDB.ExecuteScalar(strSelectSQL)
            Catch ex As Exception
                Throw
            End Try
        End Function

        Friend Function DecryptLicenseFile() As String
            Try
                m_blnExist = True
                Return DecryptLicenseFile(m_objDB, m_eApplicationType)
            Catch ex As clsK1Exception When ex.Message.Contains("License file not found.")
                m_blnExist = False
                Return Nothing
            Catch ex As Exception
                Throw
            End Try
        End Function

        Friend Shared Function DecryptLicenseFile(ByVal strFilePath As String) As String
            Dim objStream As FileStream = Nothing
            Try
                '-- Get bytes from file
                objStream = New FileStream(strFilePath, FileMode.Open, FileAccess.Read)
                Dim arrBytes(CInt(objStream.Length - 1)) As Byte
                objStream.Read(arrBytes, 0, CInt(objStream.Length))

                Return DecryptLicenseFile(arrBytes)
            Catch ex As Exception
                Throw
            Finally
                If objStream IsNot Nothing Then
                    objStream.Close()
                End If
            End Try
        End Function

        Private Shared Function DecryptLicenseFile(ByVal objDB As clsDB,
                                                  ByVal eApplicationType As enumApplicationType) As String
            Try
                Dim intID As Integer = clsLicenseFile.GetLicenseFileID(objDB, eApplicationType)

                If intID < 1 Then
                    Throw New clsK1Exception("License file not found.", True)
                End If

                Dim arrBytes() As Byte = objDB.ReadBLOBToMemory(Tables.cK1LICENCEFILE,
                                                                    Fields.K1LicenceFile.CLICENCEFILE, intID)
                Return DecryptLicenseFile(arrBytes)
            Catch ex As Exception
                Throw
            End Try
        End Function

        Private Shared Function DecryptLicenseFile(arrBytes() As Byte) As String
            Try
                Dim strEncrypted As String = Text.Encoding.UTF7.GetString(arrBytes)
                Dim objEncryption As New clsEncryption(True)
                Dim strPlainText As String = objEncryption.Decrypt(strEncrypted)

                Return strPlainText
            Catch ex As Exception
                Throw
            End Try
        End Function

        Friend Function LoadXMLFromLicenseFile() As XDocument
            Return LoadXMLFromLicenseFile(m_objDB, m_eApplicationType)
        End Function

        Friend Shared Function LoadXMLFromLicenseFile(ByVal objDB As clsDB,
                                                      ByVal eApplicationType As enumApplicationType) As XDocument
            Try
                Dim strPlainText As String = DecryptLicenseFile(objDB, eApplicationType)

                Return XDocument.Parse(strPlainText)
            Catch ex As FormatException
                Throw New clsK1Exception("Could not load the " & modGlobal.GetProductName(eApplicationType) & " license file. License file is corrupted.", True)
            Catch ex As Exception
                Throw
            End Try
        End Function

        Friend Sub Remove(ByVal objDB As clsDB)
            Try
                Remove(objDB, m_eApplicationType)
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Friend Shared Sub Remove(ByVal objDB As clsDB, ByVal eAppType As clsDBConstants.enumApplicationType)
            Try
                Dim strSelectSQL As String = "DELETE FROM [{0}] WHERE [{1}] = {2}"

                strSelectSQL = String.Format(strSelectSQL, Tables.cK1LICENCEFILE,
                                             Fields.K1LicenceFile.cAPPLICATION_TYPE,
                                             CInt(eAppType))

                objDB.ExecuteSQL(strSelectSQL)
            Catch ex As Exception
                Throw
            End Try
        End Sub
    End Class

End Namespace