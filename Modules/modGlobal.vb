Imports System.Configuration
Imports System.Runtime.CompilerServices
Imports Microsoft.Win32
Imports System.Text
Imports System.Drawing
Imports System.Xml
Imports System.Linq
Imports System.Threading
Imports System.IO.Compression

Public Module modGlobal

#Region " Public Enumerations "

    Public Enum enumFileSizeUnit As Integer
        cGB = 1073741824
        cMB = 1048576
        cKB = 1024
        cB = 1
    End Enum
#End Region

#Region " IO Functions "

    ''' <summary>
    ''' will add a "\" character to a path if it is not already the last character
    ''' </summary>
    ''' <param name="strPath"></param>
    ''' <returns></returns>
    ''' <remarks>Stupid! Why not just use System.IO.Path.CombinePath(...) !!!</remarks>
    Public Function ProperPath(ByVal strPath As String) As String
        strPath = strPath.Trim

        If Not strPath Is Nothing AndAlso strPath.Length > 0 Then
            If Not strPath.Substring(strPath.Length - 1, 1) = IO.Path.DirectorySeparatorChar Then
                strPath &= IO.Path.DirectorySeparatorChar
            End If
        End If

        Return strPath
    End Function

    Public Sub CreateDir(ByVal strPath As String)
        If Not System.IO.Directory.Exists(strPath) Then
            System.IO.Directory.CreateDirectory(strPath)
        End If
    End Sub

    Public Sub WriteToFile(ByVal strFile As String, ByVal strText As String)
        CreateDir(Path.GetDirectoryName(strFile))

        Dim objSW As New StreamWriter(strFile, False)

        objSW.WriteLine(strText)

        objSW.Close()
    End Sub

    Public Function GetMIMEType(ByVal strFileName As String) As String
        Dim strContentType As String = "application/octet-stream"

        Dim objRK As RegistryKey = Registry.ClassesRoot.OpenSubKey(Path.GetExtension(strFileName))

        If Not objRK Is Nothing Then
            strContentType = CType(objRK.GetValue("Content Type", "application/octet-stream"), String)
        End If

        Return strContentType
    End Function

    Public Function OpenFile(ByVal strFile As String) As Boolean
        Try
            System.Diagnostics.Process.Start(strFile)
            Return True
        Catch ex As Exception

        End Try

        Try
            Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & strFile, vbNormalFocus)
            Return True
        Catch ex As Exception

        End Try

        Try
            System.Diagnostics.Process.Start(IO.Path.GetDirectoryName(strFile))
            Return True
        Catch ex As Exception

        End Try

        Return False
    End Function

    ''' <summary>
    ''' This is a misleading method. This method breaks SOLID principal of single responsibility.
    ''' </summary>
    ''' <param name="strFile"></param>
    ''' <returns></returns>
    ''' <remarks>Stupid method!!!</remarks>
    Public Function ValidateFileName(ByVal strFile As String) As String
        If strFile IsNot Nothing Then
            Dim arrChars() As Char = IO.Path.GetInvalidFileNameChars

            Dim arrChr = From c In strFile
                         Where Not arrChars.Contains(c)
                         Select c

            Dim strNewFile As String = New String(arrChr.ToArray)

            If strNewFile.Length = 0 Then
                strNewFile = "temp"
            Else
                If strNewFile.IndexOf("."c) = 0 Then
                    strNewFile = "temp" & strNewFile
                End If
            End If

            Return strNewFile
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region " String Functions "

    ''' <summary>
    ''' Returns a section of a string given begin and end index strings.
    ''' </summary>
    ''' <param name="strToParse">The string to search in</param>
    ''' <param name="strBegin">Where to start when returning sub-string.</param>
    ''' <param name="strEnd">Where to end when returning sub-string.  
    ''' If the End string is not found, the sub-string from the begin string to the 
    ''' end of the initial string is returned.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ParseString(ByVal strToParse As String, ByVal strBegin As String,
    ByVal strEnd As String, Optional ByVal blnCaseSensitive As Boolean = False) As String
        Dim strReturn As String = ""
        Dim strTest As String

        Try
            If blnCaseSensitive Then
                strTest = strToParse
            Else
                strTest = strToParse.ToUpper()
            End If

            Dim intStartIndex As Integer = strTest.IndexOf(strBegin.ToUpper)
            Dim intEndIndex As Integer = strTest.IndexOf(strEnd.ToUpper, intStartIndex)

            If Not intEndIndex >= 0 Then
                intEndIndex = strToParse.Length
            End If

            If intStartIndex >= 0 Then
                strReturn = strToParse.Substring(intStartIndex + strBegin.Length,
                    intEndIndex - (intStartIndex + strBegin.Length))
            End If
        Catch ex As Exception
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' Returns a comma separated string of the IDs (integers) in the collection
    ''' </summary>
    Public Function CreateHashtableFromIDString(ByVal strIDs As String) As Hashtable
        Dim colIDs As New Hashtable
        Dim arrIds As String() = strIDs.Split(","c)

        For intLoop As Integer = 0 To arrIds.Length - 1
            Dim intId As Integer = CType(arrIds(intLoop), Integer)
            If colIDs(CStr(intId)) Is Nothing Then
                colIDs.Add(CStr(intId), intId)
            End If
        Next

        Return colIDs
    End Function

    ''' <summary>
    ''' Returns a comma separated string of the IDs (integers) in the collection
    ''' </summary>
    Public Function CreateIDStringFromCollection(ByVal colIDs As ICollection) As String
        Dim objSB As New StringBuilder
        For Each objID As Object In colIDs
            AppendToCommaString(objSB, CStr(objID))
        Next
        Return objSB.ToString
    End Function

    ''' <summary>
    ''' Returns a comma separated string of the IDs (integers) in the collection
    ''' </summary>
    ''' <param name="objDT"></param>
    ''' <returns>a coma delimited string.</returns>
    ''' <remarks>lets avoid this method, pls don't use with a large array. will crash the application</remarks>    
    Public Function CreateIDStringFromDataTable(ByVal objDT As DataTable) As String
        Dim objSB As New StringBuilder
        For Each objRow As DataRow In objDT.Rows
            AppendToCommaString(objSB, CStr(objRow(0)))
        Next
        Return objSB.ToString
    End Function

    ''' <summary>
    ''' Returns a comma separated string of the IDs (integers) in the collection
    ''' </summary>
    Public Function CreateDataTableFromIDString(ByVal strIDs As String) As DataTable
        Dim objDT As New DataTable("IDTable")
        objDT.Columns.Add("ID", GetType(System.Int64))

        Dim arrIds As String() = strIDs.Split(","c)

        For intLoop As Integer = 0 To arrIds.Length - 1
            Dim intID As Integer = CType(arrIds(intLoop), Integer)
            objDT.Rows.Add(New Object() {intID})
        Next

        Return objDT
    End Function

    ''' <summary>
    ''' Will create a comma delimited string using the original string and the new value
    ''' </summary>
    Public Sub AppendToCommaString(ByRef strString As String, ByVal strValue As String)
        If strString Is Nothing OrElse strString.Length = 0 Then
            strString = strValue
        Else
            strString &= ", " & strValue
        End If
    End Sub

    ''' <summary>
    ''' Will create a comma delimited string using the original string and the new value
    ''' </summary>
    Public Sub AppendToCommaString(ByRef objSB As StringBuilder, ByVal strValue As String)
        If objSB.Length = 0 Then
            objSB.Append(strValue)
        Else
            objSB.Append(", ")
            objSB.Append(strValue)
        End If
    End Sub

    ''' <summary>
    ''' Will create a delimited string using the original string, the new value, and delimeted by strSeparator
    ''' </summary>
    Public Sub AppendToString(ByRef strString As String, ByVal strValue As String, ByVal strSeparator As String)
        If strString Is Nothing OrElse strString.Length = 0 Then
            strString = strValue
        Else
            strString &= strSeparator & strValue
        End If
    End Sub

    ''' <summary>
    ''' Will create a comma delimited string using the original string and the new value
    ''' </summary>
    Public Sub AppendToString(ByRef objSB As StringBuilder, ByVal strValue As String, ByVal strSeparator As String)
        If objSB.Length = 0 Then
            objSB.Append(strValue)
        Else
            objSB.Append(strSeparator)
            objSB.Append(strValue)
        End If
    End Sub

    ''' <summary>
    ''' Will create a delimited string using the original string, the new value, and delimeted by strSeparator
    ''' </summary>
    Public Sub AppendToString(ByRef strString As String, ByVal strValue As String,
    ByVal strIfEmptySeparator As String, ByVal strIfNotEmptySeparator As String)
        'TODO: Use StringBuilder class?
        If strString Is Nothing OrElse strString.Length = 0 Then
            strString = strIfEmptySeparator & strValue
        Else
            strString &= strIfNotEmptySeparator & strValue
        End If
    End Sub

    ''' <summary>
    ''' Takes a noun and returns the plural form of it
    ''' </summary>
    Public Function MakePlural(ByVal strNoun As String) As String
        Dim strWord As String = strNoun

        If strWord Is Nothing OrElse strWord.Length = 0 Then
            Return strWord
        End If

        If strWord.Substring(strWord.Length - 1, 1).ToUpper = "X" OrElse
        (strWord.Length >= 2 AndAlso
        (strWord.Substring(strWord.Length - 2, 2).ToUpper = "SS" _
        OrElse strWord.Substring(strWord.Length - 2, 2).ToUpper = "SH" _
        OrElse strWord.Substring(strWord.Length - 2, 2).ToUpper = "CH")) Then
            Return strWord & "es"
        End If

        If strWord.Substring(strWord.Length - 1, 1).ToUpper = "S" Then
            Return strWord
        End If

        If strWord.Length >= 2 AndAlso
        strWord.Substring(strWord.Length - 1, 1).ToUpper = "Y" AndAlso
        Not (strWord.Substring(strWord.Length - 2, 1).ToUpper = "A" OrElse
        strWord.Substring(strWord.Length - 2, 1).ToUpper = "O" OrElse
        strWord.Substring(strWord.Length - 2, 1).ToUpper = "U" OrElse
        strWord.Substring(strWord.Length - 2, 1).ToUpper = "I" OrElse
        strWord.Substring(strWord.Length - 2, 1).ToUpper = "E") Then
            Return strWord.Substring(0, strWord.Length - 1) + "ies"
        End If

        Return strWord & "s"
    End Function

    Public Function BooleanToString(ByVal blnValue As Boolean) As String
        Return BooleanToString(blnValue, "Yes", "No")
    End Function

    Public Function BooleanToString(ByVal blnValue As Boolean, ByVal strTrueText As String, ByVal strFalseText As String) As String
        If blnValue Then
            Return strTrueText
        Else
            Return strFalseText
        End If
    End Function

    ''' <summary>
    ''' Extract the hostname out from the FQDN and return the host\instance
    ''' </summary>
    ''' <param name="strFQDN"></param>
    ''' <returns>host\instance</returns>
    Function ConvertFQDNToHostAndInst(strFQDN As String) As String
        ''remove the port number
        If strFQDN.Contains(",") Then
            strFQDN = strFQDN.Substring(0, strFQDN.IndexOf(","))
        End If

        '' if not FQDN, thats all we need
        If Not strFQDN.Contains(".") Then
            Return strFQDN
        End If

        'get the host name
        Dim hostName = strFQDN.Substring(0, strFQDN.IndexOf("."))
        Dim instanceName As String = String.Empty

        'get the instance name
        If strFQDN.Contains("\") Then
            instanceName = strFQDN.Substring(strFQDN.IndexOf("\") + 1)
            Return $"{hostName}\{instanceName}"
        End If
        Return hostName

    End Function
#End Region

#Region " Get Single Value "

    <Extension()>
    Public Function GetSingleValue(ByVal colIDS As Hashtable) As Integer
        If colIDS IsNot Nothing AndAlso colIDS.Count = 1 Then
            For Each intID As Integer In colIDS.Values
                Return intID
            Next
        End If

        Return clsDBConstants.cintNULL
    End Function

#End Region

#Region " Convert IDs To Data Table "

    Public Function ConvertIDsToDataTable(ByVal colIDs As Hashtable) As DataTable
        Return ConvertIDsToDataTable(colIDs.Values)
    End Function

    Public Function ConvertIDsToDataTable(ByVal colIDs As ICollection) As DataTable
        Dim objDT As New DataTable("IDTable")
        objDT.Columns.Add("ID", GetType(System.Int64))

        For Each intID As Integer In colIDs
            objDT.Rows.Add(New Object() {intID})
        Next

        Return objDT
    End Function
#End Region

#Region " Convert Data Table To IDs "

    Public Sub ConvertDataTableToIDs(ByVal objDT As DataTable,
                                     ByRef colIDs As Hashtable,
                                     ByVal strField As String)

        If colIDs Is Nothing Then
            colIDs = New Hashtable
        End If

        For Each objRow As DataRow In objDT.Rows
            Dim intID As Integer = CInt(objRow(strField))

            If colIDs(CStr(intID)) Is Nothing Then
                colIDs.Add(CStr(intID), intID)
            End If
        Next
    End Sub

#End Region

#Region " Reverse a datatable "

    Public Sub ReverseDataTable(ByRef objDT As DataTable)
        Dim intEndIndex As Integer = objDT.Rows.Count - 1

        For intLoop As Integer = 0 To objDT.Rows.Count - 1
            If intEndIndex > intLoop Then
                Dim objRow As DataRow = objDT.Rows(intLoop)
                Dim objRow2 As DataRow = objDT.Rows(intEndIndex)

                Dim objAddRow1 As DataRow = objDT.NewRow
                objAddRow1.ItemArray = objRow.ItemArray
                Dim objAddRow2 As DataRow = objDT.NewRow
                objAddRow2.ItemArray = objRow2.ItemArray

                objDT.Rows.Remove(objRow)
                objDT.Rows.Remove(objRow2)

                objDT.Rows.InsertAt(objAddRow2, intLoop)
                objDT.Rows.InsertAt(objAddRow1, intEndIndex)
            Else
                Exit For
            End If

            intEndIndex -= 1
        Next
    End Sub
#End Region

#Region " Convert Size To Text "

    Public Function ConvertSizeToText(ByVal strFileSize As String,
    Optional ByVal eFileSizeUnit As enumFileSizeUnit = enumFileSizeUnit.cB) As String
        Dim intFileSize As Decimal

        If strFileSize Is Nothing OrElse strFileSize.Trim.Length = 0 Then
            Return "Unknown"
        Else
            Try
                intFileSize = CType(strFileSize, Decimal)

                intFileSize *= eFileSizeUnit '-- convert to bytes

                'Const cGB As Integer = 1073741824
                'Const cMB As Integer = 1048576
                'Const cKB As Integer = 1024

                If intFileSize >= enumFileSizeUnit.cGB Then 'gigabytes
                    Return Math.Round(intFileSize / enumFileSizeUnit.cGB, 2) & " GB"
                End If

                If intFileSize >= enumFileSizeUnit.cMB Then  'megabytes
                    Return Math.Round(intFileSize / enumFileSizeUnit.cMB, 2) & " MB"
                End If

                If intFileSize >= enumFileSizeUnit.cKB Then 'kilobytes
                    Return Math.Round(intFileSize / enumFileSizeUnit.cKB, 2) & " KB"
                End If

                Return intFileSize & " Byte(s)"
            Catch ex As Exception
                Return "Unknown"
            End Try
        End If
    End Function

#End Region

#Region " Array Functions "

    ''' <summary>
    ''' Implodes an array into a string joined by strGlue
    ''' </summary>
    ''' <param name="arrValues">Array of values we want to convert into a string</param>
    ''' <param name="strGlue">Separates the values</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ImplodeArray(ByVal arrValues() As String, ByVal strGlue As String) As String
        Try
            Dim objStringBuilder As New StringBuilder
            Dim intCount As Integer = arrValues.GetUpperBound(0)
            For intIndex As Integer = 0 To intCount
                If objStringBuilder.Length > 0 Then
                    objStringBuilder.Append(strGlue)
                End If
                objStringBuilder.Append(arrValues(intIndex))
            Next

            Return objStringBuilder.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

    <Extension()>
    Public Function ToDelimitedString(Of T)(ByVal col As IEnumerable(Of T), ByVal delimiter As Char) As String
        Dim strBuilder As New StringBuilder()
        For Each val As T In col
            strBuilder.Append(val.ToString())
            strBuilder.Append(delimiter)
        Next

        If (strBuilder.ToString() = String.Empty) Then
            Return String.Empty
        End If

        Return strBuilder.Replace(delimiter, String.Empty, strBuilder.Length - 1, 1).ToString()
    End Function
#End Region

#Region " Isolated Storage Functions "

#Region " Get Isolated Storage File "

    ''' <summary>
    ''' Obtains user-scoped or machine-scoped isolated storage corresponding to the application domain 
    ''' identity and assembly or calling code's application identity or assembly identity.
    ''' </summary>
    ''' <param name="eIsolatedStorage">Isolated Storage scope we want to use.</param>
    ''' <exception cref="System.Security.SecurityException">Sufficient isolated storage permissions have not been granted.</exception>
    ''' <remarks>
    ''' eIsolatedStorage is a bitwise value with the following valid values:
    ''' Machine Scope
    '''   Application = IsolatedStorageScope.Machine or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.Machine or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.Machine or IsolatedStorageScope.Domain
    ''' 
    ''' User Scope (IsolatedStorageScope.User is optional)
    '''   Application = IsolatedStorageScope.User or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.User or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.User or IsolatedStorageScope.Domain 
    ''' </remarks>
    Private Function GetIsolatedStorageFile(ByVal eIsolatedStorage As IO.IsolatedStorage.IsolatedStorageScope) _
    As IO.IsolatedStorage.IsolatedStorageFile
        Try
            Dim blnMachineScope As Boolean = False
            Dim objIsolatedStorageFile As IO.IsolatedStorage.IsolatedStorageFile = Nothing

            Dim eValidScope As IO.IsolatedStorage.IsolatedStorageScope =
                IO.IsolatedStorage.IsolatedStorageScope.Application _
                Or IO.IsolatedStorage.IsolatedStorageScope.Assembly _
                Or IO.IsolatedStorage.IsolatedStorageScope.Domain

            If CInt(IO.IsolatedStorage.IsolatedStorageScope.Machine And eIsolatedStorage) > 0 Then
                '-- We want to use the Isolated Storage at Machine Scope rather then User
                blnMachineScope = True
            End If

            '-- Remove invalid flags from the bitwise value
            eIsolatedStorage = eIsolatedStorage And eValidScope

            Select Case eIsolatedStorage
                Case IO.IsolatedStorage.IsolatedStorageScope.Assembly
                    If blnMachineScope Then
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetMachineStoreForAssembly
                    Else
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetUserStoreForAssembly
                    End If

                Case IO.IsolatedStorage.IsolatedStorageScope.Domain
                    If blnMachineScope Then
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetMachineStoreForDomain
                    Else
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetUserStoreForDomain
                    End If

                Case Else
                    If blnMachineScope Then
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetMachineStoreForApplication
                    Else
                        objIsolatedStorageFile = IO.IsolatedStorage.IsolatedStorageFile.GetUserStoreForApplication
                    End If
            End Select

            Return objIsolatedStorageFile
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Get Isolated Storage File Stream "

    ''' <summary>
    ''' Initializes a new instance of an System.IO.IsolatedStorage.IsolatedStorageFileStream object 
    ''' giving access to the file designated by path in the specified mode.
    ''' </summary>
    ''' <param name="eIsolatedStorage">Isolated Storage scope we want to use.</param>
    ''' <param name="strFileName">The relative path of the file within isolated storage.</param>
    ''' <param name="eFileMode">One of the System.IO.FileMode values.</param>
    ''' <param name="eFileAccess">A bitwise combination of the System.IO.FileAccess values.</param>
    ''' <exception cref="System.IO.FileNotFoundException">No file was found and the mode is set to Open.</exception>
    ''' <exception cref="System.ArgumentException">The path is badly formed.</exception>
    ''' <exception cref="System.ArgumentNullException">The path is null.</exception>
    Public Function GetIsolatedStorageFileStream(ByVal eIsolatedStorage As IO.IsolatedStorage.IsolatedStorageScope,
            ByVal strFileName As String, ByVal eFileMode As IO.FileMode, ByVal eFileAccess As IO.FileAccess) As IO.IsolatedStorage.IsolatedStorageFileStream
        Try
            Dim objIsolatedStorageFile As IO.IsolatedStorage.IsolatedStorageFile =
                GetIsolatedStorageFile(eIsolatedStorage)

            Dim objFileStream As New IO.IsolatedStorage.IsolatedStorageFileStream(strFileName, eFileMode,
                eFileAccess, objIsolatedStorageFile)

            Return objFileStream
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Xml Serialize And Save "

    ''' <summary>
    ''' Serializes an object into XML and then saves it into the users application scope
    ''' </summary>
    ''' <param name="strFileName">The relative path of the file within isolated storage.</param>
    ''' <param name="objValue">Object to be serialized and saved.</param>
    ''' <remarks>
    ''' Root isolated storage folder = [Documents and Settings]\[User]\Local Settings\Application Data\IsolatedStorage\
    ''' </remarks> 
    ''' <example>
    ''' Dim objAutoComplete As AutoCompleteStringCollection = txtUserName.AutoCompleteCustomSource
    ''' XmlSerializeAndSave("user_auto.xml", objAutoComplete)
    ''' </example>
    Public Sub XmlSerializeAndSave(ByVal strFileName As String, ByVal objValue As Object)
        Try
            XmlSerializeAndSave(IsolatedStorage.IsolatedStorageScope.Domain, strFileName, objValue)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Serializes an object into XML and then saves it into the specified isolated storage scope
    ''' </summary>
    ''' <param name="eIsolatedStorage">Isolated Storage scope we want to use.</param>
    ''' <param name="strFileName">The relative path of the file within isolated storage.</param>
    ''' <param name="objValue">Object to be serialized and saved.</param>
    ''' <remarks>
    ''' Root User isolated storage folder = [Documents and Settings]\[User]\Local Settings\Application Data\IsolatedStorage\
    ''' Root Machine isolated storage folder = [Documents and Settings]\All Users\Application Data\IsolatedStorage\
    ''' 
    ''' eIsolatedStorage is a bitwise value with the following valid values:
    ''' Machine Scope
    '''   Application = IsolatedStorageScope.Machine or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.Machine or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.Machine or IsolatedStorageScope.Domain
    ''' 
    ''' User Scope (IsolatedStorageScope.User is optional)
    '''   Application = IsolatedStorageScope.User or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.User or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.User or IsolatedStorageScope.Domain 
    ''' </remarks> 
    ''' <example>
    ''' Dim eIsolatedScope as IsolatedStorageScope = IsolatedStorageScope.Machine or IsolatedStorageScope.Application
    ''' Dim objAutoComplete As AutoCompleteStringCollection = txtUserName.AutoCompleteCustomSource
    ''' XmlSerializeAndSave(eIsolatedScope, "user_auto.xml", objAutoComplete)
    ''' </example>
    Public Sub XmlSerializeAndSave(ByVal eIsolatedStorage As IO.IsolatedStorage.IsolatedStorageScope,
    ByVal strFileName As String, ByVal objValue As Object)
        Dim objFileStream As IO.IsolatedStorage.IsolatedStorageFileStream = Nothing

        Try
            objFileStream = GetIsolatedStorageFileStream(eIsolatedStorage, strFileName, FileMode.OpenOrCreate,
                FileAccess.Write)

            Dim objStreamWriter As IO.StreamWriter = New IO.StreamWriter(objFileStream)

            Dim objXmlWriter As New System.Xml.Serialization.XmlSerializer(objValue.GetType)

            objXmlWriter.Serialize(objStreamWriter, objValue)
        Catch ex As Exception
            Throw
        Finally
            If objFileStream IsNot Nothing Then
                objFileStream.Close()
                objFileStream = Nothing
            End If
        End Try
    End Sub

#End Region

#Region " Xml Deserialize And Load "

    ''' <summary>
    ''' Loads a file from the users application scope and deserializes it
    ''' </summary>
    ''' <param name="strFileName">The relative path of the file within isolated storage.</param>
    ''' <param name="objType">Type of the object that will be deserialized.</param>
    ''' <example>
    ''' Dim objAutoComplete As AutoCompleteStringCollection = XmlDeserializeAndLoad("user_auto.xml", GetType(AutoCompleteStringCollection))
    ''' </example>
    Public Function XmlDeserializeAndLoad(ByVal strFileName As String, ByVal objType As Type) As Object
        Try
            Return XmlDeserializeAndLoad(IsolatedStorage.IsolatedStorageScope.Domain, strFileName, objType)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Loads a file from the specified isolated storage scope and deserializes it
    ''' </summary>
    ''' <param name="eIsolatedStorage">Isolated Storage scope we want to use.</param>
    ''' <param name="strFileName">The relative path of the file within isolated storage.</param>
    ''' <param name="objType">Type of the object that will be deserialized.</param>
    ''' <remarks>
    ''' eIsolatedStorage is a bitwise value with the following valid values:
    ''' Machine Scope
    '''   Application = IsolatedStorageScope.Machine or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.Machine or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.Machine or IsolatedStorageScope.Domain
    ''' 
    ''' User Scope (IsolatedStorageScope.User is optional)
    '''   Application = IsolatedStorageScope.User or IsolatedStorageScope.Application
    '''   Assembly = IsolatedStorageScope.User or IsolatedStorageScope.Assembly
    '''   Domain = IsolatedStorageScope.User or IsolatedStorageScope.Domain 
    ''' </remarks> 
    ''' <example>
    ''' Dim eIsolatedScope as IsolatedStorageScope = IsolatedStorageScope.Machine or IsolatedStorageScope.Application
    ''' Dim objAutoComplete As AutoCompleteStringCollection = XmlDeserializeAndLoad(eIsolatedScope, "user_auto.xml", GetType(AutoCompleteStringCollection))
    ''' </example>
    Public Function XmlDeserializeAndLoad(ByVal eIsolatedStorage As IO.IsolatedStorage.IsolatedStorageScope,
            ByVal strFileName As String, ByVal objType As Type) As Object
        Dim objFileStream As IO.IsolatedStorage.IsolatedStorageFileStream = Nothing

        Try
            objFileStream = GetIsolatedStorageFileStream(eIsolatedStorage, strFileName, FileMode.Open,
                FileAccess.Read)

            Dim objStreamReader As IO.StreamReader = New IO.StreamReader(objFileStream)

            Dim objXmlWriter As New System.Xml.Serialization.XmlSerializer(objType)

            Dim objValue As Object = objXmlWriter.Deserialize(objStreamReader)

            Return objValue
        Catch ex As IO.FileNotFoundException
            Return Nothing
        Catch ex As Exception
            Throw
        Finally
            If objFileStream IsNot Nothing Then
                objFileStream.Close()
                objFileStream = Nothing
            End If
        End Try
    End Function

#End Region

#End Region

#Region " Version Info "

    Public Function GetVersionString() As String
        Return GetVersionString(My.Application.Info.Version)
    End Function

    Public Function GetVersionString(ByVal objVersion As System.Version) As String
        Dim strFormat As String = "{0}.{1:00}"
        Dim strVersion As String = String.Format(strFormat, objVersion.Major, objVersion.Minor)

        If objVersion.Build > 0 Then
            strVersion &= ChrW((objVersion.Build - 1) + AscW("a"))
        End If

        Return strVersion
    End Function

    ''' <summary>
    ''' Returns the Application version 
    ''' </summary>
    ''' <returns>Returns the Application version.</returns>
    Public Function GetVersion() As Double
        Dim intLengthMinor As Integer = 100
        Dim intLenghtBuild As Integer = 100 * intLengthMinor

        Return (My.Application.Info.Version.Minor / intLengthMinor) +
               (My.Application.Info.Version.Build / intLenghtBuild) +
               My.Application.Info.Version.Major
    End Function
#End Region

#Region " Image Functions "

    Public Function GetImageFileExtension(ByVal objImg As Drawing.Image) As String
        Dim strExtension As String = ""

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Jpeg) Then
            strExtension = "jpg"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Gif) Then
            strExtension = "gif"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Bmp) OrElse
        objImg.RawFormat.Equals(Imaging.ImageFormat.MemoryBmp) Then
            strExtension = "bmp"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Png) Then
            strExtension = "png"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Tiff) Then
            strExtension = "tif"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Wmf) Then
            strExtension = "Wmf"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Emf) Then
            strExtension = "Emf"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Exif) Then
            strExtension = "Exif"
        End If

        If objImg.RawFormat.Equals(Imaging.ImageFormat.Icon) Then
            strExtension = "Ico"
        End If

        Return strExtension
    End Function
#End Region

#Region " Update Connection Info "

    Public Sub UpdateConnectionInfo(ByVal strPath As String, ByRef strConn As String)
        Dim objEncryption As New clsEncryption(True)

        Dim strEncrypted As String = objEncryption.Encrypt(strConn)
        Dim objEnc As New System.Text.ASCIIEncoding
        Dim objUTF As New System.Text.UTF7Encoding
        Dim arrBytes() As Byte = objUTF.GetBytes(strEncrypted)
        strEncrypted = objEnc.GetString(arrBytes)

        Dim objConfig As System.Configuration.Configuration = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(strPath)
        Dim objAppsettings As AppSettingsSection = CType(objConfig.GetSection("connectionStrings"), AppSettingsSection)
        If objAppsettings IsNot Nothing Then
            objAppsettings.Settings("RecFind6.My.MySettings.ConnectionString").Value = strEncrypted
            objConfig.Save()
        End If
    End Sub
#End Region

#Region " Licensing "

    Public Function GetProductName(ByVal eApplicationType As clsDBConstants.enumApplicationType) As String
        Try
            Select Case eApplicationType
                Case clsDBConstants.enumApplicationType.Button
                    Return clsDBConstants.Products.cBUTTON

                Case clsDBConstants.enumApplicationType.GEM
                    Return clsDBConstants.Products.cGEM

                Case clsDBConstants.enumApplicationType.K1
                    Return clsDBConstants.Products.cK1

                Case clsDBConstants.enumApplicationType.Mini_API
                    Return clsDBConstants.Products.cMINI_API

                Case clsDBConstants.enumApplicationType.RecCapture
                    Return clsDBConstants.Products.cRECCAPTURE

                Case clsDBConstants.enumApplicationType.RecFind
                    Return clsDBConstants.Products.cRECFIND

                Case clsDBConstants.enumApplicationType.Scan
                    Return clsDBConstants.Products.cSCAN

                Case clsDBConstants.enumApplicationType.Tacit
                    Return clsDBConstants.Products.cTACIT

                Case clsDBConstants.enumApplicationType.WebClient
                    Return clsDBConstants.Products.cWEBCLIENT

                Case clsDBConstants.enumApplicationType.SharePoint
                    Return clsDBConstants.Products.cSHAREPOINT

                Case clsDBConstants.enumApplicationType.API
                    Return clsDBConstants.Products.cAPI

                Case clsDBConstants.enumApplicationType.OneilIntegration
                    Return clsDBConstants.Products.cONEIL

                Case clsDBConstants.enumApplicationType.Archive
                    Return clsDBConstants.Products.cARCHIVE

                Case clsDBConstants.enumApplicationType.RF6Connector
                    Return clsDBConstants.Products.cRF6CONNECTOR

                Case Else
                    Throw New Exception("Unknown product.")
            End Select
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetLicenceFileApplicationType(ByVal strLicenceFile As String) As clsDBConstants.enumApplicationType
        Try
            Select Case strLicenceFile
                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_BUTTON
                    Return clsDBConstants.enumApplicationType.Button

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_GEM
                    Return clsDBConstants.enumApplicationType.GEM

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_K1
                    Return clsDBConstants.enumApplicationType.K1

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_MINIAPI
                    Return clsDBConstants.enumApplicationType.Mini_API

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_CAPTURE
                    Return clsDBConstants.enumApplicationType.RecCapture

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_RECFIND
                    Return clsDBConstants.enumApplicationType.RecFind

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_HSSM
                    Return clsDBConstants.enumApplicationType.Scan

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_TACIT
                    Return clsDBConstants.enumApplicationType.Tacit

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_WEBCLIENT
                    Return clsDBConstants.enumApplicationType.WebClient

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_SHAREPOINT
                    Return clsDBConstants.enumApplicationType.SharePoint

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_API
                    Return clsDBConstants.enumApplicationType.API

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_ONEILINTEGRATION
                    Return clsDBConstants.enumApplicationType.OneilIntegration

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_ARCHIVE
                    Return clsDBConstants.enumApplicationType.Archive

                Case clsDBConstants.SystemFiles.cstrLICENCE_FILE_RF6CONNECTOR
                    Return clsDBConstants.enumApplicationType.RF6Connector

                Case Else
                    Throw New Exception("Unknown product.")
            End Select
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetApplicationType(ByVal strProductName As String) As clsDBConstants.enumApplicationType
        Try
            Select Case strProductName
                Case clsDBConstants.Products.cBUTTON
                    Return clsDBConstants.enumApplicationType.Button

                Case clsDBConstants.Products.cGEM
                    Return clsDBConstants.enumApplicationType.GEM

                Case clsDBConstants.Products.cK1
                    Return clsDBConstants.enumApplicationType.K1

                Case clsDBConstants.Products.cMINI_API
                    Return clsDBConstants.enumApplicationType.Mini_API

                Case clsDBConstants.Products.cRECCAPTURE
                    Return clsDBConstants.enumApplicationType.RecCapture

                Case clsDBConstants.Products.cRECFIND
                    Return clsDBConstants.enumApplicationType.RecFind

                Case clsDBConstants.Products.cSCAN
                    Return clsDBConstants.enumApplicationType.Scan

                Case clsDBConstants.Products.cTACIT
                    Return clsDBConstants.enumApplicationType.Tacit

                Case clsDBConstants.Products.cWEBCLIENT
                    Return clsDBConstants.enumApplicationType.WebClient

                Case clsDBConstants.Products.cSHAREPOINT
                    Return clsDBConstants.enumApplicationType.SharePoint

                Case clsDBConstants.Products.cAPI
                    Return clsDBConstants.enumApplicationType.API

                Case clsDBConstants.Products.cONEIL
                    Return clsDBConstants.enumApplicationType.OneilIntegration

                Case clsDBConstants.Products.cARCHIVE
                    Return clsDBConstants.enumApplicationType.Archive

                Case clsDBConstants.Products.cRF6CONNECTOR
                    Return clsDBConstants.enumApplicationType.RF6Connector

                Case Else
                    Throw New Exception("Unknown product.")
            End Select
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetActivationType(ByVal eProductType As clsDBConstants.enumApplicationType) As clsDBConstants.enumApplicationType
        Try
            Select Case eProductType
                Case clsDBConstants.enumApplicationType.K1
                    Return clsDBConstants.enumApplicationType.K1ActivationKey

                Case clsDBConstants.enumApplicationType.RecFind
                    Return clsDBConstants.enumApplicationType.RecFindActivationKey

                Case Else
                    Throw New Exception("Unknown product.")
            End Select
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetHostName(ByVal strHostOrAddress As String) As String
        Try
            Dim strTempHost As String = strHostOrAddress
            Dim strPreviousHost As String = strHostOrAddress
            Dim objHostEntry As Net.IPHostEntry
            Do
                strPreviousHost = strTempHost
                objHostEntry = Net.Dns.GetHostEntry(strTempHost)
                strTempHost = objHostEntry.HostName
            Loop While Not strTempHost = strPreviousHost

            Return strTempHost
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetSQLServerName(ByVal objDB As clsDB) As String

        Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT serverproperty('ServerName')")

        If objDT Is Nothing OrElse objDT.Rows.Count = 0 Then
            Throw New Exception("Server name unknown")
        End If

        Return CType(objDT.Rows(0)(0), String)

    End Function

#End Region

#Region " Emailing "

    Public Function SendEmail(ByVal strEmail As String, ByVal strSubject As String,
    ByVal strBody As String, ByVal strSender As String, ByVal strSMTPserver As String,
    Optional ByVal blnAsHTML As Boolean = False) As Boolean
        Try
            Dim objMessage As New Net.Mail.MailMessage(strSender, strEmail, strSubject, strBody)
            objMessage.IsBodyHtml = blnAsHTML

            Dim objClient As New Net.Mail.SmtpClient(strSMTPserver)

            objClient.Send(objMessage)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region " MetadataProfile Types "

    Public Function GetTypeIDs(ByVal objDB As clsDB, ByVal eMDPType As clsDBConstants.enumMDPTypeCodes) As Hashtable
        Try
            Dim colParams As New clsDBParameterDictionary
            colParams.Add(New clsDBParameter("@ID", eMDPType))

            Dim objDT As DataTable = objDB.GetDataTableBySQL("SELECT TypeID FROM K1MDPTypes " & "WHERE TypeCode = @ID", colParams)
            colParams.Dispose()

            If objDT Is Nothing Or objDT.Rows.Count = 0 Then
                Return Nothing
            Else
                Dim colIDs As New Hashtable

                ConvertDataTableToIDs(objDT, colIDs, clsDBConstants.Fields.cTYPEID)

                Return colIDs
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetTypeFilter(ByVal objDB As clsDB, ByVal eMDPType As clsDBConstants.enumMDPTypeCodes) As clsSearchFilter
        Try
            Dim colIDs As Hashtable = GetTypeIDs(objDB, eMDPType)

            If colIDs Is Nothing OrElse colIDs.Count = 0 Then
                Return Nothing
            Else
                Dim objSF As clsSearchFilter
                If colIDs.Count = 1 Then
                    objSF = New clsSearchFilter(objDB, clsDBConstants.Tables.cMETADATAPROFILE & "." &
                        clsDBConstants.Fields.cTYPEID, clsSearchFilter.enumComparisonType.EQUAL, GetSingleValue(colIDs))
                Else
                    objSF = New clsSearchFilter(objDB, clsDBConstants.Tables.cMETADATAPROFILE & "." &
                        clsDBConstants.Fields.cTYPEID, clsSearchFilter.enumComparisonType.IN, colIDs)
                End If

                objSF.Name = GetMDPTypeName(eMDPType)

                Return objSF
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetMDPTypeName(ByVal eMDPType As clsDBConstants.enumMDPTypeCodes) As String
        Select Case eMDPType
            Case clsDBConstants.enumMDPTypeCodes.ArchiveBox
                Return "Archive Boxes"
            Case clsDBConstants.enumMDPTypeCodes.DocumentProfile
                Return "Document Profiles"
            Case Else
                Return "File Folders"
        End Select
    End Function

    Public Function GetMDPTypeFromFilter(ByVal objDb As clsDB, ByVal objSF As clsSearchFilter) As Integer
        If objSF Is Nothing Then
            Return clsDBConstants.cintNULL
        End If

        Dim strTableName As String
        If objSF.TableMask IsNot Nothing Then
            strTableName = objSF.TableMask.Table.ExternalID
        Else
            strTableName = objSF.RootTable
        End If

        If objDb.SysInfo.Tables(strTableName).TypeDependent AndAlso strTableName = clsDBConstants.Tables.cMETADATAPROFILE Then
            Dim strRowFilter As String = objSF.GetTypeFilterFromSearchFilter()

            If strRowFilter Is Nothing Then
                Return clsDBConstants.cintNULL
            End If

            strRowFilter = strRowFilter.Replace(clsDBConstants.Fields.cID, "[" & clsDBConstants.Fields.cTYPEID & "]")
            Dim objDT As DataTable = objDb.GetDataTableBySQL("SELECT [{1}] FROM {0} WHERE {2} GROUP BY [{1}]", Nothing,
                                                             clsDBConstants.Tables.cK1MDPTYPES, clsDBConstants.Fields.K1MDPTypes.cTYPECODE,
                                                             strRowFilter)

            If objDT IsNot Nothing AndAlso objDT.Rows.Count = 1 Then
                Return CType(objDT.Rows(0)(clsDBConstants.Fields.K1MDPTypes.cTYPECODE), Integer)
            End If
        End If

        Return clsDBConstants.cintNULL
    End Function
#End Region

#Region "File Compression"
    Public Function CompressFile(objFileToCompress As FileInfo) As String
        If objFileToCompress Is Nothing Then
            Throw New FileNotFoundException()
        End If

        Dim outputFile = objFileToCompress.FullName & ".gz"

        Try
            Using originalFileStream As FileStream = objFileToCompress.OpenRead()
                If (File.GetAttributes(objFileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden And objFileToCompress.Extension <> ".gz" Then
                    Using compressedFileStream As FileStream = File.Create(outputFile)
                        Using compressionStream As New Compression.GZipStream(compressedFileStream, Compression.CompressionMode.Compress)
                            originalFileStream.CopyTo(compressionStream)
                        End Using
                    End Using
                End If
            End Using
        Catch
            Throw
        End Try

        Return outputFile
    End Function

    Public Function DecompressFile(ByVal fileToDecompress As FileInfo) As String
        If fileToDecompress Is Nothing Then
            Throw New FileNotFoundException()
        End If

        Dim currentFileName As String = fileToDecompress.FullName
        Dim newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length)

        Try
            Using originalFileStream As FileStream = fileToDecompress.OpenRead()
                Using decompressedFileStream As FileStream = File.Create(newFileName)
                    Using decompressionStream As Compression.GZipStream = New Compression.GZipStream(originalFileStream, Compression.CompressionMode.Decompress)
                        decompressionStream.CopyTo(decompressedFileStream)
                    End Using
                End Using
            End Using
        Catch
            Throw
        End Try

        Return newFileName
    End Function

    Public Function CompressBytes(arrBytes As Byte()) As Byte()
        Using objCompressedStream = New MemoryStream()
            Using objZipStream = New Compression.GZipStream(objCompressedStream, Compression.CompressionMode.Compress)
                objZipStream.Write(arrBytes, 0, arrBytes.Length)
                objZipStream.Close()
                Return objCompressedStream.ToArray()
            End Using
        End Using
    End Function

    Public Function DecompressBytes(ByVal arrData As Byte()) As Byte()
        Using objCompressedStream = New MemoryStream(arrData)
            Using objZipStream = New Compression.GZipStream(objCompressedStream, Compression.CompressionMode.Decompress)
                Using objResultStream = New MemoryStream()
                    objZipStream.CopyTo(objResultStream)
                    Return objResultStream.ToArray()
                End Using
            End Using
        End Using
    End Function

#End Region

    ''' <summary>
    ''' Verify that a file is less than 2147483648 bytes.
    ''' </summary>
    ''' <param name="strFileName">Name of the File.</param>
    Public Function VerifyEDOCFileSize(ByVal strFileName As String) As Boolean

        Dim objFileInfo As New IO.FileInfo(strFileName)
        Return objFileInfo.Length < 2147483648
    End Function

    ''' <summary>
    ''' Update a setting's value in the application's config file.
    ''' </summary>
    ''' <param name="strName">Name of the setting.</param>
    ''' <param name="strValue">New value for the setting.</param>
    Public Sub UpdateSettingsValue(ByVal strName As String, ByVal strValue As String)
        Dim objConfig As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim strConfigFile As String = objConfig.FilePath
        'Dim objXDoc As System.Xml.Linq.XDocument = Linq.XDocument.Load(strConfigFile)
        'Dim arrElems As Generic.IEnumerable(Of Linq.XElement) = From n In objXDoc...<setting> _
        '                                                        Where n.@name = strName _
        '                                                        Select n

        'If arrElems.Count = 1 Then
        '    arrElems...<value>.Value = strValue
        '    objXDoc.Save(strConfigFile)
        'End If
        UpdateSettingsValue(strConfigFile, strName, strValue)
    End Sub

    ''' <summary>
    ''' Update a setting's value in the specified configuration file.
    ''' </summary>
    ''' <param name="strConfigFile">File path of configuration file to modify.</param>
    ''' <param name="strName">Name of the setting.</param>
    ''' <param name="strValue">New value for the setting.</param>
    Public Sub UpdateSettingsValue(ByVal strConfigFile As String, ByVal strName As String, ByVal strValue As String)
        'Dim objConfig As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim objXDoc As System.Xml.Linq.XDocument = Linq.XDocument.Load(strConfigFile)
        Dim arrElems As Generic.IEnumerable(Of Linq.XElement) = From n In objXDoc...<setting>
                                                                Where n.@name = strName
                                                                Select n

        If arrElems.Count = 1 Then
            arrElems...<value>.Value = strValue
            objXDoc.Save(strConfigFile)
        End If
    End Sub

    Public Function IsLinkedApplicationMethod(ByVal objDb As clsDB, ByVal eAppMethod As clsApplicationMethod.enumAppMethod) As Boolean
        Dim objAppMethod As clsApplicationMethod = objDb.SysInfo.ApplicationMethods(CStr(eAppMethod))

        Return objAppMethod IsNot Nothing AndAlso (objDb.Profile.LinkApplicationMethods(CStr(objAppMethod.ID)) IsNot Nothing)
    End Function

    Public Function DeleteFolder(strPath As String) As IList(Of String)

        Dim unableToDeleteObjects = New List(Of String)()

        Dim dInfo As DirectoryInfo = New DirectoryInfo(strPath)
        If (dInfo.Exists) Then
            For Each dirInfo As DirectoryInfo In dInfo.GetDirectories()
                DeleteFolder(dirInfo.FullName)
            Next
            For Each fInfo As FileInfo In dInfo.GetFiles()
                Try
                    fInfo.IsReadOnly = False
                    File.SetAttributes(fInfo.FullName, FileAttributes.Normal)
                    fInfo.Delete()
                Catch ex As Exception
                    unableToDeleteObjects.Add(fInfo.FullName)
                End Try
            Next

            Try
                dInfo.Delete(True)
            Catch ex As Exception

                Thread.Sleep(100)
                Try
                    dInfo.Delete(True)
                Catch iex As Exception
                    unableToDeleteObjects.Add(dInfo.FullName)
                End Try

            End Try

        End If

        Return unableToDeleteObjects

    End Function



    <Extension()>
    Public Function IsImageEnabledTable(ByVal objTable As clsTable) As Boolean

        Dim blnHasImageField = objTable.Fields.Any(Function(f) f.Value.DatabaseName.ToUpper() = clsDBConstants.Fields.EDOC.cIMAGE.ToUpper())
        Dim blnHasFileNameField = objTable.Fields.Any(Function(f) f.Value.DatabaseName.ToUpper() = clsDBConstants.Fields.EDOC.cFILENAME.ToUpper())

        Return blnHasFileNameField AndAlso blnHasImageField

    End Function
    <Extension()>
    Public Function IsThisMe(ByVal objTable As clsTable, databaseName As String) As Boolean
        Return String.Equals(objTable.DatabaseName, databaseName, StringComparison.CurrentCultureIgnoreCase)
    End Function

    ''' <summary>
    ''' Determine if the extension is email type.  Leading . is optional.  Case insensitive.
    ''' </summary>
    ''' <param name="ext">expects the extension of the filename</param>
    ''' <returns>true if the extension is email type</returns>
    <Extension()>
    Public Function IsEmailExtension(ext As String) As Boolean
        Return {".msg", "msg", ".eml", "eml"}.Contains(ext.ToLower())
    End Function

    <Extension()>
    Public Function HasEnclosedQuotes(str As String) As Boolean
        ' Ara Melkonian - 2100003650
        ' Text search consumes a thread.
        Return If(str?.Count(Function(x)
                                 Return x.Equals(""""c)
                             End Function) Mod 2 = 0, True)
    End Function

    ''' <summary>
    ''' Emmanuel Cardakaris - 2200003756
    ''' For excel documents that have invalid URIs embedded, reading the contents with openXML fails. 
    ''' The way to fix this is to replace the invalid URIs with a dummy one then reimport the file.
    ''' </summary>
    ''' <param name="fs">Filestream of Excel with bad URIs</param>
    ''' <param name="invalidUriHandler">Delegate function that returns a replacement URI</param>
    Public Sub FixInvalidUri(ByVal fs As Stream, ByVal invalidUriHandler As Func(Of String, Uri))
        Dim relNs As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"

        Using za As ZipArchive = New ZipArchive(fs, ZipArchiveMode.Update)

            For Each entry In za.Entries.ToList()
                If Not entry.Name.EndsWith(".rels") Then Continue For
                Dim replaceEntry As Boolean = False
                Dim entryXDoc As XDocument = Nothing

                Using entryStream = entry.Open()

                    Try
                        entryXDoc = XDocument.Load(entryStream)

                        If entryXDoc.Root IsNot Nothing AndAlso entryXDoc.Root.Name.[Namespace] = relNs Then
                            Dim urisToCheck = entryXDoc.Descendants(relNs + "Relationship").Where(Function(r) r.Attribute("TargetMode") IsNot Nothing AndAlso CStr(r.Attribute("TargetMode")) = "External")

                            For Each rel In urisToCheck
                                Dim target = CStr(rel.Attribute("Target"))

                                If target IsNot Nothing Then

                                    Try
                                        Dim uri As Uri = New Uri(target)
                                    Catch ex As UriFormatException
                                        Dim newUri As Uri = invalidUriHandler(target)
                                        rel.Attribute("Target").Value = newUri.ToString()
                                        replaceEntry = True
                                    End Try
                                End If
                            Next
                        End If

                    Catch ex As XmlException
                        Continue For
                    End Try
                End Using

                If replaceEntry Then
                    Dim fullName = entry.FullName
                    entry.Delete()
                    Dim newEntry = za.CreateEntry(fullName)

                    Using writer As StreamWriter = New StreamWriter(newEntry.Open())

                        Using xmlWriter As XmlWriter = XmlWriter.Create(writer)
                            entryXDoc.WriteTo(xmlWriter)
                        End Using
                    End Using
                End If
            Next
        End Using
    End Sub
End Module
