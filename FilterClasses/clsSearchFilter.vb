#Region " File Information "

'==============================================================================
' This class represents a complete search filter used to refine a recordset
'==============================================================================

#Region " Revision History "

'==============================================================================
' Name      Date        Description
'------------------------------------------------------------------------------
' KSD       23/02/2007  Implemented.
'==============================================================================

#End Region

#End Region

Imports System.Xml
Imports System.Xml.Serialization

<Serializable()> Public Class clsSearchFilter
    Implements IDisposable

#Region " Members "

    <NonSerialized()> Private m_objDB As clsDB
    Private m_strUserSyntax As String           'This is the boolean search representation
    <NonSerialized()> Private m_strDisplaySyntax As String
    Private m_objFilterGroup As clsSearchGroup  'This is the group containing filters
    Private m_strRootTable As String
    <NonSerialized()> Private m_objTableMask As clsTableMask
    Private m_blnDisposedValue As Boolean
    Private m_blnHasVariables As Boolean = False
    <NonSerialized()> Private m_colVariables As FrameworkCollections.K1Dictionary(Of clsField) 'used to store the variables and their types
    <NonSerialized()> Private m_strName As String
    <NonSerialized()> Private m_intTypeID As Integer = 0
    <NonSerialized()> Private m_intDefaultTypeID As Integer = 0
    <NonSerialized()> Private m_blnDoSecurityCheck As Boolean
    <NonSerialized()> Private m_intCreatedFromSavedSearchID As Integer = clsDBConstants.cintNULL
    <NonSerialized()> Private m_blnCreateFromUser As Boolean = False
#End Region

#Region " Enumeration "

    Public Enum enumComparisonType
        EQUAL = 1
        GREATER_THAN = 2
        LESS_THAN = 3
        GREATER_THAN_EQUAL = 4
        LESS_THAN_EQUAL = 5
        [IN] = 6
        '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
        EXACT = 7
    End Enum

    Public Enum enumOperatorType
        NONE = 0
        [AND] = 1
        [OR] = 2
        ANDNOT = 3
        ORNOT = 4
        [NOT] = 5
    End Enum

    Public Enum enumTokenType
        VALUE = 0
        VARIABLE = 1
        CONSTANT = 2
    End Enum
#End Region

#Region " Constructors "

    Public Sub New()
    End Sub

    ''' <summary>
    ''' Creates a new search filter using boolean search syntax
    ''' </summary>
    Public Sub New(ByVal objDb As clsDB, _
    ByVal strUserSyntax As String, _
    ByVal strRootTable As String)
        m_objDB = objDb
        m_strUserSyntax = strUserSyntax
        m_strRootTable = strRootTable
        m_objFilterGroup = GetGroup()
    End Sub

    ''' <summary>
    ''' Creates a new search filter using a search group object
    ''' </summary>
    Public Sub New(ByVal objDb As clsDB, _
    ByVal objFilterGroup As clsSearchGroup, _
    ByVal strRootTable As String)
        m_objDB = objDb
        m_objFilterGroup = objFilterGroup
        m_strRootTable = strRootTable
    End Sub

    ''' <summary>
    ''' Creates a new search filter using a search group object and table mask object
    ''' </summary>
    Public Sub New(ByVal objDb As clsDB, _
    ByVal objFilterGroup As clsSearchGroup, _
    ByVal objTableMask As clsTableMask)
        m_objDB = objDb
        m_objFilterGroup = objFilterGroup
        m_objTableMask = objTableMask
        If objTableMask IsNot Nothing Then
            m_strRootTable = objTableMask.Table.DatabaseName
        End If
    End Sub

    ''' <summary>
    ''' Creates a basic search filter with only one filter criteria
    ''' </summary>
    Public Sub New(ByVal objDb As clsDB,
                   ByVal strRef As String,
                   ByVal eCompareType As enumComparisonType,
                   ByVal objValue As Object,
                   Optional ByVal strName As String = Nothing)

        m_objDB = objDb

        Dim objSe As New clsSearchElement(enumOperatorType.NONE, strRef, eCompareType, objValue)

        Dim colSOs As New List(Of clsSearchObjBase)
        colSOs.Add(objSe)

        m_objFilterGroup = New clsSearchGroup(enumOperatorType.NONE, colSOs)

        Dim intIndex As Integer = strRef.IndexOf(".", StringComparison.Ordinal)
        m_strRootTable = strRef.Substring(0, intIndex)
        m_strName = strName

    End Sub
#End Region

#Region " Properties "

    <XmlIgnore()> _
    Public Property Database() As clsDB
        Get
            Return m_objDB
        End Get
        Set(ByVal value As clsDB)
            m_objDB = value
        End Set
    End Property

    ''' <summary>
    ''' The XML representation of the Search Filter
    ''' </summary>
    Public ReadOnly Property XML() As String
        Get
            Return GetXML()
        End Get
    End Property

    ''' <summary>
    ''' The user syntax (boolean syntax) of the Search Filter
    ''' </summary>
    Public Property UserSyntax() As String
        Get
            If m_strUserSyntax Is Nothing OrElse m_strUserSyntax.Length = 0 Then
                m_strUserSyntax = GetUserSyntax()
            End If
            Return m_strUserSyntax
        End Get
        Set(ByVal value As String)
            m_strUserSyntax = value
        End Set
    End Property

    ''' <summary>
    ''' The Search Group representation of the filter
    ''' </summary>
    Public Property Group() As clsSearchGroup
        Get
            Return m_objFilterGroup
        End Get
        Set(ByVal value As clsSearchGroup)
            m_objFilterGroup = value
            m_strUserSyntax = Nothing
        End Set
    End Property

    <XmlIgnore()> _
    Public ReadOnly Property TableMask() As clsTableMask
        Get
            If m_objTableMask Is Nothing Then
                GetTableMask()
            End If
            Return m_objTableMask
        End Get
    End Property

    Public Property HasVariables() As Boolean
        Get
            Return m_blnHasVariables
        End Get
        Set(ByVal value As Boolean)
            m_blnHasVariables = value
            If value = False Then
                m_strUserSyntax = GetUserSyntax()
            End If
        End Set
    End Property

    <XmlIgnore()> _
    Public ReadOnly Property VariableCollection() As FrameworkCollections.K1Dictionary(Of clsField)
        Get
            Return m_colVariables
        End Get
    End Property

    <XmlIgnore()> _
    Public Property Name() As String
        Get
            Return m_strName
        End Get
        Set(ByVal value As String)
            m_strName = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TypeID() As Integer
        Get
            If m_intTypeID = 0 Then
                m_intTypeID = GetFilterType()
            End If
            Return m_intTypeID
        End Get
        Set(ByVal value As Integer)
            m_intTypeID = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property DefaultTypeID() As Integer
        Get
            If m_intDefaultTypeID = 0 Then
                m_intDefaultTypeID = GetDefaultTypeID()
            End If
            Return m_intDefaultTypeID
        End Get
        Set(ByVal value As Integer)
            m_intDefaultTypeID = value
        End Set
    End Property

    Public ReadOnly Property RootTable() As String
        Get
            Return m_strRootTable
        End Get
    End Property

    <XmlIgnore()> _
    Public Property CreatedFromSavedSearchID() As Integer
        Get
            Return m_intCreatedFromSavedSearchID
        End Get
        Set(ByVal value As Integer)
            m_intCreatedFromSavedSearchID = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property DisplaySyntax() As String
        Get
            If m_strDisplaySyntax Is Nothing OrElse String.IsNullOrEmpty(m_strUserSyntax) Then
                m_strDisplaySyntax = GetDisplaySyntax()
            End If
            Return m_strDisplaySyntax
        End Get
        Set(ByVal value As String)
            m_strDisplaySyntax = value
        End Set
    End Property
#End Region

#Region " Methods "

#Region " Enumeration String Representations "

    Public Shared Function GetComparisonString(ByVal eCompareType As enumComparisonType) As String
        Dim strType As String = ""

        Select Case eCompareType
            '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
            Case enumComparisonType.EQUAL, enumComparisonType.EXACT
                strType = "="
            Case enumComparisonType.GREATER_THAN
                strType = ">"
            Case enumComparisonType.GREATER_THAN_EQUAL
                strType = ">="
            Case enumComparisonType.IN
                strType = "IN"
            Case enumComparisonType.LESS_THAN
                strType = "<"
            Case enumComparisonType.LESS_THAN_EQUAL
                strType = "<="
        End Select

        Return strType
    End Function

    ''' <summary>
    ''' Returns the SQL representation of the Operator Type
    ''' </summary>
    Public Shared Function GetSQLOperator(ByVal eOpType As enumOperatorType) As String
        Select Case eOpType
            Case clsSearchFilter.enumOperatorType.AND
                Return "AND"
            Case clsSearchFilter.enumOperatorType.ANDNOT
                Return "AND NOT"
            Case clsSearchFilter.enumOperatorType.OR
                Return "OR"
            Case clsSearchFilter.enumOperatorType.ORNOT
                Return "OR NOT"
            Case clsSearchFilter.enumOperatorType.NOT
                Return "NOT"
            Case Else
                Return ""
        End Select
    End Function
#End Region

#Region " XML Methods "

    Public Shared Function Serialize(ByVal strXML As String) As clsSearchFilter
        Dim objSerializer As New XmlSerializer(GetType(clsSearchFilter))
        Dim objTextReader As New StringReader(strXML)
        Dim objXmlReader As New XmlTextReader(objTextReader)

        If objSerializer.CanDeserialize(objXmlReader) Then
            Return CType(objSerializer.Deserialize(objXmlReader), clsSearchFilter)
        Else
            Return Nothing
        End If
    End Function

    Public Function DeSerialize() As String
        Dim objSB As New Text.StringBuilder
        Dim objSW As New System.IO.StringWriter(objSB)
        Dim objXmlSerializer As New XmlSerializer(Me.GetType())

        objXmlSerializer.Serialize(objSW, Me)
        objSW.Close()

        Return objSB.ToString
    End Function

    Private Function GetXML() As String
        Dim strXML As String

        strXML = DeSerialize()

        Return strXML
    End Function

#End Region

#Region " User Syntax Methods "

    Private Function GetDisplaySyntax() As String
        Dim strUserSyntax As String = ""

        If m_objFilterGroup IsNot Nothing AndAlso m_strRootTable IsNot Nothing Then
            RecurseTableMask(strUserSyntax, m_objFilterGroup, True)
        End If

        Return strUserSyntax
    End Function

    Private Function GetUserSyntax() As String
        Dim strUserSyntax As String = ""

        If m_objFilterGroup IsNot Nothing AndAlso m_strRootTable IsNot Nothing Then
            RecurseTableMask(strUserSyntax, m_objFilterGroup, True)
        End If

        Return strUserSyntax
    End Function

    Private Sub RecurseTableMask(ByRef strSyntax As String, ByVal objSG As clsSearchGroup, Optional ByVal blnDisplayOnly As Boolean = False)
        strSyntax &= "("

        For Each objSO As clsSearchObjBase In objSG.SearchObjs
            Dim strOp As String = clsSearchFilter.GetSQLOperator(objSO.OperatorType)

            If strOp.Length > 0 Then
                strSyntax &= " " & strOp & " "
            End If

            If TypeOf objSO Is clsSearchGroup Then
                RecurseTableMask(strSyntax, CType(objSO, clsSearchGroup), blnDisplayOnly)
            Else
                Dim objSE As clsSearchElement = CType(objSO, clsSearchElement)

                If objSE.IsVariable Then
                    strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                        "[" & CType(objSE.Value, String) & "]"
                ElseIf objSE.IsConstant Then
                    'System variable
                    If blnDisplayOnly Then
                        Dim objField As clsField = GetFieldFromRef(objSE.FieldRef)
                        Dim dtDate As Date = CDate(objSE.Value)

                        'Want to display the variables in the search filter tooltip
                        Dim strValue As String = objSE.ConstantValue

                        'If objField.IsDateType Then
                        '    If (objField.DateType = clsDBConstants.enumDateTypes.DATE_AND_TIME OrElse _
                        '    objField.DateType = clsDBConstants.enumDateTypes.DATE_ONLY) Then
                        '        strValue = objField.Database.SysInfo.ToLocalTime(strValue)
                        '    End If

                        '    If (objField.DateType = clsDBConstants.enumDateTypes.TIME_ONLY) Then
                        '        strValue = CDate(objSE.Value).ToString("hh:mm:ss tt")
                        '    End If
                        'End If

                        strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                              strValue
                    Else
                        strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                             "#" & CType(objSE.Value, String) & "#"
                    End If
                Else
                    If TypeOf objSE.Value Is Hashtable Then
                        strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                            """" & CreateIDStringFromCollection(CType(objSE.Value, Hashtable).Values) & """"
                    ElseIf TypeOf objSE.Value Is Date Then
                        Dim objField As clsField = GetFieldFromRef(objSE.FieldRef)

                        Dim strValue As String = CType(objSE.Value, String)
                        If objField.IsDateType Then
                            If (objField.DateType = clsDBConstants.enumDateTypes.DATE_AND_TIME OrElse _
                            objField.DateType = clsDBConstants.enumDateTypes.DATE_ONLY) Then
                                strValue = CType(objField.Database.Profile.ToLocalTime(CDate(strValue)), String)
                            End If

                            If (objField.DateType = clsDBConstants.enumDateTypes.TIME_ONLY) Then
                                strValue = CDate(objSE.Value).ToString("hh:mm:ss tt")
                            End If
                        End If

                        strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                            """" & strValue & """"
                    Else
                        strSyntax &= objSE.FieldRef & " " & GetComparisonString(objSE.CompareType) & " " & _
                            """" & CType(objSE.Value, String) & """"
                    End If
                End If
            End If
        Next

        strSyntax &= ")"
    End Sub
#End Region

#Region " Group Methods "

    Private Function GetGroup() As clsSearchGroup
        Dim objGroup As clsSearchGroup = Nothing

        If Not m_strUserSyntax Is Nothing Then
            objGroup = GetGroupFromUserSyntax()

            'check if we have just added unnecessary parentheses around the group
            Dim objNewGroup As clsSearchGroup = Nothing

            If objGroup.SearchObjs IsNot Nothing AndAlso _
            objGroup.SearchObjs.Count = 1 AndAlso _
            TypeOf objGroup.SearchObjs(0) Is clsSearchGroup Then
                objNewGroup = CType(objGroup.SearchObjs(0), clsSearchGroup)
            End If

            'if so, remove the outer group and re-assign group to first sub-group
            If Not objNewGroup Is Nothing Then
                objGroup.SearchObjs.Clear()
                objGroup = objNewGroup
            End If
        End If

        Return objGroup
    End Function

#Region " GetGroupFromUserSyntax "

    Public Function GetGroupFromUserSyntax() As clsSearchGroup
        Dim strQuery As String
        Dim objSG As clsSearchGroup = Nothing

        strQuery = m_strUserSyntax.Trim
        strQuery = strQuery.Replace(vbCr, "")
        strQuery = strQuery.Replace(vbLf, "")

        If strQuery.Length = 0 Then
            Return objSG
        Else
            objSG = New clsSearchGroup(enumOperatorType.NONE, Nothing)
        End If

        Dim colTokens As Hashtable = GetTokens(strQuery)

        RunQueryProcess(strQuery, 0, enumOperatorType.NONE, objSG, 1, colTokens)

        m_colVariables = New FrameworkCollections.K1Dictionary(Of clsField)
        m_blnDoSecurityCheck = m_objDB.Profile IsNot Nothing

        m_blnCreateFromUser = True

        'Calls during saving and searching
        ValidateGroup(objSG, m_colVariables, True)
        m_blnCreateFromUser = False

        Return objSG
    End Function
#End Region

#Region " Tokens "

    Private Function GetTokens(ByRef strQuery As String) As Hashtable
        Dim colTokens As New Hashtable

        'Must do values before variables
        AddValueToken(strQuery, colTokens, enumTokenType.VALUE)
        AddValueToken(strQuery, colTokens, enumTokenType.CONSTANT)
        AddValueToken(strQuery, colTokens, enumTokenType.VARIABLE)

        Return colTokens
    End Function

    Private Sub AddValueToken(ByRef strQuery As String, ByVal colTokens As Hashtable, _
    ByVal eTokenType As enumTokenType)
        Dim intStart As Integer
        Dim intRef As Integer
        Dim intEnd As Integer
        Dim chrStart As Char
        Dim chrEnd As Char

        Select Case eTokenType
            Case enumTokenType.VALUE
                chrStart = """"c
                chrEnd = """"c
            Case enumTokenType.VARIABLE
                chrStart = "["c
                chrEnd = "]"c
            Case enumTokenType.CONSTANT
                chrStart = "#"c
                chrEnd = "#"c
        End Select

        intStart = strQuery.IndexOf(chrStart)
        intRef = intStart

        While intStart >= 0

            intEnd = strQuery.IndexOf(chrEnd, intRef + 1)

            If intEnd < 0 Then
                Select Case eTokenType
                    Case enumTokenType.VALUE
                        Throw New Exception("A value in the search string is missing it's end "" character." & vbCrLf & _
                            "All values should be enclosed in double quotes." & vbCrLf & vbCrLf & _
                            "Note: If you would like to include the double quote character in a value, " & vbCrLf & _
                            "please use two double quotes to represent a single double quote.")
                    Case enumTokenType.VARIABLE
                        Throw New Exception("A variable in the search string is missing it's end ] character." & vbCrLf & _
                            "All variables should be enclosed in quare brackets.")
                    Case enumTokenType.CONSTANT
                        Throw New Exception("A constant used in the search string is missing it's end # character." & vbCrLf & _
                            "All constants should be enclosed in hash characters.")
                End Select
            End If

            If eTokenType = enumTokenType.VALUE AndAlso _
            intEnd + 1 <= strQuery.Length - 1 AndAlso _
            strQuery.Chars(intEnd + 1) = chrEnd Then
                'skip two double quotes contained within a value
                intRef = intEnd + 1
            Else
                Dim strKey As String = "~" & colTokens.Count & "~"
                Dim strValue As String

                'Get the whole value string
                strValue = strQuery.Substring(intStart, (intEnd + 1) - intStart)
                'Strip off double quotes from start and end, and replace all paired double quotes with a single double quote
                strValue = strValue.Substring(1, strValue.Length - 2).Replace("""""", """")

                If strValue.Trim.Length = 0 Then
                    strValue = Nothing
                End If

                colTokens.Add(strKey, New clsSearchValue(strValue, eTokenType))

                'replace the value from the original query with the key representation
                strQuery = strQuery.Substring(0, intStart) & strKey & _
                    strQuery.Substring(intEnd + 1, strQuery.Length - (intEnd + 1))

                intStart = strQuery.IndexOf(chrStart)
                intRef = intStart

                If eTokenType = enumTokenType.VARIABLE Then
                    m_blnHasVariables = True
                End If
            End If
        End While
    End Sub
#End Region

#Region " Group Creation "

    Private Sub SkipWhiteSpace(ByVal strQuery As String, ByRef intIndex As Integer)
        While strQuery.Chars(intIndex) = " "c
            intIndex += 1
        End While
    End Sub

    Private Sub RunQueryProcess(ByVal strQuery As String, ByRef intIndex As Integer, _
    ByVal eOpType As enumOperatorType, ByVal objSG As clsSearchGroup, ByRef intGroupLevel As Integer, _
    ByVal colTokens As Hashtable)
        Dim objSE As clsSearchElement
        Dim blnRepeat As Boolean = True

        While blnRepeat AndAlso intIndex < strQuery.Length
            objSE = ProcessStep1(strQuery, intIndex, eOpType, objSG, intGroupLevel, colTokens)
            If objSE IsNot Nothing Then
                ProcessStep2(strQuery, intIndex, objSE)
                ProcessStep3(strQuery, intIndex, objSE, colTokens)
            End If
            blnRepeat = ProcessStep4(strQuery, intIndex, eOpType, intGroupLevel)
        End While
    End Sub

    Private Function ProcessStep1(ByVal strQuery As String, _
    ByRef intIndex As Integer, ByVal eOpType As enumOperatorType, _
    ByVal objSG As clsSearchGroup, ByRef intGroupLevel As Integer, _
    ByVal colTokens As Hashtable) As clsSearchElement
        'Move through the "(" and white space in the user syntax
        While intIndex <= strQuery.Length - 1 AndAlso _
        (strQuery.Chars(intIndex) = "("c OrElse _
        strQuery.Chars(intIndex) = " "c)

            If strQuery.Chars(intIndex) = "("c Then
                Dim objNewGroup As New clsSearchGroup(eOpType)
                eOpType = enumOperatorType.NONE
                intGroupLevel += 1
                intIndex += 1

                If objSG.SearchObjs Is Nothing Then
                    objSG.SearchObjs = New List(Of clsSearchObjBase)
                End If
                objSG.SearchObjs.Add(objNewGroup)

                RunQueryProcess(strQuery, intIndex, eOpType, objNewGroup, intGroupLevel, colTokens)

                Return Nothing
            Else
                intIndex += 1
            End If
        End While

        If intIndex = strQuery.Length Then
            Throw New Exception("Missing a search criteria field.")
        End If

        'Check for the next expected characters
        Dim arrEndChars As Char() = {"="c, ">"c, "<"c, " "c}
        Dim intEndIndex As Integer = strQuery.IndexOfAny(arrEndChars, intIndex)

        If intEndIndex = -1 Then
            Throw New Exception("The following search criteria is missing a proper operator: " & _
                strQuery.Substring(intIndex, strQuery.Length - intIndex) & " (Expected: =, >, <, >=, <=, or IN).")
        End If

        'Get the criteria string (i.e. MetadataProfile.ExternalID)
        Dim strCriteria As String = strQuery.Substring(intIndex, intEndIndex - intIndex)
        If strCriteria.ToUpper = "NOT" Then
            Return Nothing
        End If

        intIndex = intEndIndex

        SkipWhiteSpace(strQuery, intIndex)

        'Check for the unexpected characters in the search criteria
        Dim arrBadChars As Char() = {"("c, "~"c, ")"c}
        If strCriteria.IndexOfAny(arrBadChars) >= 0 Then
            Throw New Exception("Invalid search criteria field: " & strCriteria)
        End If

        Dim objSE As New clsSearchElement(eOpType, strCriteria)
        If objSG.SearchObjs Is Nothing Then
            objSG.SearchObjs = New List(Of clsSearchObjBase)
        End If
        objSG.SearchObjs.Add(objSE)

        Return objSE
    End Function

    Private Sub ProcessStep2(ByVal strQuery As String, ByRef intIndex As Integer, _
    ByVal objSE As clsSearchElement)
        Dim arrEndChars As Char() = {"~"c, " "c}
        Dim intEndIndex As Integer = strQuery.IndexOfAny(arrEndChars, intIndex)

        If intEndIndex = -1 Then
            Throw New Exception("The following search criteria is missing a value: " & objSE.FieldRef)
        End If

        Dim strValueType As String = strQuery.Substring(intIndex, intEndIndex - intIndex)
        intIndex = intEndIndex

        SkipWhiteSpace(strQuery, intIndex)

        Select Case strValueType
            Case "="
                objSE.CompareType = enumComparisonType.EQUAL
            Case ">"
                objSE.CompareType = enumComparisonType.GREATER_THAN
            Case "<"
                objSE.CompareType = enumComparisonType.LESS_THAN
            Case "<="
                objSE.CompareType = enumComparisonType.LESS_THAN_EQUAL
            Case ">="
                objSE.CompareType = enumComparisonType.GREATER_THAN_EQUAL
            Case "IN"
                objSE.CompareType = enumComparisonType.IN
            Case Else
                Throw New Exception("The following search criteria is missing a proper operator: " & _
                strQuery.Substring(intIndex, strQuery.Length - intIndex) & " (Expected: =, >, <, >=, <=, or IN).")
        End Select
    End Sub

    Private Sub ProcessStep3(ByVal strQuery As String, ByRef intIndex As Integer, _
    ByVal objSE As clsSearchElement, ByVal colTokens As Hashtable)
        Dim intEndIndex As Integer = strQuery.IndexOf("~", intIndex + 1)

        If intEndIndex = -1 Then
            Throw New Exception("The following search criteria is missing a value: " & objSE.FieldRef)
        Else
            intEndIndex += 1
        End If

        Dim strKey As String = strQuery.Substring(intIndex, intEndIndex - intIndex)
        intIndex = intEndIndex

        Dim objSV As clsSearchValue = CType(colTokens(strKey), clsSearchValue)

        If objSV Is Nothing Then
            Throw New Exception("The following search criteria is missing a value: " & objSE.FieldRef)
        Else
            objSE.IsVariable = objSV.IsVariable
            objSE.IsConstant = objSV.IsConstant
            If objSV.IsConstant Then
                ' objSE.Value = GetConstantValue(objSV.Value)
                'objSE.Value = objSV.Value
                objSE.ConstantValue = "#" & CType(objSV.Value, String) & "#"

            Else
                objSE.Value = objSV.Value
            End If

            If objSV.IsVariable Then
                m_blnHasVariables = True
            End If
        End If
    End Sub

    Private Function ProcessStep4(ByVal strQuery As String, ByRef intIndex As Integer, _
    ByRef eOpType As enumOperatorType, ByRef intGroupLevel As Integer) As Boolean
        While intIndex <= strQuery.Length - 1 AndAlso _
        (strQuery.Chars(intIndex) = ")"c OrElse strQuery.Chars(intIndex) = " "c)
            If strQuery.Chars(intIndex) = ")"c Then
                If intGroupLevel = 1 Then
                    Throw New Exception("Too many end parentheses in the query.")
                Else
                    intIndex += 1
                    intGroupLevel -= 1
                    Return False
                End If
            Else
                intIndex += 1
            End If
        End While

        If intIndex = strQuery.Length Then
            If Not intGroupLevel = 1 Then
                Throw New Exception("Too many begin parentheses in the query.")
            Else
                Return False
            End If
        End If

        Dim arrEndChars As Char() = {" "c, "("c}
        Dim intEndIndex As Integer = strQuery.IndexOfAny(arrEndChars, intIndex)

        If intEndIndex = -1 Then
            Throw New Exception("Unexpected characters at the end of the query.")
        End If

        Dim strOp As String = strQuery.Substring(intIndex, intEndIndex - intIndex)
        intIndex = intEndIndex
        SkipWhiteSpace(strQuery, intIndex)

        Dim strNot As String = ""
        intEndIndex = strQuery.IndexOfAny(arrEndChars, intIndex)

        If intEndIndex >= 0 Then
            strNot = strQuery.Substring(intIndex, intEndIndex - intIndex)
            If strNot.ToUpper = "NOT" Then
                strOp &= " NOT"
                intIndex = intEndIndex
            End If
        End If

        Select Case strOp.ToUpper
            Case "AND"
                eOpType = enumOperatorType.AND
            Case "OR"
                eOpType = enumOperatorType.OR
            Case "AND NOT", "ANDNOT"
                eOpType = enumOperatorType.ANDNOT
            Case "OR NOT", "ORNOT"
                eOpType = enumOperatorType.ORNOT
            Case "NOT"
                eOpType = enumOperatorType.NOT
            Case Else
                Throw New Exception("Invalid connector: " & strOp & vbCrLf & _
                    "Expected: AND, OR, AND NOT, OR NOT.")
        End Select

        Return True
    End Function
#End Region

#Region " Tables and Variable XML "
    'Added blnSavingSearch optional boolean so that system variables are either evaluated or saved depending on a search or a save

    Public Sub ValidateFilter(Optional ByVal strRootTable As String = Nothing, Optional ByVal blnSavingSearch As Boolean = False)
        If m_objFilterGroup Is Nothing Then
            Throw New Exception("This is an invalid search filter as it has no search criteria.")
        End If

        If m_strRootTable Is Nothing Then
            If strRootTable Is Nothing Then
                Throw New Exception("To validate this saved search, " & _
                    "please pass in the root table to the ValidateFilter function.")
            Else
                m_strRootTable = strRootTable
            End If
        End If

        If m_colVariables IsNot Nothing Then
            m_colVariables.Dispose()
            m_colVariables = Nothing
        End If

        m_colVariables = New FrameworkCollections.K1Dictionary(Of clsField)
        m_blnDoSecurityCheck = m_objDB.Profile IsNot Nothing

        ValidateGroup(m_objFilterGroup, m_colVariables, blnSavingSearch)
    End Sub

    Private Sub ValidateGroup(ByVal objSG As clsSearchGroup, _
    ByRef colVariableFields As FrameworkCollections.K1Dictionary(Of clsField), Optional ByVal blnSavingSearch As Boolean = False)
        If objSG.SearchObjs Is Nothing Then
            Throw New Exception("There is an empty group in the filter. The filter is invalid.")
        End If

        For Each objSO As clsSearchObjBase In objSG.SearchObjs
            If TypeOf objSO Is clsSearchGroup Then
                ValidateGroup(CType(objSO, clsSearchGroup), colVariableFields, blnSavingSearch)
            Else
                ValidateElement(CType(objSO, clsSearchElement), colVariableFields, blnSavingSearch)
            End If
        Next
    End Sub

    Private Sub ValidateElement(ByVal objSE As clsSearchElement, _
    ByVal colVariableFields As FrameworkCollections.K1Dictionary(Of clsField), Optional ByVal blnSavingSearch As Boolean = False)
        Dim objField As clsField = GetFieldFromRef(objSE.FieldRef)

        'Check our system variable constants
        If objSE.IsConstant Then
            'Remove # characters from variable string in order to perform checking
            Dim strConstantValue As String = objSE.ConstantValue.Substring(1, objSE.ConstantValue.Length - 2).Replace("""""", """")
            If blnSavingSearch Then
                'Just check if system variable is valid, but keep the system variable name
                If Not objSE.IsConstantSaved Then
                    GetConstantValue(strConstantValue)
                End If
                Return
            Else
                'Replaying search, evaluate system variable value
                If Not objSE.IsConstantSaved Then
                    objSE.Value = GetConstantValue(strConstantValue)
                    objSE.IsConstantSaved = True
                End If
            End If
        End If

        If objSE.IsVariable Then
            If objSE.Value Is Nothing Then
                Throw New Exception("Variable names cannot be blank.")
            End If

            If colVariableFields(CStr(objSE.Value)) Is Nothing Then
                colVariableFields.Add(CStr(objSE.Value), objField)
            Else
                Dim objCompareField As clsField = CType(colVariableFields(CStr(objSE.Value)), clsField)
                If Not CompareFields(objCompareField, objField) Then
                    Throw New Exception("A variable is used for more than one criteria but with conflicting data types.")
                End If
            End If
        Else
            If objSE.CompareType = enumComparisonType.IN Then
                If objSE.Value Is Nothing Then
                    Throw New Exception("Must include at least one value when using the IN operator.")
                End If

                If Not objField.IsForeignKey AndAlso Not objField.DatabaseName.ToUpper = clsDBConstants.Fields.cID Then
                    Throw New Exception("The IN operator can only be used with foreign key fields or the ID field.")
                End If

                If TypeOf objSE.Value Is Hashtable Then
                    Return
                End If

                If TypeOf objSE.Value Is String Then
                    Dim arrIds As String() = objSE.Value.ToString.Split(","c)

                    objSE.Value = New Hashtable

                    For intLoop As Integer = 0 To arrIds.Length - 1
                        Dim objVal As Object = CheckProperValue(objSE, arrIds(intLoop), objField)
                        CType(objSE.Value, Hashtable).Add(CStr(objVal), objVal)
                    Next
                Else
                    Throw New Exception("The value '" & CType(objSE.Value, String) & "' is inconsistent with the data type " & _
                        "for the field '" & objSE.FieldRef & "'.")
                End If
            Else
                If objSE.Value Is Nothing Then
                    If objSE.CompareType = enumComparisonType.GREATER_THAN_EQUAL Then
                        Throw New Exception("You must have a value when the comparison is greater than or equal.")
                    ElseIf objSE.CompareType = clsSearchFilter.enumComparisonType.LESS_THAN Then
                        Throw New Exception("You must have a value when the comparison is less than.")
                    Else
                        Return
                    End If
                End If

                objSE.Value = CheckProperValue(objSE, objSE.Value, objField)
            End If
        End If

    End Sub

    Private Function GetFieldFromRef(ByVal strFieldRef As String) As clsField
        Dim arrRefs As String() = Split(strFieldRef, ".")
        Dim objTable As clsTable = Nothing
        Dim objField As clsField = Nothing
        Dim objRelTable As clsTable = Nothing
        Dim strTable As String
        Dim strField As String
        Dim strRef As String = ""

        'TODO: Need to validate user has security to all tables, field, and field links involved

        For intLoop As Integer = 0 To arrRefs.Length - 1
            If objTable Is Nothing Then
                strTable = arrRefs(intLoop)

                If Not CStr(arrRefs(intLoop)).ToUpper = m_strRootTable.ToUpper Then
                    Throw New Exception("The search criteria '" & strFieldRef & "' does not begin " & _
                        "with the root table '" & m_strRootTable & "'.")
                End If

                objTable = m_objDB.SysInfo.Tables(strTable)

                If objTable Is Nothing Then
                    Throw New Exception("A table called '" & strTable & "' does not exist in the RecFind 6 database.")
                ElseIf m_blnDoSecurityCheck AndAlso Not objTable.HasAccess Then
                    Throw New Exception("You do not have the required security to access the table '" & strTable & "'.")
                End If

                strRef &= objTable.DatabaseName
            Else
                strField = arrRefs(intLoop)

                If strField.Chars(0) = "*"c Then
                    'foreign key table or link table
                    strTable = strField.Substring(1, strField.Length - 1)
                    objRelTable = m_objDB.SysInfo.Tables(strTable)

                    If objRelTable Is Nothing Then
                        Throw New Exception("A table called '" & strTable & "' does not exist in the RecFind 6 database.")
                    ElseIf m_blnDoSecurityCheck AndAlso Not objTable.HasAccess Then
                        Throw New Exception("You do not have the required security to access the table '" & strTable & "'.")
                    End If
                Else
                    If objRelTable Is Nothing Then
                        objField = m_objDB.SysInfo.Fields(objTable.ID & "_" & strField)

                        If objField Is Nothing Then
                            Throw New Exception("A field called '" & strField & "' belonging to table '" & _
                                objTable.DatabaseName & "' does not exist in the RecFind 6 database.")
                        ElseIf m_blnDoSecurityCheck _
                        AndAlso ((Not objField.IsForeignKey _
                                  AndAlso Not m_objDB.Profile.HasFieldAccess(objField, clsDBConstants.cintNULL)) _
                                 OrElse (objField.IsForeignKey _
                                         AndAlso Not m_objDB.Profile.HasAccess(objField.SecurityID))) Then
                            Throw New Exception("You do not have the required security to access the field '" & strField & _
                                                "' belonging to table '" & objTable.DatabaseName & "'.")
                        ElseIf m_blnDoSecurityCheck AndAlso objField.IsForeignKey AndAlso Not objField.FieldLink.IdentityTable.HasAccess Then
                            Throw New Exception("You do not have the required security to access the table '" & objField.FieldLink.IdentityTable.CaptionText & "'.")
                        End If

                        If Not intLoop = arrRefs.Length - 1 Then
                            If objField.IsForeignKey Then
                                objTable = objField.FieldLink.IdentityTable
                            Else
                                Throw New Exception("Field '" & strField & "' within table '" & objTable.DatabaseName & "' " & _
                                    "is not a foreign key field, and therfore the syntax in the search criteria '" & _
                                    strFieldRef & "' is incorrect.")
                            End If
                        End If
                    Else
                        objField = m_objDB.SysInfo.Fields(objRelTable.ID & "_" & strField)

                        If objField Is Nothing Then
                            Throw New Exception("A field called '" & strField & "' belonging to table '" & _
                                objRelTable.DatabaseName & "' does not exist in the RecFind 6 database.")
                        ElseIf m_blnDoSecurityCheck _
                        AndAlso ((Not objField.IsForeignKey _
                                  AndAlso Not m_objDB.Profile.HasFieldAccess(objField, clsDBConstants.cintNULL)) _
                                 OrElse (objField.IsForeignKey _
                                         AndAlso Not m_objDB.Profile.HasAccess(objField.SecurityID))) Then
                            Throw New Exception("You do not have the required security to access the field '" & strField & _
                                                "' belonging to table '" & objRelTable.DatabaseName & "'.")
                        End If

                        If Not objField.IsForeignKey Then
                            Throw New Exception("Field '" & strField & "' within table '" & objRelTable.DatabaseName & "' " & _
                                "is not a foreign key field, and therfore the syntax in the search criteria '" & _
                                        strFieldRef & "' is incorrect.")
                        End If

                        If Not objField.FieldLink.IdentityTable.DatabaseName = objTable.DatabaseName Then
                            Throw New Exception("Field '" & strField & "' within table '" & objRelTable.DatabaseName & "' " & _
                                "is not a foreign key to the parent table '" & objTable.DatabaseName & _
                                "', and therfore the syntax in the search criteria '" & _
                                strFieldRef & "' is incorrect.")
                        End If

                        objTable = objRelTable
                        objRelTable = Nothing
                    End If
                End If
            End If
        Next

        Return objField
    End Function

    Private Function CheckProperValue(ByVal objSE As clsSearchElement, ByVal objStartValue As Object, _
    ByVal objField As clsField) As Object
        Dim objValue As Object = objStartValue

        Try
            Select Case objField.DataType
                Case SqlDbType.BigInt, SqlDbType.Int, SqlDbType.SmallInt, _
                SqlDbType.TinyInt
                    objValue = CType(objValue, Int64)
                Case SqlDbType.Decimal, SqlDbType.Float, SqlDbType.Money, _
                SqlDbType.Real, SqlDbType.SmallMoney
                    objValue = CType(objValue, Double)
                Case SqlDbType.Bit
                    objValue = CType(objValue, Boolean)
                Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                    objValue = CType(objValue, Date)
                    If m_blnCreateFromUser AndAlso _
                    (objField.DateType = clsDBConstants.enumDateTypes.DATE_AND_TIME OrElse _
                    objField.DateType = clsDBConstants.enumDateTypes.DATE_ONLY) Then
                        objValue = objField.Database.Profile.ToLocalTime(CDate(objValue), True)
                    End If
            End Select
        Catch ex As Exception
            Throw New Exception("The value '" & CType(objValue, String) & "' is inconsistent with the data type " & _
                "for the field '" & objSE.FieldRef & "'.")
        End Try

        Select Case objField.DataType
            Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.VarChar
                If Not objField.Length = clsDBConstants.cintNULL AndAlso objValue.ToString.Replace("*", "").Length > objField.Length Then
                    Throw New Exception("The length of the value """ & CType(objValue, String) & """ is greater " & _
                        "than the designated length (" & objField.Length & " characters) for field """ & objSE.FieldRef & """")
                End If
            Case SqlDbType.Decimal
                CheckScale(objField, CType(objValue, String), objField.Scale, objField.Length)
                Dim intNumChars As Integer = objField.Length - objField.Scale
                CheckRange(objField, CType(objValue, String), _
                    Math.Max((Int32.MinValue + 1), (0 - CType("".PadRight(intNumChars, "9"c), Int64))), _
                    CType("".PadRight(intNumChars, "9"c), Int64))
            Case SqlDbType.Float
                CheckScale(objField, CType(objValue, String), 16)
                CheckRange(objField, CType(objValue, String), (Int32.MinValue + 1), Decimal.MaxValue)
            Case SqlDbType.Real
                CheckScale(objField, CType(objValue, String), 8)
                CheckRange(objField, CType(objValue, String), (Int32.MinValue + 1), Decimal.MaxValue)
            Case SqlDbType.Money
                CheckScale(objField, CType(objValue, String), 4)
                CheckRange(objField, CType(objValue, String), _
                    CType(Int32.MinValue / 1000, Decimal), _
                    CType(Int64.MaxValue / 1000, Decimal))
            Case SqlDbType.SmallMoney
                CheckScale(objField, CType(objValue, String), 4)
                CheckRange(objField, CType(objValue, String), _
                    CType(Int32.MinValue / 1000, Decimal), _
                    CType(Int32.MaxValue / 1000, Decimal))
            Case SqlDbType.BigInt
                CheckRange(objField, CType(objValue, String), (Int32.MinValue + 1), Int64.MaxValue)
            Case SqlDbType.Int
                CheckRange(objField, CType(objValue, String), (Int32.MinValue + 1), Int32.MaxValue)
            Case SqlDbType.SmallInt
                CheckRange(objField, CType(objValue, String), Int16.MinValue, Int16.MaxValue)
            Case SqlDbType.TinyInt
                CheckRange(objField, CType(objValue, String), Byte.MinValue, Byte.MaxValue)
        End Select

        Return objValue
    End Function

    Private Sub CheckScale(ByVal objField As clsField, ByVal strValue As String, ByVal intScale As Integer, _
    Optional ByVal intPrecision As Integer = 0)
        Dim strCheck As String
        strValue = strValue.Trim
        Dim intIndex As Integer = strValue.IndexOf(".")

        If intIndex >= 0 Then
            strCheck = strValue.Substring(intIndex + 1, strValue.Length - (intIndex + 1))
            If strCheck.Length > intScale Then
                Throw New Exception(objField.DatabaseName & " should be a decimal number with max digits " & _
                    "right of the decimal point equal to " & intScale & ".")
            End If
        End If

        If intPrecision > 0 Then
            If strValue.Chars(0) = "-"c Then
                strValue = strValue.Substring(1, strValue.Length - 1)
                intIndex = strValue.IndexOf(".")
            End If

            If intIndex >= 0 Then
                strCheck = strValue.Substring(0, intIndex)
            Else
                strCheck = strValue
            End If

            If strCheck.Length > (intPrecision - intScale) Then
                Throw New Exception(objField.DatabaseName & " should be a decimal number with max digits " & _
                    "left of the decimal point equal to " & (intPrecision - intScale) & ".")
            End If
        End If
    End Sub

    Private Sub CheckRange(ByVal objField As clsField, ByVal strValue As String, _
    ByVal intMin As Double, ByVal intMax As Double)
        Dim objNum As Double = CDbl(strValue)

        If objNum < intMin OrElse objNum > intMax Then
            Throw New Exception(objField.DatabaseName & " should be a number between " & intMin & " and " & intMax & ".")
        End If
    End Sub

    Public Shared Function CompareFields(ByVal objField1 As clsField, ByVal objField2 As clsField) As Boolean
        If (objField1.DataType = objField2.DataType) OrElse _
        (objField1.IsTextType AndAlso objField2.IsTextType) OrElse _
        (objField1.IsNumericType AndAlso objField2.IsNumericType) OrElse _
        ((objField1.DataType = SqlDbType.DateTime AndAlso objField2.DataType = SqlDbType.SmallDateTime) OrElse _
        (objField2.DataType = SqlDbType.DateTime AndAlso objField1.DataType = SqlDbType.SmallDateTime)) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#End Region

#Region " Type Dependent Methods "

    ''' <summary>
    ''' Returns a single type if filter only includes a single type filter, otherwise returns NULL_INT
    ''' </summary>
    Private Function GetFilterType() As Integer
        If String.IsNullOrEmpty(m_strRootTable) OrElse _
            m_objFilterGroup Is Nothing Then
            Return clsDBConstants.cintNULL
        End If

        Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strRootTable)
        If objTable Is Nothing Then
            Return clsDBConstants.cintNULL
        End If

        Dim objDT As DataTable = m_objDB.GetDataTableByField(clsDBConstants.Tables.cTYPE, _
            clsDBConstants.Fields.Type.cTABLEID, objTable.ID)

        Dim strRowFilter As String = GetTypeFilterFromSearchFilter()

        If Not strRowFilter Is Nothing Then
            objDT.DefaultView.RowFilter = strRowFilter
        End If

        If objDT.DefaultView.Count = 1 Then
            Return CInt(objDT.DefaultView(0)(clsDBConstants.Fields.cID))
        Else
            Return clsDBConstants.cintNULL
        End If
    End Function

    Public Function GetTypeFilterFromSearchFilter() As String
        Try
            If String.IsNullOrEmpty(m_strRootTable) OrElse _
            m_objFilterGroup Is Nothing Then
                Return Nothing
            End If

            Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strRootTable)
            If objTable Is Nothing Then
                Return Nothing
            End If

            Dim colIncludeIDs As New Hashtable
            Dim colExcludeIDs As New Hashtable

            RecurseGetTypeIDS(objTable, m_objFilterGroup, colIncludeIDs, colExcludeIDs, True, False)

            If colIncludeIDs Is Nothing AndAlso colExcludeIDs Is Nothing Then
                Return Nothing
            Else
                Dim strReturn As String = Nothing

                If colIncludeIDs IsNot Nothing AndAlso colIncludeIDs.Count > 0 Then
                    AppendToString(strReturn, clsDBConstants.Fields.cID & " IN (" & CreateIDStringFromCollection(colIncludeIDs.Values) & ")", " AND ")
                End If

                If colExcludeIDs IsNot Nothing AndAlso colExcludeIDs.Count > 0 Then
                    AppendToString(strReturn, clsDBConstants.Fields.cID & " NOT IN (" & CreateIDStringFromCollection(colExcludeIDs.Values) & ")", " AND ")
                End If

                Return strReturn
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Who knows?
    ''' </summary>
    ''' <param name="objTable"></param>
    ''' <param name="objSO"></param>
    ''' <param name="colIncludeIDs"></param>
    ''' <param name="colExcludeIDs"></param>
    ''' <param name="blnOK"></param>
    ''' <param name="blnNot"></param>
    ''' <remarks>Method is too difficult to follow. Select Cases with no default control flow logic :( This is a bad piece of code!!!</remarks>
    Private Sub RecurseGetTypeIDS(ByVal objTable As clsTable,
                                  ByVal objSO As clsSearchObjBase,
                                  ByRef colIncludeIDs As Hashtable,
                                  ByRef colExcludeIDs As Hashtable,
                                  ByRef blnOK As Boolean,
                                  ByRef blnNot As Boolean)
        If Not blnOK Then
            Return
        End If

        Dim eOpType As clsSearchFilter.enumOperatorType = objSO.OperatorType

        If blnNot Then
            Select Case eOpType
                Case clsSearchFilter.enumOperatorType.AND
                    eOpType = clsSearchFilter.enumOperatorType.ORNOT
                Case clsSearchFilter.enumOperatorType.ANDNOT
                    eOpType = clsSearchFilter.enumOperatorType.AND
                Case clsSearchFilter.enumOperatorType.NONE
                    eOpType = clsSearchFilter.enumOperatorType.NOT
                Case clsSearchFilter.enumOperatorType.NOT
                    eOpType = clsSearchFilter.enumOperatorType.NONE
                Case clsSearchFilter.enumOperatorType.OR
                    eOpType = clsSearchFilter.enumOperatorType.ANDNOT
                Case clsSearchFilter.enumOperatorType.ORNOT
                    eOpType = clsSearchFilter.enumOperatorType.OR
            End Select
        End If

        If eOpType = clsSearchFilter.enumOperatorType.OR OrElse
            eOpType = clsSearchFilter.enumOperatorType.ORNOT Then
            blnOK = False
            'colIncludeIDs = Nothing
            'colExcludeIDs = Nothing
            Return
        End If

        If TypeOf objSO Is clsSearchGroup Then
            Dim objSG As clsSearchGroup = CType(objSO, clsSearchGroup)

            If objSG.OperatorType = clsSearchFilter.enumOperatorType.NOT OrElse
                objSG.OperatorType = clsSearchFilter.enumOperatorType.ANDNOT OrElse
                objSG.OperatorType = clsSearchFilter.enumOperatorType.ORNOT Then
                blnNot = Not blnNot
            End If

            blnOK = True
            Dim colTempIncludeIDs As New Hashtable
            Dim colTempExcludeIDs As New Hashtable

            For Each objSeachObj As clsSearchObjBase In objSG.SearchObjs
                RecurseGetTypeIDS(objTable, objSeachObj, colTempIncludeIDs, colTempExcludeIDs, blnOK, blnNot)
            Next

            If Not blnOK Then
                blnOK = True
                colTempIncludeIDs = Nothing
                colTempExcludeIDs = Nothing
            Else
                For Each objEntry In colTempIncludeIDs.Values
                    If colIncludeIDs(objEntry.ToString) Is Nothing Then
                        colIncludeIDs.Add(objEntry.ToString, objEntry)
                    End If
                Next
                For Each objEntry In colTempExcludeIDs.Values
                    If colExcludeIDs(objEntry.ToString) Is Nothing Then
                        colExcludeIDs.Add(objEntry.ToString, objEntry)
                    End If
                Next
            End If
        Else
            Dim objSE As clsSearchElement = CType(objSO, clsSearchElement)

            Dim strCompareRef As String = objTable.DatabaseName & "." & clsDBConstants.Fields.cTYPEID

            If objSE.FieldRef.ToUpper = strCompareRef.ToUpper Then
                Select Case objSE.CompareType
                    '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
                    Case clsSearchFilter.enumComparisonType.EQUAL, clsSearchFilter.enumComparisonType.EXACT
                        If eOpType = clsSearchFilter.enumOperatorType.ANDNOT OrElse _
                        eOpType = clsSearchFilter.enumOperatorType.NOT Then
                            If colExcludeIDs(objSE.Value.ToString) Is Nothing Then
                                colExcludeIDs.Add(objSE.Value.ToString, objSE.Value)
                            End If
                        Else
                            If colIncludeIDs(objSE.Value.ToString) Is Nothing Then
                                colIncludeIDs.Add(objSE.Value.ToString, objSE.Value)
                            End If
                        End If
                    Case clsSearchFilter.enumComparisonType.IN
                        If eOpType = clsSearchFilter.enumOperatorType.ANDNOT OrElse
                            eOpType = clsSearchFilter.enumOperatorType.NOT Then

                            For Each intID As Integer In CType(objSE.Value, Hashtable).Values
                                If colExcludeIDs(intID.ToString) Is Nothing Then
                                    colExcludeIDs.Add(intID.ToString, intID)
                                End If
                            Next
                        Else
                            '2016-03-17 -- Peter Melisi -- Bug fix for #1600003054
                            If (TypeOf objSE.Value Is String) Then
                                Dim searchVals = objSE.Value.ToString.Split(New Char() {","c})

                                If (searchVals IsNot Nothing) Then
                                    For Each intId As String In searchVals
                                        If colIncludeIDs(intId) Is Nothing Then
                                            colIncludeIDs.Add(intId, intId)
                                        End If
                                    Next
                                End If
                            Else
                                Dim searchVals = TryCast(objSE.Value, Hashtable)
                                If (searchVals IsNot Nothing) Then
                                    For Each intId As Integer In searchVals.Values
                                        If colIncludeIDs(intId.ToString) Is Nothing Then
                                            colIncludeIDs.Add(intId.ToString, intId)
                                        End If
                                    Next
                                End If
                            End If
                        End If
                End Select
            End If
        End If
    End Sub

    Private Function GetDefaultTypeID() As Integer
        If String.IsNullOrEmpty(m_strRootTable) OrElse _
        m_objFilterGroup Is Nothing Then
            Return clsDBConstants.cintNULL
        End If

        Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strRootTable)
        If objTable Is Nothing Then
            Return clsDBConstants.cintNULL
        End If

        Dim objDT As DataTable = m_objDB.GetDataTableByField(clsDBConstants.Tables.cTYPE, _
            clsDBConstants.Fields.Type.cTABLEID, objTable.ID)

        Dim strRowFilter As String = GetTypeFilterFromSearchFilter()

        If Not strRowFilter Is Nothing Then
            objDT.DefaultView.RowFilter = strRowFilter
        End If

        Dim intMDPType As Integer = -1

        For intLoop As Integer = 0 To objDT.DefaultView.Count - 1
            Dim intID As Integer = CInt(objDT.DefaultView(intLoop)(clsDBConstants.Fields.cID))

            If m_objDB.SysInfo.K1Groups.TypeGroups.ContainsKey(CStr(intID)) Then
                If intMDPType = -1 Then
                    intMDPType = m_objDB.SysInfo.K1Groups.TypeGroups(CStr(intID))
                Else
                    If Not intMDPType = m_objDB.SysInfo.K1Groups.TypeGroups(CStr(intID)) Then
                        intMDPType = -1
                        Exit For
                    End If
                End If
            Else
                intMDPType = -1
                Exit For
            End If
        Next

        If intMDPType = -1 Then
            Return clsDBConstants.cintNULL
        Else
            Return m_objDB.SysInfo.K1Groups.GetDefaultType(CType(intMDPType, clsDBConstants.enumMDPTypeCodes))
        End If
    End Function
#End Region

#Region " Table Mask Methods "

    Private Sub GetTableMask()
        If String.IsNullOrEmpty(m_strRootTable) OrElse _
        m_objFilterGroup Is Nothing Then
            Return
        End If

        Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strRootTable)
        If objTable Is Nothing Then
            Return
        End If

        Dim intTypeID As Integer = clsDBConstants.cintNULL
        If objTable.TypeDependent Then
            intTypeID = GetFilterType()

            If intTypeID = clsDBConstants.cintNULL AndAlso _
            objTable.DatabaseName = clsDBConstants.Tables.cMETADATAPROFILE Then
                intTypeID = GetDefaultTypeID()
            End If
        End If

        Dim colFieldRefs As New FrameworkCollections.K1Dictionary(Of clsSearchElement)
        Dim blnOK As Boolean = True

        RecurseTestGroup(m_objFilterGroup, colFieldRefs, blnOK, False)

        If Not blnOK Then
            Return
        End If

        Dim objTableMask As New clsTableMask(objTable, clsTableMask.enumMaskType.SEARCH, intTypeID:=intTypeID)

        Try
            For Each objSE As clsSearchElement In colFieldRefs.Values
                AddToTableMask(objSE, objTableMask, blnOK)
                If Not blnOK Then Return
            Next

            m_objTableMask = objTableMask
        Catch ex As Exception
            If Not m_objTableMask Is Nothing Then
                m_objTableMask.Dispose()
            End If
            m_objTableMask = Nothing
        End Try
    End Sub

    Private Sub RecurseTestGroup(ByVal objSO As clsSearchObjBase, _
    ByRef colFieldRefs As FrameworkCollections.K1Dictionary(Of clsSearchElement), _
    ByRef blnOK As Boolean, ByRef blnNot As Boolean)
        If Not blnOK Then
            Return
        End If

        Dim eOpType As enumOperatorType = objSO.OperatorType

        If blnNot Then
            Select Case eOpType
                Case enumOperatorType.AND
                    eOpType = enumOperatorType.ORNOT
                Case enumOperatorType.ANDNOT
                    eOpType = enumOperatorType.AND
                Case enumOperatorType.NONE
                    eOpType = enumOperatorType.NOT
                Case enumOperatorType.NOT
                    eOpType = enumOperatorType.NONE
                Case enumOperatorType.OR
                    eOpType = enumOperatorType.ANDNOT
                Case enumOperatorType.ORNOT
                    eOpType = enumOperatorType.OR
            End Select
        End If

        If eOpType = enumOperatorType.OR OrElse _
        eOpType = enumOperatorType.ORNOT Then
            blnOK = False 'can't use or on metadata search
            Return
        End If

        If TypeOf objSO Is clsSearchGroup Then
            Dim objSG As clsSearchGroup = CType(objSO, clsSearchGroup)

            If objSG.OperatorType = enumOperatorType.NOT OrElse _
            objSG.OperatorType = enumOperatorType.ANDNOT OrElse _
            objSG.OperatorType = enumOperatorType.ORNOT Then
                blnNot = Not blnNot
            End If

            For Each objSeachObj As clsSearchObjBase In objSG.SearchObjs
                RecurseTestGroup(objSeachObj, colFieldRefs, blnOK, blnNot)
            Next
        Else
            Dim objSE As clsSearchElement = CType(objSO, clsSearchElement)

            Dim strKey As String

            '2016-07-21 -- Peter Melisi -- Bug fix for #1600003135
            Select Case objSE.CompareType
                Case enumComparisonType.EQUAL, enumComparisonType.GREATER_THAN_EQUAL, enumComparisonType.IN, enumComparisonType.EXACT
                    strKey = objSE.FieldRef

                Case enumComparisonType.LESS_THAN_EQUAL
                    strKey = objSE.FieldRef & "*"

                Case Else
                    blnOK = False
                    Return

            End Select

            If colFieldRefs(strKey) Is Nothing Then
                colFieldRefs.Add(strKey, New clsSearchElement(eOpType, _
                    objSE.FieldRef, objSE.CompareType, objSE.Value, objSE.Field))
            Else
                blnOK = False 'can only use a fieldref once in metatadata search
                Return
            End If
        End If
    End Sub

    Private Sub AddToTableMask(ByVal objSE As clsSearchElement, ByVal objParentTableMask As clsTableMask, _
    ByRef blnOk As Boolean)
        Dim arrRefs As String() = Split(objSE.FieldRef, ".")
        Dim objTable As clsTable = m_objDB.SysInfo.Tables(m_strRootTable)
        Dim objField As clsField
        Dim strField As String
        Dim objMaskField As clsMaskField
        Dim objTableMask As clsTableMask = objParentTableMask

        For intLoop As Integer = 1 To arrRefs.Length - 1
            strField = arrRefs(intLoop)

            If strField.Chars(0) = "*"c Then
                blnOk = False
                Return 'TODO: Implement for linked tables for K1?
            Else
                objField = m_objDB.SysInfo.Fields(objTable.ID & "_" & strField)

                objMaskField = objTableMask.MaskFieldCollection(objField.DatabaseName)

                If Not intLoop = arrRefs.Length - 1 AndAlso _
                Not (objMaskField.AllowFreeTextEntry AndAlso intLoop = arrRefs.Length - 2) Then
                    objTable = objField.FieldLink.IdentityTable

                    If Not objMaskField.CheckState = Windows.Forms.CheckState.Unchecked Then
                        blnOk = False
                        Return
                    End If

                    If objSE.OperatorType = enumOperatorType.NOT OrElse _
                    objSE.OperatorType = enumOperatorType.ANDNOT Then
                        objMaskField.CheckState = Windows.Forms.CheckState.Indeterminate
                    Else
                        objMaskField.CheckState = Windows.Forms.CheckState.Checked
                    End If

                    objTableMask = New clsTableMask(objTable, clsTableMask.enumMaskType.SEARCH, _
                        objParentMask:=objMaskField.Value1.MaskField)
                    objMaskField.Value1.MaskField.TableMask = objTableMask
                Else
                    If objMaskField.AllowFreeTextEntry AndAlso intLoop = arrRefs.Length - 2 Then
                        intLoop += 1
                    End If

                    If Not objMaskField.TableMask Is Nothing Then
                        blnOk = False
                        Return
                    End If

                    If objSE.OperatorType = enumOperatorType.NOT OrElse _
                    objSE.OperatorType = enumOperatorType.ANDNOT Then
                        'See if there is a conflicting value for the range search
                        If objMaskField.CheckState = Windows.Forms.CheckState.Checked Then
                            blnOk = False
                            Return
                        End If

                        objMaskField.CheckState = Windows.Forms.CheckState.Indeterminate
                    Else
                        'See if there is a conflicting value for the range search
                        If objMaskField.CheckState = Windows.Forms.CheckState.Indeterminate Then
                            blnOk = False
                            Return
                        End If

                        objMaskField.CheckState = Windows.Forms.CheckState.Checked
                    End If

                    If objSE.CompareType = enumComparisonType.GREATER_THAN_EQUAL OrElse _
                    objSE.CompareType = enumComparisonType.LESS_THAN_EQUAL OrElse _
                    objSE.CompareType = enumComparisonType.IN Then
                        objMaskField.MaskSearchType = clsMaskBase.enumMaskSearchType.RANGE
                    End If

                    If objSE.CompareType = enumComparisonType.LESS_THAN_EQUAL OrElse _
                    objSE.CompareType = enumComparisonType.IN Then
                        If objMaskField.AllowFreeTextEntry Then
                            objMaskField.Value2.FreeText = CType(objSE.Value, String)
                        Else
                            objMaskField.Value2.Value = objSE.Value
                        End If
                    Else
                        If objMaskField.AllowFreeTextEntry Then
                            objMaskField.Value1.FreeText = CType(objSE.Value, String)
                        Else
                            objMaskField.Value1.Value = objSE.Value
                        End If

                        If objField.IsForeignKey Then
                            clsMaskField.LoadLinkedData(objMaskField, objField.FieldLink.IdentityTable)
                        End If
                    End If
                End If
            End If
        Next
    End Sub
#End Region

#Region " Interpret Constant Methods "

    Private Function GetConstantValue(ByVal strValue As String) As Object
        strValue = strValue.ToUpper

        Dim intIndex As Integer = 0
        Dim arrEndChars As Char() = {"-"c, "+"c}
        Dim intEndIndex As Integer = strValue.IndexOfAny(arrEndChars, intIndex)
        Dim objReturn As Object = Nothing
        Dim strToken As String

        If intEndIndex > -1 Then
            strToken = Trim(strValue.Substring(intIndex, intEndIndex))
        Else
            strToken = Trim(strValue)
        End If

        Select Case strToken
            Case "CURRENT_DATE"
                objReturn = Today
            Case "CURRENT_DATE_AND_TIME"
                objReturn = Now
            Case "START_OF_MONTH"
                Dim dtToday As Date = Today
                objReturn = New Date(dtToday.Year, dtToday.Month, 1)
            Case "START_OF_YEAR"
                Dim dtToday As Date = Today
                objReturn = New Date(dtToday.Year, 1, 1)
            Case "END_OF_MONTH"
                Dim dtToday As Date = Today
                objReturn = New Date(dtToday.Year, dtToday.Month, 1)
                objReturn = CDate(objReturn).AddMonths(1).AddMilliseconds(-1)
            Case "END_OF_YEAR"
                Dim dtToday As Date = Today
                objReturn = New Date(dtToday.Year, 1, 1)
                objReturn = CDate(objReturn).AddYears(1).AddMilliseconds(-1)
            Case Else
                Throw New Exception("Unrecognized constant: " & strToken & vbCrLf & _
                    "Expected: CURRENT_DATE, CURRENT_DATE_AND_TIME, START_OF_MONTH, START_OF_YEAR, END_OF_MONTH, END_OF_YEAR")
        End Select

        While intEndIndex > -1
            Dim chChar As Char = strValue.Chars(intEndIndex)

            intIndex = intEndIndex + 1
            intEndIndex = strValue.IndexOfAny(arrEndChars, intIndex)

            If intEndIndex > -1 Then
                strToken = Trim(strValue.Substring(intIndex, intEndIndex - intIndex))
            Else
                strToken = Trim(strValue.Substring(intIndex, Len(strValue) - intIndex))
            End If

            Dim arrTokens As String() = strToken.Split(" "c)
            If arrTokens.Length = 2 Then
                Dim intNum As Integer
                Dim strNum As String = arrTokens(0).Trim
                Dim blnIsNumeric As Boolean = Integer.TryParse(strNum, intNum)
                If Not blnIsNumeric Then
                    Throw New Exception("Invalid number used in expression: " & strToken & vbCrLf & _
                        "Expected: [Number] [Period] (ex. 1 DAY).  Exceptable [Period] values: SECOND, MINUTE, HOUR, DAY, WEEK, MONTH, YEAR")
                End If

                If chChar = "-"c Then
                    intNum = 0 - intNum
                End If

                strToken = arrTokens(1).Trim
                Select Case strToken
                    Case "SECOND"
                        objReturn = CDate(objReturn).AddSeconds(intNum)
                    Case "MINUTE"
                        objReturn = CDate(objReturn).AddMinutes(intNum)
                    Case "HOUR"
                        objReturn = CDate(objReturn).AddHours(intNum)
                    Case "DAY"
                        objReturn = CDate(objReturn).AddDays(intNum)
                    Case "WEEK"
                        objReturn = CDate(objReturn).AddDays((intNum * 7))
                    Case "MONTH"
                        objReturn = CDate(objReturn).AddMonths(intNum)
                    Case "YEAR"
                        objReturn = CDate(objReturn).AddYears(intNum)
                    Case Else
                        Throw New Exception("Unrecognized period used: " & strToken & vbCrLf & _
                            "Expected: [Number] [Period] (ex. 1 DAY).  Exceptable [Period] values: SECOND, MINUTE, HOUR, DAY, WEEK, MONTH, YEAR")
                End Select
            Else
                Throw New Exception("Unrecognized constant: " & strToken & vbCrLf & _
                "Expected: [Number] [Period] (ex. 1 DAY).  Exceptable [Period] values: SECOND, MINUTE, HOUR, DAY, WEEK, MONTH, YEAR")
            End If
        End While

        Return objReturn
    End Function
#End Region

#End Region

#Region " IDisposable Support "

    Protected Overridable Sub Dispose(ByVal blnDisposing As Boolean)
        If Not m_blnDisposedValue Then
            If blnDisposing Then
                m_objDB = Nothing
                m_objTableMask = Nothing

                If m_colVariables IsNot Nothing Then
                    m_colVariables.Clear()
                    m_colVariables = Nothing
                End If

                m_objFilterGroup = Nothing
            End If
        End If
        m_blnDisposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
