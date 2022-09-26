#Region " File Information "

'=====================================================================
' This class represents the table AutoNumberFormat in the Database.
' It is used to automatically assign values using a specified format
' to certain mask fields.
'=====================================================================

#Region " Revision History "

'=====================================================================
' Name      Date         Description
'---------------------------------------------------------------------
' M.O.      03/02/2005    Implemented.
'=====================================================================
#End Region

#End Region
Imports K1Library.FrameworkCollections

Public Class clsAutoNumberFormat
    Inherits clsDBObjBase

#Region " Members "

    Private m_strFormat As String
    Private m_intSequentialNumberLength As Integer
    Private m_blnSequentialNumberPad As Boolean
    Private m_strMultipleSequenceMask As String
    Private m_colMultipleSequences As FrameworkCollections.K1Dictionary(Of clsMultipleSequence)
    Private m_intMaskCount As Integer
    Private m_intLastValue As Integer
    Private m_colLocks As New K1Library.FrameworkCollections.K1Dictionary(Of clsRecordLock)
    Private m_colTokens As FrameworkCollections.K1Collection(Of clsAutoNumberToken)
    Private m_strFormattedMask As String
    Private m_objAppliesToField As clsField
    Private m_intAppliesToTypeID As Integer = clsDBConstants.cintNULL
#End Region

#Region " Constants "

    Public Const cSEQUENCE_MASK As String = "(Sequence)"

    Public Class ErrorMessage
        ' TODO: these should go into the error messages table
        Public Const cDEFAULT As String = "[FIELD] is an auto-number format field but has " & _
            "an invalid value. The field should be in the format [FORMAT]."

        Public Const cINCOMPLETE As String = "[FIELD] is an auto-number format field but it " & _
            "has not been completed. The field should be in the format [FORMAT]."

        Public Const cNOMULTIPLESEQUENCES As String = "[FIELD] is an auto-number format field that " & _
            "requires a multiple sequence but none have been defined. Please " & _
            "contact your RecFind Administrator. The field should be in the format [FORMAT]."

        Public Const cINVALIDMULTIPLESEQUENCES As String = "[FIELD] is an auto-number format field " & _
            "that requires a multiple sequence but the one you have entered has not been " & _
            "defined. Please contact your RecFind Administrator. The field should be in the format [FORMAT]."

        Public Const cINVALIDAUTONUMBERFORMAT As String = "The auto-number format for " & _
            "[FIELD] is invalid.  The multiple sequence mask cannot be applied to a " & _
            "variable length format (D, M, S).  Please contact your RecFind Administrator."

        Public Const cMISSINGTITLE As String = "The auto-number format for [FIELD] " & _
            "depends on a value in [FIELD2].  There is currently no value " & _
            "in [FIELD2]. This could be the result of a configuration issue. " & _
            "Please contact your RecFind Administrator for further assistance."
    End Class
#End Region

#Region " Constructors "

    Public Sub New()
        MyBase.New()
        m_strFormat = clsDBConstants.cstrNULL
        m_intSequentialNumberLength = clsDBConstants.cintNULL
        m_blnSequentialNumberPad = False
        m_strMultipleSequenceMask = clsDBConstants.cstrNULL
        m_intLastValue = clsDBConstants.cintNULL
    End Sub

    Private Sub New(ByVal objDR As DataRow, ByVal objDB As clsDB)

        MyBase.new(objDR, objDB)

        m_strFormat = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormat.cFORMAT, clsDBConstants.cstrNULL), String)
        m_intSequentialNumberLength = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormat.cSEQUENTIALNUMBERLENGTH, clsDBConstants.cintNULL), Integer)
        m_blnSequentialNumberPad = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormat.cSEQUENTIALNUMBERPAD, False), Boolean)
        m_strMultipleSequenceMask = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormat.cMULTIPLESEQUENCEMASK, clsDBConstants.cstrNULL), String)
        m_intLastValue = CType(clsDB_Direct.DataRowValue(objDR, clsDBConstants.Fields.AutoNumberFormat.cSEQUENTIALNUMBERLASTVALUE, 0), Integer)

        '2013-12-10 -- Naing Thein -- Fix for #1300002569
        If (m_colTokens Is Nothing) Then
            m_colTokens = TokenizeString(m_strFormat, m_objDB)
        End If

    End Sub

#End Region

#Region " Properties "

    Public ReadOnly Property Format() As String
        Get
            Return m_strFormat
        End Get
    End Property

    Public ReadOnly Property SequentialNumberLength() As Integer
        Get
            Return m_intSequentialNumberLength
        End Get
    End Property

    Public ReadOnly Property SequentialNumberPad() As Boolean
        Get
            Return m_blnSequentialNumberPad
        End Get
    End Property

    Public ReadOnly Property MultipleSequenceMask() As String
        Get
            Return m_strMultipleSequenceMask
        End Get
    End Property

    Public ReadOnly Property MultipleSequences() As FrameworkCollections.K1Dictionary(Of clsMultipleSequence)
        Get
            If m_colMultipleSequences Is Nothing Then
                m_colMultipleSequences = clsMultipleSequence.GetList(Me, Me.Database)
            End If
            Return m_colMultipleSequences
        End Get
    End Property

    Public ReadOnly Property IsMultipleSequenceAutoNumber() As Boolean
        Get
            If String.IsNullOrEmpty(m_strMultipleSequenceMask) OrElse _
            m_strMultipleSequenceMask.IndexOf("*"c) = -1 OrElse _
            Not m_strMultipleSequenceMask.Length <= FormattedMask.Length Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property AutoNumberTokens() As K1Collection(Of clsAutoNumberToken)
        Get
            If m_colTokens Is Nothing Then
                m_colTokens = TokenizeString(m_strFormat, m_objDB)
            Else
                '[Naing] Bug Fix: 1300002493 
                'Must refresh or the tokens will be out dated if the Date changes but the global cache is not yet updated.
                'Global cache is only updated after the Application is restarted or IIS app pool is restarted.
                Dim blnRefresh = m_colTokens.Any(Function(token)
                                                     Return token.TokenType = clsAutoNumberToken.enumTokenType.YEAR_4 OrElse
                                                         token.TokenType = clsAutoNumberToken.enumTokenType.YEAR_2 OrElse
                                                         token.TokenType = clsAutoNumberToken.enumTokenType.MONTH OrElse
                                                         token.TokenType = clsAutoNumberToken.enumTokenType.MONTH_PAD OrElse
                                                         token.TokenType = clsAutoNumberToken.enumTokenType.DAY OrElse
                                                         token.TokenType = clsAutoNumberToken.enumTokenType.DAY_PAD
                                                 End Function)
                If (blnRefresh) Then
                    m_colTokens = TokenizeString(m_strFormat, m_objDB)
                End If

            End If

            Return m_colTokens
        End Get
    End Property

    Public ReadOnly Property MaskCount() As Integer
        Get
            If m_intMaskCount = 0 Then
                For Each objToken As clsAutoNumberToken In m_colTokens ' AutoNumberTokens 2013-11-08 -- Peter Melisi -- Fix for #1300002521
                    If objToken.TokenType = clsAutoNumberToken.enumTokenType.ANY OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.ANY_NOSPACE OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.LETTER OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.NUMBER Then
                        m_intMaskCount += 1
                    End If
                Next
            End If

            Return m_intMaskCount
        End Get
    End Property

    Public ReadOnly Property RequiresUserInput() As Boolean
        Get
            Return (MaskCount > 0)
        End Get
    End Property

    Public ReadOnly Property LastValue() As Integer
        Get
            Return m_intLastValue
        End Get
    End Property

    Public ReadOnly Property HasSequence() As Boolean
        Get
            For Each objToken As clsAutoNumberToken In m_colTokens ' AutoNumberTokens 2013-11-08 -- Peter Melisi -- Fix for #1300002521
                If objToken.TokenType = clsAutoNumberToken.enumTokenType.SEQUENCE Then
                    Return True
                    Exit For
                End If
            Next
            Return False
        End Get
    End Property

    Public ReadOnly Property FormattedMask() As String
        Get
            If m_strFormattedMask Is Nothing Then
                m_strFormattedMask = CreateMaskMultiSequenceString()
            End If
            Return m_strFormattedMask
        End Get
    End Property

    Public ReadOnly Property AppliesToField() As clsField
        Get
            If m_objAppliesToField Is Nothing Then
                GetAppliesToFieldAndType()
            End If
            Return m_objAppliesToField
        End Get
    End Property

    Public ReadOnly Property AppliesToTypeID() As Integer
        Get
            If m_objAppliesToField Is Nothing Then
                GetAppliesToFieldAndType()
            End If
            Return m_intAppliesToTypeID
        End Get
    End Property

#End Region

#Region " GetItem, GetList "

    Public Shared Function GetItem(ByVal intID As Integer, ByVal objDB As clsDB) As clsAutoNumberFormat
        Try
            Dim objDT As DataTable = objDB.GetItem(clsDBConstants.Tables.cAUTONUMBERFORMAT, intID)

            Return New clsAutoNumberFormat(objDT.Rows(0), objDB)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetList(ByVal objDB As clsDB) As FrameworkCollections.K1Dictionary(Of clsAutoNumberFormat)
        Try
            Dim colObjects As New FrameworkCollections.K1Dictionary(Of clsAutoNumberFormat)
            Dim objDT As DataTable = objDB.GetDataTable( _
                clsDBConstants.Tables.cAUTONUMBERFORMAT & clsDBConstants.StoredProcedures.cGETLIST)

            For Each objDataRow As DataRow In objDT.Rows
                Dim objItem As New clsAutoNumberFormat(objDataRow, objDB)
                colObjects.Add(CType(objItem.ID, String), objItem)
            Next

            Return colObjects
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region " Business Logic "

#Region " TokenizeFormatString "

    '=====================================================================
    ' Parses the format string and produces an array of tokens
    '=====================================================================
    Public Shared Function TokenizeString(ByVal strFormat As String, Optional ByVal objDb As clsDB = Nothing) As K1Collection(Of clsAutoNumberToken)
        Dim chrChar As Char
        Dim intLoop As Integer
        Dim blnHasSequence As Boolean = False
        Dim blnIsConstant As Boolean = False
        Dim colTokens As New FrameworkCollections.K1Collection(Of clsAutoNumberToken)

        intLoop = 0

        Dim objNow = Now

        If (objDb IsNot Nothing) Then
            objNow = objDb.GetCurrentTime()
        End If

        Do While intLoop < strFormat.Length

            chrChar = strFormat.Chars(intLoop)

            If blnIsConstant Then
                If chrChar = """"c Then
                    blnIsConstant = False
                Else
                    colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))
                End If
            Else
                Select Case Char.ToUpper(chrChar)
                    Case "A"c
                        colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.LETTER))

                    Case "9"c
                        colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.NUMBER))

                    Case "*"c
                        colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.ANY))

                    Case "X"c
                        colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.ANY_NOSPACE))

                    Case "S"c
                        If Not blnHasSequence Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.SEQUENCE))
                            blnHasSequence = True
                        Else
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))
                        End If

                    Case "D"c
                        If strFormat.Substring(intLoop).ToUpper.IndexOf("DD", System.StringComparison.InvariantCultureIgnoreCase) = 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.DAY_PAD, objNow.Day.ToString("00")))
                            intLoop += 1
                        Else
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.DAY, objNow.Day.ToString))
                        End If

                    Case "M"c
                        If strFormat.Substring(intLoop).ToUpper.IndexOf("MM", System.StringComparison.InvariantCultureIgnoreCase) = 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.MONTH_PAD, objNow.Month.ToString("00")))
                            intLoop += 1
                        Else
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.MONTH, objNow.Month.ToString))
                        End If

                    Case "Y"c
                        If strFormat.Substring(intLoop).ToUpper.IndexOf("YYYY", System.StringComparison.InvariantCultureIgnoreCase) = 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.YEAR_4, objNow.Year.ToString))
                            intLoop += 3
                        ElseIf strFormat.Substring(intLoop).IndexOf("YY", System.StringComparison.InvariantCultureIgnoreCase) = 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.YEAR_2, objNow.Year.ToString.Substring(2)))
                            intLoop += 1
                        Else
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))
                        End If

                    Case """"c
                        Dim intPos As Integer = strFormat.Substring(intLoop + 1).IndexOf(""""c)

                        If intPos < 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))
                        ElseIf intPos = 0 Then
                            colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))
                            intLoop += 1
                        Else
                            blnIsConstant = True
                        End If

                    Case Else
                        colTokens.Add(New clsAutoNumberToken(clsAutoNumberToken.enumTokenType.CONSTANT, chrChar))

                End Select
            End If

            intLoop += 1
        Loop

        Return colTokens
    End Function

#End Region

#Region " CreateMask "

    '=====================================================================
    ' Converts the format string into tokens and returns the mask to
    ' display to the user.
    '=====================================================================
    Public Function CreateMask(ByVal colTokens As FrameworkCollections.K1Collection(Of clsAutoNumberToken)) As String
        Dim strMask As String = ">"

        For Each objToken As clsAutoNumberToken In colTokens
            Select Case objToken.TokenType
                Case clsAutoNumberToken.enumTokenType.LETTER
                    strMask &= "L"

                Case clsAutoNumberToken.enumTokenType.NUMBER
                    strMask &= "0"

                Case clsAutoNumberToken.enumTokenType.ANY
                    strMask &= "C"

                Case clsAutoNumberToken.enumTokenType.ANY_NOSPACE
                    strMask &= "&"

                Case clsAutoNumberToken.enumTokenType.SEQUENCE
                    strMask &= cSEQUENCE_MASK

                Case clsAutoNumberToken.enumTokenType.DAY
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.DAY_PAD
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.MONTH
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.MONTH_PAD
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.YEAR_2
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.YEAR_4
                    strMask &= objToken.Value.Replace("0", "\0").Replace("9", "\9")

                Case clsAutoNumberToken.enumTokenType.CONSTANT
                    strMask &= "\" & objToken.Value

            End Select
        Next

        Return strMask
    End Function

    Private Function CreateMaskMultiSequenceString() As String
        Dim strReturn As String = ""

        ' AutoNumberTokens 2013-11-08 -- Peter Melisi -- Fix for #1300002521
        For Each objToken As clsAutoNumberToken In m_colTokens
            Select Case objToken.TokenType
                Case clsAutoNumberToken.enumTokenType.ANY
                    strReturn &= "~"
                Case clsAutoNumberToken.enumTokenType.ANY_NOSPACE
                    strReturn &= "~"
                Case clsAutoNumberToken.enumTokenType.CONSTANT
                    strReturn &= objToken.Value
                Case clsAutoNumberToken.enumTokenType.DAY
                    strReturn &= "D"
                Case clsAutoNumberToken.enumTokenType.DAY_PAD
                    strReturn &= objToken.Value
                Case clsAutoNumberToken.enumTokenType.LETTER
                    strReturn &= "~"
                Case clsAutoNumberToken.enumTokenType.MONTH
                    strReturn &= "M"
                Case clsAutoNumberToken.enumTokenType.MONTH_PAD
                    strReturn &= objToken.Value
                Case clsAutoNumberToken.enumTokenType.NUMBER
                    strReturn &= "~"
                Case clsAutoNumberToken.enumTokenType.SEQUENCE
                    strReturn &= "S"
                Case clsAutoNumberToken.enumTokenType.YEAR_2
                    strReturn &= objToken.Value
                Case clsAutoNumberToken.enumTokenType.YEAR_4
                    strReturn &= objToken.Value
            End Select
        Next

        Return strReturn
    End Function

    Public Function GetMaskValue(ByVal strMultipleSequence As String, _
    Optional ByVal strCurrentValue As String = Nothing) As String
        Dim strMSMask As String = FormattedMask
        Dim arrChars() As Char = m_strMultipleSequenceMask.ToArray
        Dim strReturn As String = ""

        Dim intSeqCounter As Integer = 0
        For intChar As Integer = 0 To arrChars.Length - 1
            If arrChars(intChar) = "*" Then
                If strMSMask(intChar) = "~" Then
                    strReturn &= strMultipleSequence(intSeqCounter)
                End If
                intSeqCounter += 1
            Else
                'if the character as this location is user input, make it a space
                If intChar < strMSMask.Length AndAlso strMSMask(intChar) = "~"c Then
                    strReturn &= " "
                End If
            End If
        Next

        If strReturn.IndexOf(" "c) >= 0 AndAlso Not String.IsNullOrEmpty(strCurrentValue) Then
            arrChars = strReturn.ToCharArray
            Dim strNewValue As String = ""
            Dim arrCurrentVal As Char() = strCurrentValue.ToCharArray

            For intLoop As Integer = 0 To arrChars.Length - 1
                If arrChars(intLoop) = " "c AndAlso _
                arrCurrentVal.Length >= intLoop + 1 Then
                    strNewValue &= arrCurrentVal(intLoop)
                Else
                    strNewValue &= arrChars(intLoop)
                End If
            Next

            strReturn = strNewValue
        End If

        Return strReturn
    End Function

    Public Function GetMultipleSequenceValue(ByVal strMask As String) As String
        Dim strMSMask As String = FormattedMask
        Dim arrChars() As Char = m_strMultipleSequenceMask.ToArray
        Dim strReturn As String = ""

        If strMask Is Nothing Then
            strMask = ""
        End If

        Dim intSeqCounter As Integer = 0
        For intChar As Integer = 0 To arrChars.Length - 1
            If arrChars(intChar) = "*" Then
                If strMSMask(intChar) = "~" Then
                    If intSeqCounter >= strMask.Length Then
                        Exit For
                    End If

                    strReturn &= strMask(intSeqCounter)
                    intSeqCounter += 1
                Else
                    strReturn &= strMSMask(intChar)
                End If
            Else
                If strMSMask(intChar) = "~" Then
                    intSeqCounter += 1
                End If
            End If
        Next

        Return strReturn
    End Function
#End Region

#Region " Validation "

    '=====================================================================
    ' Validate the auto-number format
    '=====================================================================
    Public Function Validate(ByVal objMaskValue As clsMaskFieldValue) As clsAutoNumberValidation
        Dim objValidation As New clsAutoNumberValidation
        Dim strField As String = String.Empty

        Try
            strField = objMaskValue.MaskField.Caption
            If String.IsNullOrEmpty(strField) Then
                strField = objMaskValue.MaskField.Field.DatabaseName
            End If

            If IsMultipleSequenceAutoNumber Then
                For Each objMask As clsMaskField In objMaskValue.MaskField.MaskFieldCollection.Values
                    If objMask.DeterminesMultipleSequence AndAlso _
                    objMask.Value1.Value Is Nothing Then
                        Dim strTitle As String = objMask.Caption
                        If String.IsNullOrEmpty(strTitle) Then
                            strTitle = objMask.Field.DatabaseName
                        End If

                        objValidation.SetError(ErrorMessage.cMISSINGTITLE.Replace("[FIELD]", _
                            strField).Replace("[FIELD2]", strTitle), clsAutoNumberValidation.enumAutonumberErrorType.OTHER)
                        Return objValidation
                    End If
                Next
            End If

            'find out if we expect user input, and how many characters to expect if we do
            objValidation.HasSequence = HasSequence

            m_intMaskCount = MaskCount

            If m_intMaskCount > 0 Then
                'Make sure we have a value
                If objMaskValue.Value Is Nothing Then
                    objValidation.SetError(ErrorMessage.cINCOMPLETE.Replace("[FIELD]", _
                        strField).Replace("[FORMAT]", m_strFormat), clsAutoNumberValidation.enumAutonumberErrorType.MASK_INCOMPLETE)
                    Return objValidation
                End If

                'check that the length of the value is the number of input characters expected
                Dim strValue As String = CStr(objMaskValue.Value)

                If strValue.Length < m_intMaskCount Then
                    strValue = strValue.PadRight(m_intMaskCount, " "c)
                End If

                'Assign the entered values to the token collection
                Dim intValueIndex As Integer = 0
                For Each objToken As clsAutoNumberToken In m_colTokens ' AutoNumberTokens 2013-11-08 -- Peter Melisi -- Fix for #1300002521
                    If objToken.TokenType = clsAutoNumberToken.enumTokenType.ANY OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.ANY_NOSPACE OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.LETTER OrElse _
                    objToken.TokenType = clsAutoNumberToken.enumTokenType.NUMBER Then
                        objToken.Value = strValue.Chars(intValueIndex)

                        If Not objToken.TokenType = clsAutoNumberToken.enumTokenType.ANY AndAlso _
                        objToken.Value = " "c Then
                            objValidation.SetError(ErrorMessage.cINCOMPLETE.Replace("[FIELD]", _
                                strField).Replace("[FORMAT]", m_strFormat), clsAutoNumberValidation.enumAutonumberErrorType.MASK_INCOMPLETE)
                            Return objValidation
                        End If

                        intValueIndex += 1
                    End If
                Next
            End If

            'Check if it has Multiple Sequences
            If Not IsMultipleSequenceAutoNumber Then
                Return objValidation
            End If

            Dim intTokenIndex As Integer = 0
            Dim strMultipleSequence As String = ""
            Dim intCount As Integer = 0

            For intMaskIndex As Integer = 0 To m_strMultipleSequenceMask.Length - 1
                Dim objToken As clsAutoNumberToken = m_colTokens(intTokenIndex) ' AutoNumberTokens(intTokenIndex) 2013-11-08 -- Peter Melisi -- Fix for #1300002521

                Select Case objToken.TokenType
                    Case clsAutoNumberToken.enumTokenType.CONSTANT, _
                    clsAutoNumberToken.enumTokenType.ANY, _
                    clsAutoNumberToken.enumTokenType.ANY_NOSPACE, _
                    clsAutoNumberToken.enumTokenType.LETTER, _
                    clsAutoNumberToken.enumTokenType.NUMBER, _
                    clsAutoNumberToken.enumTokenType.SEQUENCE, _
                    clsAutoNumberToken.enumTokenType.DAY, _
                    clsAutoNumberToken.enumTokenType.MONTH
                        intTokenIndex += 1

                    Case clsAutoNumberToken.enumTokenType.DAY_PAD, _
                    clsAutoNumberToken.enumTokenType.MONTH_PAD, _
                    clsAutoNumberToken.enumTokenType.YEAR_2
                        If intCount = 0 Then
                            intCount = 1
                        Else
                            intCount = 0
                            intTokenIndex += 1
                        End If

                    Case clsAutoNumberToken.enumTokenType.YEAR_4
                        If intCount = 3 Then
                            intCount = 0
                            intTokenIndex += 1
                        Else
                            intCount += 1
                        End If
                End Select

                If m_strMultipleSequenceMask.Substring(intMaskIndex, 1) = "*" Then
                    Select Case objToken.TokenType
                        Case clsAutoNumberToken.enumTokenType.CONSTANT, _
                        clsAutoNumberToken.enumTokenType.ANY, _
                        clsAutoNumberToken.enumTokenType.ANY_NOSPACE, _
                        clsAutoNumberToken.enumTokenType.LETTER, _
                        clsAutoNumberToken.enumTokenType.NUMBER
                            strMultipleSequence &= objToken.Value

                        Case clsAutoNumberToken.enumTokenType.SEQUENCE, _
                        clsAutoNumberToken.enumTokenType.DAY, _
                        clsAutoNumberToken.enumTokenType.MONTH
                            objValidation.SetError(ErrorMessage.cINVALIDAUTONUMBERFORMAT.Replace( _
                                "[FIELD]", strField))
                            Return objValidation

                        Case clsAutoNumberToken.enumTokenType.DAY_PAD, _
                        clsAutoNumberToken.enumTokenType.MONTH_PAD, _
                        clsAutoNumberToken.enumTokenType.YEAR_2
                            If intCount = 0 Then
                                strMultipleSequence &= objToken.Value.Chars(1)
                            Else
                                strMultipleSequence &= objToken.Value.Chars(0)
                            End If

                        Case clsAutoNumberToken.enumTokenType.YEAR_4
                            If intCount = 0 Then
                                strMultipleSequence &= objToken.Value.Chars(3)
                            Else
                                strMultipleSequence &= objToken.Value.Chars(intCount - 1)
                            End If

                    End Select
                End If
            Next

            objValidation.MultipleSequenceMask = strMultipleSequence
            objValidation.IsAutoGenerated = IsAutoGeneratedMultiSequence(objValidation)

            If objValidation.IsAutoGenerated Then
                Return objValidation
            End If

            If GetMultipleSequence(strMultipleSequence) Is Nothing Then
                objValidation.SetError(ErrorMessage.cINVALIDMULTIPLESEQUENCES.Replace("[FIELD]", _
                        strField).Replace("[FORMAT]", m_strFormat))
                Return objValidation
            End If
        Catch ex As Exception
            objValidation.SetError(ErrorMessage.cDEFAULT.Replace("[FIELD]", _
                strField).Replace("[FORMAT]", m_strFormat))
        End Try

        Return objValidation
    End Function
#End Region

#Region " IsAutoGeneratedMultiSequence "

    Private Function IsAutoGeneratedMultiSequence(ByRef objValidation As clsAutoNumberValidation) As Boolean

        Dim strMultipleSequenceFormat As String = ""

        For intIndex As Integer = 0 To m_strMultipleSequenceMask.Length - 1
            If m_strMultipleSequenceMask.Substring(intIndex, 1) = "*" Then
                strMultipleSequenceFormat &= m_strFormat.Chars(intIndex)
            End If
        Next

        Dim colTokens As K1Collection(Of clsAutoNumberToken)
        colTokens = TokenizeString(strMultipleSequenceFormat, m_objDB)

        '[Naing] All the tokens for the multi sequence must satisfy the condition that they are auto generated!
        Return colTokens.All(Function(objToken) (objToken.TokenType = clsAutoNumberToken.enumTokenType.CONSTANT OrElse
                                                 objToken.TokenType = clsAutoNumberToken.enumTokenType.DAY_PAD OrElse
                                                 objToken.TokenType = clsAutoNumberToken.enumTokenType.MONTH_PAD OrElse
                                                 objToken.TokenType = clsAutoNumberToken.enumTokenType.YEAR_2 OrElse
                                                 objToken.TokenType = clsAutoNumberToken.enumTokenType.YEAR_4))

    End Function
#End Region

#Region " Generate "

    '=====================================================================
    ' Generate the final output for this auto-number format
    '=====================================================================
    Public Sub Generate(ByVal objMaskValue As clsMaskFieldValue, Optional ByVal blnOverride As Boolean = False)
        If objMaskValue.AutoNumberGenerated AndAlso Not blnOverride Then
            Return
        End If

        Dim objValidation As clsAutoNumberValidation = Validate(objMaskValue)
        If Not objValidation.ErrorType = clsAutoNumberValidation.enumAutonumberErrorType.NONE Then
            Throw New clsK1Exception(objValidation.ErrorMsg)
        End If

        Dim intSequence As Integer

        If objValidation.HasSequence Then
            If objMaskValue.SequentialNumber = clsDBConstants.cintNULL Then
                Dim objTable As clsTable

                If String.IsNullOrEmpty(m_strMultipleSequenceMask) Then
                    'Use the Autonumber Format's sequence
                    objTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cAUTONUMBERFORMAT)
                    intSequence = GetNextSequenceNumber(objTable, m_intID, _
                        clsDBConstants.Fields.AutoNumberFormat.cSEQUENTIALNUMBERLASTVALUE)
                Else
                    Dim objMS As clsMultipleSequence = GetMultipleSequence(objValidation.MultipleSequenceMask)

                    If objMS Is Nothing AndAlso _
                    objValidation.IsAutoGenerated Then
                        'Create the MultipleSequence
                        objMS = CreateMultipleSequence(objValidation.MultipleSequenceMask)
                    End If

                    'Use the Multiple Sequence
                    objTable = m_objDB.SysInfo.Tables(clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE)
                    intSequence = GetNextSequenceNumber(objTable, objMS.ID, _
                        clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cMULTIPLESEQUENCELASTVALUE)
                End If

                objMaskValue.SequentialNumber = intSequence
            Else
                intSequence = objMaskValue.SequentialNumber
            End If
        End If

        Dim strValue As String = ""
        For Each objToken As clsAutoNumberToken In m_colTokens ' AutoNumberTokens 2013-11-08 -- Peter Melisi -- Fix for #1300002521
            If objToken.TokenType = clsAutoNumberToken.enumTokenType.SEQUENCE Then
                If m_blnSequentialNumberPad Then
                    strValue &= intSequence.ToString("".PadLeft(m_intSequentialNumberLength, "0"c))
                Else
                    strValue &= intSequence
                End If
            Else
                strValue &= objToken.Value
            End If
        Next

        objMaskValue.RollbackValue = objMaskValue.Value
        objMaskValue.InitializeValue(strValue)
        objMaskValue.AutoNumberGenerated = True
    End Sub

    Public Function GetMultipleSequence(ByVal strExternalID As String) As clsMultipleSequence
        Dim objMS As clsMultipleSequence = Nothing

        If Me.MultipleSequences(strExternalID) Is Nothing Then
            Dim colParams As New clsDBParameterDictionary

            If strExternalID IsNot Nothing Then
                colParams.Add(New clsDBParameter(clsDB.ParamName( _
                    clsDBConstants.Fields.cEXTERNALID), strExternalID))
            End If

            colParams.Add(New clsDBParameter(clsDB.ParamName( _
                clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cAUTONUMBERFORMATID), m_intID))

            Dim objDT As DataTable = m_objDB.GetDataTable( _
                clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE & _
                clsDBConstants.StoredProcedures.cGETLIST, colParams)

            If Not objDT Is Nothing AndAlso objDT.Rows.Count = 1 Then
                objMS = clsMultipleSequence.GetItem( _
                    CInt(objDT.Rows(0).Item(clsDBConstants.Fields.cID)), m_objDB)

                m_colMultipleSequences.Add(strExternalID, objMS)
            End If
        Else
            objMS = m_colMultipleSequences(strExternalID)
        End If

        Return objMS
    End Function

    '=====================================================================
    ' Generate the next sequence number
    '=====================================================================
    Private Function GetNextSequenceNumber(ByVal objTable As clsTable, ByVal intID As Integer, _
    ByVal strField As String) As Integer
        Dim intLoop As Integer = 0
        Dim blnNotUpdated As Boolean = True
        Dim strUser As String = String.Empty
        Dim intSequence As Integer

        'try 10 times to get an exclusive lock to update the sequence
        While intLoop < 10 And blnNotUpdated
            strUser = clsRecordLock.GetLock(m_objDB, m_colLocks, objTable.ID, intID)

            If (String.IsNullOrEmpty(strUser)) Then
                Try
                    Dim colParams As New clsDBParameterDictionary
                    colParams.Add(New clsDBParameter("@ID", intID))

                    Dim strSQL As String = "SELECT [" & strField & "] FROM [" & _
                        objTable.DatabaseName & "] WHERE [" & clsDBConstants.Fields.cID & "] = @ID"

                    Dim objDT As DataTable = m_objDB.GetDataTableBySQL(strSQL, colParams)
                    colParams.Dispose()

                    If objDT Is Nothing OrElse Not objDT.Rows.Count = 1 Then
                        intSequence = 1
                    Else
                        intSequence = CInt(clsDB.NullValue(objDT.Rows(0)(0), 0)) + 1
                    End If

                    Dim colMasks As clsMaskFieldDictionary = _
                        clsMaskField.CreateMaskCollection(objTable, intID)

                    colMasks.UpdateMaskObj(strField, intSequence)

                    colMasks.Update(objTable.Database)
                    blnNotUpdated = False
                Catch ex As Exception
                    Throw
                Finally
                    clsRecordLock.ReleaseLock(m_objDB, m_colLocks, objTable.ID, intID)
                End Try
            Else
                Threading.Thread.Sleep(100)
                intLoop += 1
            End If
        End While

        If intLoop = 10 And blnNotUpdated Then
            Throw New clsK1Exception("The Auto-number Format record is exclusively locked to user " & _
                strUser & " and the sequence cannot get updated.")
        End If

        Return intSequence
    End Function

    Private Function CreateMultipleSequence(ByVal strExternalID As String) As clsMultipleSequence
        Dim objTable As clsTable = m_objDB.SysInfo.Tables( _
            clsDBConstants.Tables.cAUTONUMBERFORMATMULTIPLESEQUENCE)

        Dim colMasks As clsMaskFieldDictionary = clsMaskField.CreateMaskCollection(objTable)

        colMasks.UpdateMaskObj(clsDBConstants.Fields.cEXTERNALID, strExternalID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.cSECURITYID, m_intSecurityID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cAUTONUMBERFORMATID, m_intID)
        colMasks.UpdateMaskObj(clsDBConstants.Fields.AutoNumberFormatMultipleSequence.cMULTIPLESEQUENCELASTVALUE, 0)

        Dim objMS As clsMultipleSequence

        Try
            Dim intID As Integer = colMasks.Insert(m_objDB)
        Catch ex As clsK1Exception
            If Not ex.ErrorNumber = clsDB_Direct.enumSQLExceptions.INDEX_VIOLATION Then
                Throw
            End If
        End Try

        objMS = GetMultipleSequence(strExternalID)

        Return objMS
    End Function
#End Region

#Region " Get Associated Field And Type "

    Private Sub GetAppliesToFieldAndType()
        Dim strSQL As String
        Dim objDT As DataTable
        Dim intFieldID As Integer

        m_intAppliesToTypeID = clsDBConstants.cintNULL

        strSQL = "SELECT [" & clsDBConstants.Fields.cID & "] " & _
            "FROM [" & clsDBConstants.Tables.cFIELD & "] WHERE " & _
            "[" & clsDBConstants.Fields.Field.cAUTONUMBERFORMATID & "] = @ID"

        Dim colParams As New clsDBParameterDictionary
        colParams.Add(New clsDBParameter("@ID", m_intID))

        objDT = m_objDB.GetDataTableBySQL(strSQL, colParams)
        If objDT IsNot Nothing AndAlso objDT.Rows.Count > 0 Then
            intFieldID = CInt(objDT.Rows(0)(0))

            m_objAppliesToField = m_objDB.SysInfo.Fields(intFieldID)

            Return
        End If

        strSQL = "SELECT [" & clsDBConstants.Fields.TypeFieldInfo.cFIELDID & "], " & _
            "[" & clsDBConstants.Fields.TypeFieldInfo.cAPPLIESTOTYPEID & "] " & _
            "FROM [" & clsDBConstants.Tables.cTYPEFIELDINFO & "] WHERE " & _
            "[" & clsDBConstants.Fields.TypeFieldInfo.cAUTONUMBERFORMATID & "] = @ID"

        objDT = m_objDB.GetDataTableBySQL(strSQL, colParams)
        If objDT IsNot Nothing AndAlso objDT.Rows.Count > 0 Then
            intFieldID = CInt(objDT.Rows(0)(0))
            m_intAppliesToTypeID = CInt(objDT.Rows(0)(1))

            m_objAppliesToField = m_objDB.SysInfo.Fields(intFieldID)
        End If
    End Sub
#End Region

    Public Sub RollBack(ByVal objMaskValue As clsMaskFieldValue)
        If objMaskValue.RollbackValue IsNot Nothing Then
            objMaskValue.InitializeValue(objMaskValue.RollbackValue)
        End If
        objMaskValue.RollbackValue = Nothing
        objMaskValue.AutoNumberGenerated = False
        objMaskValue.SequentialNumber = clsDBConstants.cintNULL
    End Sub

    Public Sub FormatMaskField(ByVal objMaskValue As clsMaskFieldValue)
        Try
            If objMaskValue.Value Is Nothing Then
                Return
            End If

            Dim strMaskValue As String = CStr(objMaskValue.Value)
            Dim strNewValue As String = ""

            Dim blnAfter As Boolean = False
            Dim intIndex As Integer = 0

            For Each objToken As clsAutoNumberToken In objMaskValue.AutoNumber.AutoNumberTokens
                Select Case objToken.TokenType
                    Case clsAutoNumberToken.enumTokenType.CONSTANT
                        intIndex += 1

                    Case clsAutoNumberToken.enumTokenType.ANY, _
                    clsAutoNumberToken.enumTokenType.ANY_NOSPACE, _
                    clsAutoNumberToken.enumTokenType.LETTER, _
                    clsAutoNumberToken.enumTokenType.NUMBER
                        If blnAfter OrElse _
                        strMaskValue.Length < intIndex + 1 Then
                            '2016-08-11 -- Peter Melisi -- Bug fix for change in Auto Number length.
                            Exit For
                        End If
                        strNewValue &= strMaskValue.Substring(intIndex, 1)
                        intIndex += 1

                    Case clsAutoNumberToken.enumTokenType.DAY, _
                    clsAutoNumberToken.enumTokenType.MONTH
                        blnAfter = True

                    Case clsAutoNumberToken.enumTokenType.DAY_PAD, _
                    clsAutoNumberToken.enumTokenType.MONTH_PAD, _
                    clsAutoNumberToken.enumTokenType.YEAR_2
                        intIndex += 2

                    Case clsAutoNumberToken.enumTokenType.YEAR_4
                        intIndex += 4

                    Case clsAutoNumberToken.enumTokenType.SEQUENCE
                        If objMaskValue.AutoNumber.SequentialNumberPad = False Then
                            blnAfter = True
                        Else
                            intIndex += objMaskValue.AutoNumber.SequentialNumberLength
                        End If

                End Select
            Next

            objMaskValue.Value = strNewValue
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region " IDisposable Support "

    Protected Overrides Sub DisposeDBObject()
        Try
            If Not m_colLocks Is Nothing Then
                For Each objLock As clsRecordLock In m_colLocks.Values
                    Try
                        objLock.Delete(m_objDB)
                    Catch ex As Exception
                    End Try
                Next
                m_colLocks.Dispose()
                m_colLocks = Nothing
            End If

            If m_colTokens IsNot Nothing Then
                m_colTokens.Dispose()
                m_colTokens = Nothing
            End If

            m_colMultipleSequences = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
