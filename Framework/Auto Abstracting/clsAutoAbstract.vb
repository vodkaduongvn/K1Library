Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text

Public Class clsAutoAbstractBuilder

#Region " Members "

    Private m_strBody As String
    Private m_colSeedList As FrameworkCollections.K1Dictionary(Of Integer)
    Private m_colStopWords As List(Of String)
    Private m_intMaxSentences As Int64
    Private m_colSentenceCollection As Generic.List(Of clsSentence)
    Private m_colHitScoreIndex As SortedList(Of Int64, List(Of Integer))

#End Region

#Region " Constructors "

    Public Sub New(ByVal strBody As String, ByVal colSeedList As FrameworkCollections.K1Dictionary(Of Integer), _
                   ByVal colStopWords As List(Of String), ByVal blnRemoveStopWords As Boolean, ByVal intMaxSentences As Int64)
        Me.New(colSeedList, colStopWords, blnRemoveStopWords, intMaxSentences)

        m_strBody = strBody
    End Sub

    Public Sub New(ByVal colSeedList As FrameworkCollections.K1Dictionary(Of Integer), ByVal colStopWords As List(Of String), _
                   ByVal blnRemoveStopWords As Boolean, ByVal intMaxSentences As Int64)
        m_strBody = String.Empty
        m_colSeedList = colSeedList
        If blnRemoveStopWords Then
            m_colStopWords = colStopWords
        End If
        m_intMaxSentences = intMaxSentences

        m_colSentenceCollection = New Generic.List(Of clsSentence)
        m_colHitScoreIndex = New SortedList(Of Int64, List(Of Integer))
    End Sub

#End Region

#Region " Set Text "

    ''' <summary>
    ''' Sets given text as the text to build the abstract from.
    ''' </summary>
    ''' <param name="strText">Text to set for abstracting.</param>
    Public Sub SetText(ByVal strText As String)
        m_strBody = strText
    End Sub

#End Region

#Region " Load text "

    ''' <summary>
    ''' Uses the extracted text from the given file to build the abstract from.
    ''' </summary>
    ''' <param name="strFilePath">Full path to the file to abstract.</param>
    Public Sub LoadText(ByVal strFilePath As String)
        Dim objTextFilter As New K1Library.clsFileReader(strFilePath)
        m_strBody = objTextFilter.FullText
    End Sub

#End Region

#Region " Build "

    ''' <summary>
    ''' Builds an abstract from the given text and seed list.
    ''' </summary>
    ''' <returns>Abstract of given text.</returns>
    Public Function Build() As String
        Try
            Dim strAbstract As String = String.Empty

            If m_colSeedList Is Nothing OrElse m_colSeedList.Count = 0 Then
                '-- No Seed Words to use
                Return strAbstract
            End If

            Dim intCount As Int64
            Dim objBuilder As New StringBuilder
            InitaliseSentenceCollection()

            Do Until m_colSentenceCollection.Count = 0 OrElse intCount >= m_intMaxSentences
                '-- Get Sentence with Highest Hit Score
                Dim intIndex As Integer = m_colHitScoreIndex.Last.Value(0)
                Dim objSentence As clsSentence = m_colSentenceCollection(intIndex)

                objBuilder.Append(objSentence.Sentence & ". ")
                RemoveHighestSentence()

                intCount += 1
            Loop

            strAbstract = objBuilder.ToString

            ' Remove any matching words from arrStopWord
            If m_colStopWords IsNot Nothing Then
                strAbstract = RemoveStopWordsFromAbstract(strAbstract)
            End If

            Return strAbstract
        Finally
            m_colSentenceCollection.Clear()
            m_colHitScoreIndex.Clear()
        End Try
    End Function

#End Region

#Region " Do Auto Abstract "

    ''' <summary>
    ''' Main Function that performs auto abstracting. We assume that keywords and stop words
    ''' are case insentive and we are matching only whole words to them.
    ''' (Deprecated see Build)
    ''' </summary>
    ''' <param name="strBody">Body of text.</param>
    ''' <param name="dicSeedList">Seed list used to determine the importance of sentences.</param>
    ''' <param name="colStopWords">Words that are considered noise.</param>
    ''' <param name="blnRemoveStopWords">Should stop words be removed from the the final abstract.</param>
    ''' <param name="intMaxSentences">Maximum number of sentences to include in the abstract.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' This function is left only to support compatability with products using the old code. All products should 
    ''' use Build instead.
    ''' </remarks>
    Public Shared Function DoAutoAbstract(ByVal strBody As String, ByVal dicSeedList As FrameworkCollections.K1Dictionary(Of Integer), _
        ByVal colStopWords As List(Of String), ByVal blnRemoveStopWords As Boolean, ByVal intMaxSentences As Int64) As String
        Try
            Dim objAbstractBuilder As New clsAutoAbstractBuilder(strBody, dicSeedList, colStopWords, blnRemoveStopWords, intMaxSentences)
            Return objAbstractBuilder.Build
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Initalise Sentence Collection "

    ''' <summary>
    ''' Splits the text extracted from the document into sentences and builds an index 
    ''' ordering the Hit Score in ascending order. i.e. most relevant last
    ''' </summary>
    ''' <remarks>We consider End of Sentence (EOS) punctuation as either . or ? or ! or tab or carriage return.</remarks>
    Private Sub InitaliseSentenceCollection()
        Try
            Dim arrSentences() As String = Regex.Split(m_strBody, "[.?!\t\r]+")

            For Each strSentence As String In arrSentences
                strSentence = strSentence.Trim

                If strSentence.Length > 0 Then
                    Dim colFoundWords As New List(Of String)
                    Dim intHitScore As Int64 = GetHitScore(strSentence, colFoundWords)

                    If intHitScore > 0 Then
                        AddSentence(strSentence, intHitScore, colFoundWords)
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region " Add Sentence "

    ''' <summary>
    ''' Adds the sentence to the default list and orders it by its hit score. 
    ''' </summary>
    ''' <param name="strSentence">Sentence to add to the list.</param>
    ''' <param name="intHitScore">Hit Score value of the sentence.</param>
    ''' <param name="colFoundWords">Seed words found in the sentence.</param>
    Private Sub AddSentence(ByVal strSentence As String, ByVal intHitScore As Int64, ByVal colFoundWords As List(Of String))
        Dim objSentence As New clsSentence(strSentence, intHitScore, colFoundWords)
        AddSentence(m_colSentenceCollection, m_colHitScoreIndex, objSentence)
    End Sub

    ''' <summary>
    ''' Adds the sentence to the given list and orders it by its hit score. 
    ''' </summary>
    ''' <param name="colSentenceCollection"></param>
    ''' <param name="colHitScoreIndex"></param>
    ''' <param name="objSentence"></param>
    ''' <remarks></remarks>
    Private Sub AddSentence(ByVal colSentenceCollection As List(Of clsSentence), _
                            ByVal colHitScoreIndex As SortedList(Of Int64, List(Of Integer)), ByVal objSentence As clsSentence)
        Dim intIndex As Integer

        colSentenceCollection.Add(objSentence)
        intIndex = colSentenceCollection.Count - 1
        If Not colHitScoreIndex.ContainsKey(objSentence.HitScore) Then
            colHitScoreIndex.Add(objSentence.HitScore, New List(Of Integer))
        End If
        colHitScoreIndex(objSentence.HitScore).Add(intIndex)
    End Sub

#End Region

#Region " Remove Highest Sentence "

    ''' <summary>
    ''' Removes the sentence with the highest hit score from the list.
    ''' </summary>
    ''' <remarks>
    ''' When a sentence is removed from the list the seed words are also removed.
    ''' This means that sentences that share seed words need their hit score recalculated
    ''' and those that are no longer relevant (i.e. hit score = 0) are also removed.
    ''' </remarks>
    Private Sub RemoveHighestSentence()
        Dim intIndex As Integer = m_colHitScoreIndex.Last.Value(0)
        Dim objDeletedSentence As clsSentence = m_colSentenceCollection(intIndex)

        '-- Remove Sentence for the list
        m_colSentenceCollection.RemoveAt(intIndex)
        m_colHitScoreIndex(objDeletedSentence.HitScore).Remove(intIndex)

        Dim colSentenceCollection As New List(Of clsSentence)
        Dim colHitScoreIndex As New SortedList(Of Int64, List(Of Integer))

        '-- Remove seed words in deleted sentence from list before recacluting
        For Each strSeed As String In objDeletedSentence.FoundSeedWords
            m_colSeedList.Remove(strSeed)
        Next

        '-- Recalculate HitScore for sentences with removed Seed Words
        For Each objCurrentSentence As clsSentence In m_colSentenceCollection
            Dim blnFound As Boolean = False

            '-- Find Sentences that share seed words with the sentence that was removed
            For Each strSeed As String In objDeletedSentence.FoundSeedWords
                If objCurrentSentence.FoundSeedWords.Contains(strSeed) Then
                    Dim colFoundWords As New List(Of String)
                    Dim intHitScore As Int64 = GetHitScore(objCurrentSentence.Sentence, colFoundWords)
                    blnFound = True

                    If intHitScore > 0 Then
                        '-- Sentence is still relevant so update
                        objCurrentSentence.HitScore = intHitScore
                        objCurrentSentence.FoundSeedWords = colFoundWords

                        AddSentence(colSentenceCollection, colHitScoreIndex, objCurrentSentence)
                    End If

                    Exit For
                End If
            Next

            If Not blnFound Then
                '-- Include sentences that do not share seed words as hit score does not change
                AddSentence(colSentenceCollection, colHitScoreIndex, objCurrentSentence)
            End If
        Next

        '-- Save new sentence list
        m_colSentenceCollection = colSentenceCollection
        m_colHitScoreIndex = colHitScoreIndex
    End Sub

#End Region

#Region " Get Hit Score "

    ''' <summary>
    ''' Calculates a sentence's relevance score given a dictionary of seed words
    ''' and their value. 
    ''' </summary>
    ''' <param name="strSentence">Sentence that a hit score should be calculate for.</param>
    ''' <param name="colFoundWords">List of seed list words found in the sentence.</param>
    ''' <returns>relevance score of string.</returns>
    ''' <remarks>
    ''' We assume that keywords are case insentive and we are matching only
    ''' whole words to them. Higher score = more relevant keywords found.
    ''' </remarks>
    Private Function GetHitScore(ByVal strSentence As String, ByRef colFoundWords As List(Of String)) As Int64
        Dim intTotalHits As Int64
        Dim intHits As Integer
        Dim arrPhrases() As String

        Try
            arrPhrases = strSentence.Split(" "c)

            ' Search String for each Keyword frequency
            For Each strSeedWord As String In m_colSeedList.Keys
                intHits = 0

                For Each strWord As String In arrPhrases
                    If String.Compare(strWord, strSeedWord, True) = 0 Then
                        intHits += 1
                    End If
                Next

                intTotalHits += (intHits * m_colSeedList(strSeedWord))

                If intHits > 0 Then
                    If colFoundWords IsNot Nothing AndAlso Not colFoundWords.Contains(strSeedWord) Then
                        '-- We found a new seed list word
                        colFoundWords.Add(strSeedWord)
                    End If
                End If
            Next

            Return intTotalHits
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region " Remove Stop Words From Abstract "

    ''' <summary>
    ''' Removes stop words from the final abstract. 
    ''' </summary>
    ''' <param name="strAbstract">The abstract text to remove stop words from.</param>
    ''' <returns>New abstract without stop words.</returns>
    ''' <remarks>
    ''' We assume that stop words are case insentive and we are matching
    ''' only whole words to them.
    ''' </remarks>
    Private Function RemoveStopWordsFromAbstract(ByVal strAbstract As String) As String
        Dim arrPhrases() As String
        Dim strBuilder As New StringBuilder
        Dim blnStopWordExists As Boolean = False

        Try
            'Remove unwanted characters and then split for processing
            arrPhrases = strAbstract.Split(" "c)

            For Each strAbstractWord As String In arrPhrases
                blnStopWordExists = False
                For Each strStopWord As String In m_colStopWords
                    If String.Compare(clsSentence.RemovePunctuations(strAbstractWord), strStopWord, True) = 0 Then
                        blnStopWordExists = True
                        Exit For
                    End If
                Next

                ' If no matches then add to new Abstract
                If Not blnStopWordExists Then
                    If strBuilder.Length > 0 Then
                        strBuilder.Append(" ")
                    End If

                    strBuilder.Append(strAbstractWord)
                End If
            Next

            Return strBuilder.ToString()
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class
