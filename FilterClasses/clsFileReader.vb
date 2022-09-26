Imports K1.IFilter
Imports System.IO


Public Class clsFileReader

#Region "members"
    Private m_charFullText As String
    Private m_colLines As List(Of String)

#End Region

#Region "properties"


    ''' <summary>
    ''' Returns a string containing the all the extracted text from the file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FullText() As String
        Get
            Return m_charFullText
        End Get
    End Property


    ''' <summary>
    ''' Returns line(s) of text
    ''' </summary>
    ''' <param name="intLineNumber">The initial line number to be returned</param>
    ''' <param name="intCount">The number of consecutive lines to be returned after the initial line</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetLine(Optional ByVal intLineNumber As Integer = 0, Optional ByVal intCount As Integer = 1) As String()
        Get
            'vbCr
            Return GetLines(Chr(13), intLineNumber, intCount)
        End Get
    End Property

    ''' <summary>
    ''' Returns line(s) of text, with a soft break after 80 characters
    ''' </summary>
    Public ReadOnly Property Lines() As List(Of String)
        Get
            If m_colLines Is Nothing Then
                'vbCr
                m_colLines = SplitText(Chr(13), 80)
            End If

            Return m_colLines
        End Get
    End Property

    'Public ReadOnly Property GetParagraph(Optional ByVal intLineNumber As Integer = 0, Optional ByVal intCount As Integer = 1)
    '    Get
    '        Return GetLines(vbTab, intLineNumber, intCount)
    '    End Get
    'End Property

    'Public ReadOnly Property GetSentence(Optional ByVal intLineNumber As Integer = 0, Optional ByVal intCount As Integer = 1)
    '    Get
    '        Return GetLines(".", intLineNumber, intCount)
    '    End Get
    'End Property


#End Region

#Region "Constructors"
    ''' <summary>
    ''' Reads the text from a file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal strFilePath As String)
        'm_colLines.Clear()
        Try
            Dim objReader As TextReader = New K1.IFilter.clsFilterReader(strFilePath)
            Using objReader
                m_charFullText = objReader.ReadToEnd().Trim
            End Using
        Catch ex As Runtime.InteropServices.COMException
            Throw New Exception("COM Exception", ex)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Methods"

    Public Function SplitText(ByVal chrSplit As Char, ByVal intSoftBreak As Integer) As List(Of String)
        If m_charFullText.IndexOf(chrSplit) = -1 AndAlso chrSplit <> vbCr Then
            Return New List(Of String) From {
                m_charFullText
            }
        End If

        Dim colLines As New List(Of String)
        Dim colFiltered As New List(Of String)
        Dim textCopy As String = m_charFullText

        '2000003590 - XLSX files only have line-feeds
        If chrSplit = vbCr AndAlso Not m_charFullText.Contains(vbCr) Then
            textCopy = textCopy.Replace(vbLf, vbCr & vbLf)
        End If

        colFiltered = textCopy _
               .Replace(vbTab, String.Empty) _
               .Split(chrSplit) _
               .Where(Function(x) Not String.IsNullOrEmpty(x)) _
               .ToList()

        If colFiltered Is Nothing Then
            Return New List(Of String) From {
                m_charFullText
            }
        End If

        'Limits each line to 80 character preserving the word
        Try
            For Each line As String In colFiltered
                Dim splitStr As String() = line.Split(" "c)
                Dim sentencePart As New List(Of String) From {
                    String.Empty
                }
                Dim partCounter As Integer = 0

                For Each str As String In splitStr
                    If sentencePart(partCounter).Length + str.Length > intSoftBreak Then
                        partCounter += 1
                        sentencePart.Add(String.Empty)
                    End If
                    sentencePart(partCounter) += str + " "
                Next

                colLines.AddRange(sentencePart.Where(Function(x) x.Trim().Length <> 0).Select(Function(y) y.Trim()))
            Next
        Catch ex As Exception
            Throw
        End Try

        Return colLines
    End Function

    Private Function GetLines(ByVal chrCheck As Char, ByVal intLineNumber As Integer, ByVal intCount As Integer) As String()
        If m_colLines Is Nothing Then
            SplitText(chrCheck, 0)
        End If

        If intCount <= m_colLines.Count And intLineNumber >= 0 And intCount >= 0 Then
            Dim arrReturn(intCount - 1) As String
            m_colLines.CopyTo(intLineNumber, arrReturn, 0, intCount)

            Return arrReturn
        Else
            Return Nothing ' trying to return more lines than available or using a -ve array value
        End If
    End Function

#End Region

End Class

