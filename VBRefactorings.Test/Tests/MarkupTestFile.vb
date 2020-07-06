' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.CodeAnalysis.Text

Namespace Roslyn.UnitTestFramework
    ''' <summary>
    ''' To aid with testing, we define a special type of text file that can encode additional
    ''' information in it.  This prevents a test writer from having to carry around multiple sources
    ''' of information that must be reconstituted.  For example, instead of having to keep around the
    ''' contents of a file *and* and the location of the cursor, the tester can just provide a
    ''' string with the "$" character in it.  This allows for easy creation of "FIT" tests where all
    ''' that needs to be provided are strings that encode every bit of state necessary in the string
    ''' itself.
    '''
    ''' The current set of encoded features we support are:
    '''
    ''' $$ - The position in the file.  There can be at most one of these.
    '''
    ''' [| ... |] - A span of text in the file.  There can be many of these and they can be nested
    ''' and/or overlap the $ position.
    '''
    ''' {|Name: ... |} A span of text in the file annotated with an identifier.  There can be many of
    ''' these, including ones with the same name.
    '''
    ''' Additional encoded features can be added on a case by case basis.
    ''' </summary>
    Public Module MarkupTestFile
        Private Const PositionString As String = "$$"
        Private Const SpanStartString As String = "[|"
        Private Const SpanEndString As String = "|]"
        Private Const NamedSpanStartString As String = "{|"
        Private Const NamedSpanEndString As String = "|}"

        Private ReadOnly s_namedSpanStartRegex As New Regex("\{\| ([^:]+) \:", RegexOptions.Compiled Or RegexOptions.Multiline Or RegexOptions.IgnorePatternWhitespace)

        Private Sub Parse(input As String, ByRef output As String, ByRef position? As Integer, ByRef spans As IDictionary(Of String, IList(Of TextSpan)))
            position = Nothing
            spans = New Dictionary(Of String, IList(Of TextSpan))()

            Dim outputBuilder As New StringBuilder()

#If False Then
			output = NamedSpanStartRegex.Replace(input, String.Empty)
			output = output.Replace(PositionString, String.Empty).Replace(SpanStartString, String.Empty).Replace(SpanEndString, String.Empty).Replace(NamedSpanEndString, String.Empty)
#End If

            Dim currentIndexInInput As Integer = 0
            Dim inputOutputOffset As Integer = 0

            ' A stack of span starts along with their associated annotation name.  [||] spans simply
            ' have empty string for their annotation name.
            Dim spanStartStack As New Stack(Of Tuple(Of Integer, String))()

            Do
                Dim matches As New List(Of Tuple(Of Integer, String))()
                AddMatch(input, PositionString, currentIndexInInput, matches)
                AddMatch(input, SpanStartString, currentIndexInInput, matches)
                AddMatch(input, SpanEndString, currentIndexInInput, matches)
                AddMatch(input, NamedSpanEndString, currentIndexInInput, matches)

                Dim namedSpanStartMatch As Match = s_namedSpanStartRegex.Match(input, currentIndexInInput)
                If namedSpanStartMatch.Success Then
                    matches.Add(Tuple.Create(namedSpanStartMatch.Index, namedSpanStartMatch.Value))
                End If

                If matches.Count = 0 Then
                    ' No more markup to process.
                    Exit Do
                End If

                Dim orderedMatches As List(Of Tuple(Of Integer, String)) = matches.OrderBy(Function(t1, t2) t1.Item1 - t2.Item1).ToList()
                If orderedMatches.Count >= 2 AndAlso spanStartStack.Count > 0 AndAlso matches(0).Item1 = matches(1).Item1 - 1 Then
                    ' We have a slight ambiguity with cases like these:
                    '
                    ' [|]    [|}
                    '
                    ' Is it starting a new match, or ending an existing match.  As a workaround, we
                    ' special case these and consider it ending a match if we have something on the
                    ' stack already.
                    If (matches(0).Item2 = SpanStartString AndAlso matches(1).Item2 = SpanEndString AndAlso String.IsNullOrEmpty(spanStartStack.Peek().Item2)) OrElse (matches(0).Item2 = SpanStartString AndAlso matches(1).Item2 = NamedSpanEndString AndAlso Not String.IsNullOrEmpty(spanStartStack.Peek().Item2)) Then
                        orderedMatches.RemoveAt(0)
                    End If
                End If

                ' Order the matches by their index
                Dim firstMatch As Tuple(Of Integer, String) = orderedMatches.First()

                Dim matchIndexInInput As Integer = firstMatch.Item1
                Dim matchString As String = firstMatch.Item2

                Dim matchIndexInOutput As Integer = matchIndexInInput - inputOutputOffset
                outputBuilder.Append(input.Substring(currentIndexInInput, matchIndexInInput - currentIndexInInput))

                currentIndexInInput = matchIndexInInput + matchString.Length
                inputOutputOffset += matchString.Length

                Select Case matchString.Substring(0, 2)
                    Case PositionString
                        If position.HasValue Then
                            Throw New ArgumentException(
                                String.Format(CultureInfo.InvariantCulture,
                                              "Saw multiple occurrences of {0}",
                                              PositionString)
                                              )
                        End If

                        position = matchIndexInOutput

                    Case SpanStartString
                        spanStartStack.Push(Tuple.Create(matchIndexInOutput, String.Empty))

                    Case SpanEndString
                        If spanStartStack.Count = 0 Then
                            Throw New ArgumentException(
                                String.Format(CultureInfo.InvariantCulture,
                                              "Saw {0} without matching {1}",
                                              SpanEndString,
                                              SpanStartString)
                                              )
                        End If

                        If spanStartStack.Peek().Item2.Length > 0 Then
                            Throw New ArgumentException(
                                String.Format(CultureInfo.InvariantCulture,
                                              "Saw {0} without matching {1}",
                                              NamedSpanStartString,
                                              NamedSpanEndString)
                                              )
                        End If

                        PopSpan(spanStartStack, spans, matchIndexInOutput)

                    Case NamedSpanStartString
                        Dim name As String = namedSpanStartMatch.Groups(1).Value
                        spanStartStack.Push(Tuple.Create(matchIndexInOutput, name))

                    Case NamedSpanEndString
                        If spanStartStack.Count = 0 Then
                            Throw New ArgumentException(
                                String.Format(CultureInfo.InvariantCulture,
                                              "Saw {0} without matching {1}",
                                              NamedSpanEndString,
                                              NamedSpanStartString)
                                              )
                        End If

                        If spanStartStack.Peek().Item2.Length = 0 Then
                            Throw New ArgumentException(
                                String.Format(CultureInfo.InvariantCulture,
                                              "Saw {0} without matching {1}",
                                              SpanStartString,
                                              SpanEndString)
                                              )
                        End If

                        PopSpan(spanStartStack, spans, matchIndexInOutput)

                    Case Else
                        Throw New InvalidOperationException()
                End Select
            Loop

            If spanStartStack.Count > 0 Then
                Throw New ArgumentException(
                    String.Format(CultureInfo.InvariantCulture,
                                  "Saw {0} without matching {1}",
                                  SpanStartString,
                                  SpanEndString)
                                  )
            End If

            ' Append the remainder of the string.
            outputBuilder.Append(input.Substring(currentIndexInInput))
            output = outputBuilder.ToString()
        End Sub

        Private Sub PopSpan(spanStartStack As Stack(Of Tuple(Of Integer, String)), spans As IDictionary(Of String, IList(Of TextSpan)), finalIndex As Integer)
            Dim spanStartTuple As Tuple(Of Integer, String) = spanStartStack.Pop()

            Dim span As TextSpan = TextSpan.FromBounds(spanStartTuple.Item1, finalIndex)
            spans.GetOrAdd(spanStartTuple.Item2, Function() New List(Of TextSpan)()).Add(span)
        End Sub
        Private Sub AddMatch(input As String, value As String, currentIndex As Integer, matches As List(Of Tuple(Of Integer, String)))
            Dim index As Integer = input.IndexOf(value, currentIndex, StringComparison.InvariantCulture)
            If index >= 0 Then
                matches.Add(Tuple.Create(index, value))
            End If
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPositionOpt? As Integer, <Out()> ByRef spans As IDictionary(Of String, IList(Of TextSpan)))
            Parse(input, output, cursorPositionOpt, spans)
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef cursorPositionOpt? As Integer, <Out()> ByRef spans As IDictionary(Of String, IList(Of TextSpan)))
            Dim output As String = Nothing
            GetPositionAndSpans(input, output, cursorPositionOpt, spans)
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPosition As Integer, <Out()> ByRef spans As IDictionary(Of String, IList(Of TextSpan)))
            Dim cursorPositionOpt? As Integer = Nothing
            GetPositionAndSpans(input, output, cursorPositionOpt, spans)

            cursorPosition = cursorPositionOpt.Value
        End Sub

        Public Sub GetSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef spans As IDictionary(Of String, IList(Of TextSpan)))
            Dim cursorPositionOpt? As Integer = Nothing
            GetPositionAndSpans(input, output, cursorPositionOpt, spans)
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPositionOpt? As Integer, <Out()> ByRef spans As IList(Of TextSpan))
            Dim dictionary As IDictionary(Of String, IList(Of TextSpan)) = Nothing
            Parse(input, output, cursorPositionOpt, dictionary)

            spans = dictionary.GetOrAdd(String.Empty, Function() New List(Of TextSpan)())
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef cursorPositionOpt? As Integer, <Out()> ByRef spans As IList(Of TextSpan))
            Dim output As String = Nothing
            GetPositionAndSpans(input, output, cursorPositionOpt, spans)
        End Sub

        Public Sub GetPositionAndSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPosition As Integer, <Out()> ByRef spans As IList(Of TextSpan))
            Dim pos? As Integer = Nothing
            GetPositionAndSpans(input, output, pos, spans)

            cursorPosition = If(pos, 0)
        End Sub

        Public Sub GetPosition(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPosition As Integer)
            Dim spans As IList(Of TextSpan) = Nothing
            GetPositionAndSpans(input, output, cursorPosition, spans)
        End Sub

        Public Sub GetPositionAndSpan(input As String, <Out()> ByRef output As String, <Out()> ByRef cursorPosition As Integer, <Out()> ByRef span As TextSpan)
            Dim spans As IList(Of TextSpan) = Nothing
            GetPositionAndSpans(input, output, cursorPosition, spans)

            span = spans.Single()
        End Sub

        Public Sub GetSpans(input As String, <Out()> ByRef output As String, <Out()> ByRef spans As IList(Of TextSpan))
            Dim pos? As Integer = Nothing
            GetPositionAndSpans(input, output, pos, spans)
        End Sub

        Public Sub GetSpan(input As String, <Out()> ByRef output As String, <Out()> ByRef span As TextSpan)
            Dim spans As IList(Of TextSpan) = Nothing
            GetSpans(input, output, spans)
            span = spans.Single()
        End Sub

        Public Function CreateTestFile(code As String, cursor As Integer) As String
            If String.IsNullOrWhiteSpace(code) Then
                Throw New ArgumentException($"'{NameOf(code)}' cannot be null or whitespace", NameOf(code))
            End If

            Return CreateTestFile(code, DirectCast(Nothing, IDictionary(Of String, IList(Of TextSpan))), cursor)
        End Function

        Public Function CreateTestFile(code As String, spans As IList(Of TextSpan), Optional cursor As Integer = -1) As String
            If String.IsNullOrWhiteSpace(code) Then
                Throw New ArgumentException($"'{NameOf(code)}' cannot be null or whitespace", NameOf(code))
            End If

            Return CreateTestFile(code, New Dictionary(Of String, IList(Of TextSpan)) From {
            {String.Empty, spans}
        },
        cursor)
        End Function

        Friend Function CreateTestFile(code As String, spans As IDictionary(Of String, IList(Of TextSpan)), Optional cursor As Integer = -1) As String
            Dim sb As New StringBuilder()
            Dim anonymousSpans As IList(Of TextSpan) = spans.GetOrAdd(String.Empty, Function() New List(Of TextSpan)())

            For i As Integer = 0 To code.Length
                If i = cursor Then
                    sb.Append(PositionString)
                End If

                AddSpanString(sb, spans.Where(Function(kvp) Not String.IsNullOrEmpty(kvp.Key)), i, start:=True)
                AddSpanString(sb, spans.Where(Function(kvp) String.IsNullOrEmpty(kvp.Key)), i, start:=True)
                AddSpanString(sb, spans.Where(Function(kvp) String.IsNullOrEmpty(kvp.Key)), i, start:=False)
                AddSpanString(sb, spans.Where(Function(kvp) Not String.IsNullOrEmpty(kvp.Key)), i, start:=False)

                If i < code.Length Then
                    sb.Append(code.Chars(i))
                End If
            Next i

            Return sb.ToString()
        End Function

        Private Sub AddSpanString(sb As StringBuilder, items As IEnumerable(Of KeyValuePair(Of String, IList(Of TextSpan))), position As Integer, start As Boolean)
            For Each kvp As KeyValuePair(Of String, IList(Of TextSpan)) In items
                For Each span As TextSpan In kvp.Value
                    If start AndAlso span.Start = position Then
                        If String.IsNullOrEmpty(kvp.Key) Then
                            sb.Append(SpanStartString)
                        Else
                            sb.Append(NamedSpanStartString)
                            sb.Append(kvp.Key)
                            sb.Append(":"c)
                        End If
                    ElseIf Not start AndAlso span.End = position Then
                        If String.IsNullOrEmpty(kvp.Key) Then
                            sb.Append(SpanEndString)
                        Else
                            sb.Append(NamedSpanEndString)
                        End If
                    End If
                Next span
            Next kvp
        End Sub
    End Module
End Namespace
