' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis

Imports VB = Microsoft.CodeAnalysis.VisualBasic
Imports VBFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
Imports VBS = Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Utilities
    Public Module SyntaxTriviaExtensions

        <Extension>
        Friend Function ContainsEOLTrivia(node As VB.VisualBasicSyntaxNode) As Boolean
            If Not node.HasTrailingTrivia Then
                Return False
            End If
            Dim TriviaList As SyntaxTriviaList = node.GetTrailingTrivia
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsCommentOrDirectiveTrivia(TriviaList As List(Of SyntaxTrivia)) As Boolean
            If TriviaList Is Nothing OrElse TriviaList.Count = 0 Then
                Return False
            End If
            For Each t As SyntaxTrivia In TriviaList
                If t.IsWhitespaceOrEndOfLine Then
                    Continue For
                End If
                If t.IsNone Then
                    Continue For
                End If
                If t.IsCommentOrDirectiveTrivia Then
                    Return True
                End If
                If t.IsKind(VB.SyntaxKind.LineContinuationTrivia) Then
                    Return True
                End If
                If t.IsKind(VB.SyntaxKind.DocumentationCommentTrivia) Then
                    Return True
                End If
                If t.IsKind(VB.SyntaxKind.DocumentationCommentExteriorTrivia) Then
                    Return True
                End If

                Throw UnexpectedValue(t.ToString)
            Next
            Return False
        End Function

        ''' <summary>
        ''' Syntax Trivia in any Language
        ''' </summary>
        ''' <param name="TriviaList"></param>
        ''' <returns>True if any Trivia is a Comment or a Directive</returns>
        <Extension>
        Public Function ContainsCommentOrDirectiveTriviaOrContinuation(TriviaList As SyntaxTriviaList) As Boolean
            If TriviaList.Count = 0 Then
                Return False
            End If
            Return TriviaList.ToList.Any(Function(t) t.IsKind(VB.SyntaxKind.LineContinuationTrivia) OrElse t.IsCommentOrDirectiveTrivia)

        End Function

        ''' <summary>
        ''' Syntax Trivia in any Language
        ''' </summary>
        ''' <param name="TriviaList"></param>
        ''' <returns>True if any Trivia is a Comment or a Directive</returns>
        <Extension>
        Public Function ContainsDirectiveTrivia(TriviaList As SyntaxTriviaList, ParamArray Kinds() As VB.SyntaxKind) As Boolean
            If TriviaList.Count = 0 Then Return False
            For Each t As SyntaxTrivia In TriviaList
                If t.IsDirective Then
                    If Kinds.Length = 0 Then
                        Return True
                    End If
                    For Each k As VB.SyntaxKind In Kinds
                        If t.RawKind = k Then
                            Return True
                        End If
                    Next
                    Return False
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsCommentOrLineContinueTrivia(TriviaList As SyntaxTriviaList) As Boolean
            Return ContainsCommentOrLineContinueTrivia(TriviaList.ToList)
        End Function

        <Extension>
        Public Function ContainsCommentOrLineContinueTrivia(TriviaList As List(Of SyntaxTrivia)) As Boolean
            If TriviaList.Count = 0 Then
                Return False
            End If
            For Each t As SyntaxTrivia In TriviaList
                If t.IsWhitespaceOrEndOfLine Then
                    Continue For
                End If
                If t.IsNone Then
                    Continue For
                End If
                If t.IsKind(VB.SyntaxKind.LineContinuationTrivia) Then
                    Return True
                End If
                If t.IsComment Then
                    Return True
                End If
                If t.IsDirective Then
                    Continue For
                End If
                Stop
            Next
            Return False
        End Function

        ''' <summary>
        ''' Syntax Trivia in any Language
        ''' </summary>
        ''' <param name="TriviaList"></param>
        ''' <returns>True if any Trivia is EndIf Directive</returns>
        <Extension>
        Public Function ContainsEndIfTrivia(TriviaList As SyntaxTriviaList) As Boolean
            If TriviaList.Count = 0 Then Return False
            For Each t As SyntaxTrivia In TriviaList
                If t.IsKind(VB.SyntaxKind.EndIfDirectiveTrivia) Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsEOLTrivia(Token As SyntaxToken) As Boolean
            If Not Token.HasTrailingTrivia Then
                Return False
            End If
            Dim TriviaList As SyntaxTriviaList = Token.TrailingTrivia
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function ContainsEOLTrivia(TriviaList As SyntaxTriviaList) As Boolean
            For Each t As SyntaxTrivia In TriviaList
                If t.IsEndOfLine Then
                    Return True
                End If
            Next
            Return False
        End Function

        <Extension>
        Public Function FullWidth(trivia As SyntaxTrivia) As Integer
            Return trivia.FullSpan.Length
        End Function

        <Extension>
        Public Function IsComment(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineComment
        End Function

        <Extension>
        Public Function IsCommentOrDirectiveTrivia(t As SyntaxTrivia) As Boolean
            If t.IsSingleLineComment Then
                Return True
            End If
            If t.IsDirective Then
                Return True
            End If
            If t.IsKind(VB.SyntaxKind.DocumentationCommentTrivia) Then
                Return True
            End If
            If t.IsKind(VB.SyntaxKind.DocumentationCommentExteriorTrivia) Then
                Return True
            End If
            Return False
        End Function

        <Extension>
        Public Function IsEndOfLine(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(VB.SyntaxKind.EndOfLineTrivia)
        End Function


        <Extension>
        Public Function IsNone(trivia As SyntaxTrivia) As Boolean
            Return trivia.RawKind = 0
        End Function

        <Extension>
        Public Function IsRegularOrDocComment(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsSingleLineComment()
        End Function

        <Extension>
        Public Function IsSingleLineComment(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(VB.SyntaxKind.CommentTrivia)
        End Function

        <Extension>
        Public Function IsWhitespace(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) OrElse
                trivia.IsKind(VB.SyntaxKind.WhitespaceTrivia)
        End Function

        <Extension>
        Public Function IsWhitespaceOrEndOfLine(trivia As SyntaxTrivia) As Boolean
            Return trivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) OrElse
                trivia.IsKind(VB.SyntaxKind.WhitespaceTrivia)
        End Function

        <Extension>
        Public Function MatchesKind(trivia As SyntaxTrivia, ParamArray kinds() As VB.SyntaxKind) As Boolean
            For Each kind As VB.SyntaxKind In kinds
                If trivia.IsKind(kind) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Remove directive trivia
        ''' </summary>
        ''' <param name="Token"></param>
        ''' <returns></returns>
        <Extension>
        Public Function RemoveDirectiveTrivia(Token As SyntaxToken, ByRef FoundEOL As Boolean) As SyntaxToken
            Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
            Dim NewTrailingTrivia As New List(Of SyntaxTrivia)

            For Each trivia As SyntaxTrivia In Token.LeadingTrivia
                Select Case trivia.RawKind
                    Case VB.SyntaxKind.WhitespaceTrivia, VB.SyntaxKind.CommentTrivia
                        NewLeadingTrivia.Add(trivia)
                        FoundEOL = False
                    Case VB.SyntaxKind.EndOfLineTrivia
                        If Not FoundEOL Then
                            NewLeadingTrivia.Add(trivia)
                            FoundEOL = True
                        End If
                    Case VB.SyntaxKind.DisabledTextTrivia,
                        VB.SyntaxKind.IfDirectiveTrivia,
                        VB.SyntaxKind.ElseDirectiveTrivia,
                        VB.SyntaxKind.ElseIfDirectiveTrivia,
                        VB.SyntaxKind.EndIfDirectiveTrivia
                        ' skip
                    Case Else
                        Stop
                End Select
            Next

            For Each trivia As SyntaxTrivia In Token.TrailingTrivia
                Select Case trivia.RawKind
                    Case VB.SyntaxKind.WhitespaceTrivia, VB.SyntaxKind.CommentTrivia
                        NewTrailingTrivia.Add(trivia)
                        FoundEOL = False
                    Case VB.SyntaxKind.EndOfLineTrivia
                        If Not FoundEOL Then
                            NewTrailingTrivia.Add(trivia)
                            FoundEOL = True
                        End If
                    Case VB.SyntaxKind.DisableWarningDirectiveTrivia,
                         VB.SyntaxKind.IfDirectiveTrivia,
                         VB.SyntaxKind.ElseDirectiveTrivia,
                         VB.SyntaxKind.ElseIfDirectiveTrivia,
                         VB.SyntaxKind.EndIfDirectiveTrivia
                        ' skip
                    Case Else
                        Stop
                End Select
            Next

            Return Token.With(NewLeadingTrivia, NewTrailingTrivia)
        End Function

        <Extension>
        Public Function ToSyntaxTriviaList(l As IEnumerable(Of SyntaxTrivia)) As SyntaxTriviaList
            Dim NewSyntaxTriviaList As New SyntaxTriviaList
            Return NewSyntaxTriviaList.AddRange(l)
        End Function

        <Extension>
        Public Function ToSyntaxTriviaList(l As List(Of SyntaxTrivia)) As SyntaxTriviaList
            Dim NewSyntaxTriviaList As New SyntaxTriviaList
            Return NewSyntaxTriviaList.AddRange(l)
        End Function

    End Module
End Namespace
