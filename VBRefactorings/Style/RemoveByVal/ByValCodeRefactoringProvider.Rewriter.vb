' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Namespace Style

    Partial Class ByValCodeRefactoringProvider

        Class ByValCodeRefactoringProviderRewriter
            Inherits VisualBasicSyntaxRewriter

            Private ReadOnly _predicate As Func(Of SyntaxToken, Boolean)

            Public Sub New(predicate As Func(Of SyntaxToken, Boolean))
                _predicate = predicate
            End Sub

            Public Overrides Function VisitToken(token As SyntaxToken) As SyntaxToken
                If token.Kind = SyntaxKind.ByValKeyword AndAlso _predicate(token) Then
                    Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                    Dim TrailingTrivia As SyntaxTriviaList = token.TrailingTrivia
                    Dim TriviaUBound As Integer = TrailingTrivia.Count - 1
                    Dim FirstContinuation As Boolean = True
                    If TriviaUBound > 1 Then
                        For i As Integer = 0 To TriviaUBound
                            Dim Trivia As SyntaxTrivia = TrailingTrivia(i)
                            Dim NextTrivia As SyntaxTrivia = If(i < TriviaUBound, TrailingTrivia(i + 1), Nothing)
                            If Trivia.IsKind(SyntaxKind.WhitespaceTrivia) AndAlso NextTrivia.IsKind(SyntaxKind.LineContinuationTrivia) Then
                                If FirstContinuation Then
                                    i += 2
                                    FirstContinuation = False
                                    Continue For
                                End If
                                If Trivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                                    FirstContinuation = False
                                End If
                            End If
                            NewTrailingTrivia.Add(Trivia)
                        Next
                    End If
                    Return SyntaxFactory.Token(token.LeadingTrivia, SyntaxKind.EmptyToken, NewTrailingTrivia.ToSyntaxTriviaList)
                End If

                Return token
            End Function

        End Class

    End Class

End Namespace
