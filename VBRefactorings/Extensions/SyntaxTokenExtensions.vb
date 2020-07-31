' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.
'

Imports System.Runtime.CompilerServices

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports VB = Microsoft.CodeAnalysis.VisualBasic
Imports VBFactory = Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory

Namespace Utilities
    Public Module SyntaxTokenExtensions

        <Extension>
        Friend Function Contains(Tokens As SyntaxTokenList, Kind As SyntaxKind) As Boolean
            Return Tokens.Contains(Function(m As SyntaxToken) m.IsKind(Kind))
        End Function

        <Extension>
        Friend Function Contains(Tokens As IEnumerable(Of SyntaxToken), ParamArray Kind() As VB.SyntaxKind) As Boolean
            Return Tokens.Contains(Function(m As SyntaxToken) m.IsKind(Kind))
        End Function

        <Extension>
        Friend Function IndexOf(Tokens As IEnumerable(Of SyntaxToken), Kind As VB.SyntaxKind) As Integer
            For i As Integer = 0 To Tokens.Count - 1
                If Tokens(i).IsKind(Kind) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        <Extension>
        Friend Function WithPrependedLeadingTrivia(Token As SyntaxToken, Trivia As IEnumerable(Of SyntaxTrivia)) As SyntaxToken
            Return Token.WithPrependedLeadingTrivia(Trivia.ToSyntaxTriviaList())
        End Function

        <Extension>
        Public Function [With](token As SyntaxToken, leading As List(Of SyntaxTrivia), trailing As List(Of SyntaxTrivia)) As SyntaxToken
            Return token.WithLeadingTrivia(leading).WithTrailingTrivia(trailing)
        End Function

        <Extension>
        Public Function [With](token As SyntaxToken, leading As SyntaxTriviaList, trailing As SyntaxTriviaList) As SyntaxToken
            Return token.WithLeadingTrivia(leading).WithTrailingTrivia(trailing)
        End Function

        <Extension>
        Public Function IsKind(token As SyntaxToken, ParamArray kinds() As VB.SyntaxKind) As Boolean
            Return kinds.Contains(CType(token.RawKind, VB.SyntaxKind))
        End Function

        <Extension()>
        Public Function RemoveExtraEOL(Token As SyntaxToken) As SyntaxToken
            Dim LeadingTrivia As New List(Of SyntaxTrivia)
            LeadingTrivia.AddRange(Token.LeadingTrivia())
            Select Case LeadingTrivia.Count
                Case 0
                    Return Token
                Case 1
                    If LeadingTrivia.First.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                        Return Token.WithLeadingTrivia(New SyntaxTriviaList)
                    End If
                Case 2
                    Select Case LeadingTrivia.First.RawKind
                        Case VB.SyntaxKind.WhitespaceTrivia
                            If LeadingTrivia.Last.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                                Return Token.WithLeadingTrivia(New SyntaxTriviaList)
                            End If
                            Return Token
                        Case VB.SyntaxKind.EndOfLineTrivia
                            If LeadingTrivia.Last.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                                Return Token.WithLeadingTrivia(New SyntaxTriviaList)
                            End If
                            Return Token.WithLeadingTrivia(LeadingTrivia.Last)
                        Case Else
                    End Select
                Case Else
            End Select
            Dim NewLeadingTrivia As New List(Of SyntaxTrivia)
            For Each e As IndexClass(Of SyntaxTrivia) In Token.LeadingTrivia.WithIndex
                Dim trivia As SyntaxTrivia = e.Value
                Dim nextTrivia As SyntaxTrivia = If(Not e.IsLast, Token.LeadingTrivia(e.Index + 1), Nothing)
                If trivia.IsKind(VB.SyntaxKind.WhitespaceTrivia) AndAlso nextTrivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                    Continue For
                End If
                If trivia.IsKind(VB.SyntaxKind.CommentTrivia) AndAlso nextTrivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                    NewLeadingTrivia.Add(trivia)
                    Continue For
                End If

                If trivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) AndAlso nextTrivia.IsKind(VB.SyntaxKind.EndOfLineTrivia) Then
                    Continue For
                End If
                NewLeadingTrivia.Add(trivia)
            Next

            Return Token.WithLeadingTrivia(NewLeadingTrivia)
        End Function

        <Extension>
        Public Function WithAppendedTrailingTrivia(token As SyntaxToken, trivia As IEnumerable(Of SyntaxTrivia)) As SyntaxToken
            Return token.WithTrailingTrivia(token.TrailingTrivia.Concat(trivia))
        End Function

        <Extension>
        Public Function WithPrependedLeadingTrivia(token As SyntaxToken, trivia As SyntaxTriviaList) As SyntaxToken
            If trivia.Count = 0 Then
                Return token
            End If

            Return token.WithLeadingTrivia(trivia.Concat(token.LeadingTrivia))
        End Function

    End Module
End Namespace
