' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Editing
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Imports VBRefactorings.Utilities

Namespace Style

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(ByValAnalyzerFixCodeFixProvider)), [Shared]>
    Partial Public Class ByValAnalyzerFixCodeFixProvider
        Inherits CodeFixProvider

        Private Const Title As String = "Remove ByVal"

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(ByValAnalyzerFixAnalyzer.DiagnosticId)
            End Get
        End Property

        Private Async Function RemoveByValOccuranceAsync(document As Document, parameterList As ParameterListSyntax, CancelToken As CancellationToken) As Task(Of Document)
            Dim editor As DocumentEditor = Await DocumentEditor.CreateAsync(document, CancelToken)
            Dim parametersCount As Integer = parameterList.Parameters.Count - 1
            Dim newParamsList As New List(Of ParameterSyntax)
            For i As Integer = 0 To parametersCount
                Dim Param As ParameterSyntax = parameterList.Parameters(i)
                If Param.Modifiers.Count = 0 Then
                    newParamsList.Add(Param)
                    Continue For
                End If
                Dim indexOfByVal As Integer = Param.Modifiers.IndexOf(SyntaxKind.ByValKeyword)
                If indexOfByVal < 0 Then
                    newParamsList.Add(Param)
                    Continue For
                End If

                Dim byValToken As SyntaxToken = Param.Modifiers(indexOfByVal)
                Dim newLeadingTrivia As List(Of SyntaxTrivia) = byValToken.LeadingTrivia.ToList
                If byValToken.LeadingTrivia.Any AndAlso byValToken.LeadingTrivia.ContainsCommentOrLineContinueTrivia Then
                    If newLeadingTrivia.Last.IsKind(SyntaxKind.WhitespaceTrivia) Then
                        newLeadingTrivia.RemoveAt(newLeadingTrivia.Count - 1)
                    End If
                End If

                Dim newTrailingTrivia As List(Of SyntaxTrivia) = byValToken.TrailingTrivia.ToList
                If newTrailingTrivia.Any AndAlso newTrailingTrivia.ContainsCommentOrLineContinueTrivia Then
                    If newTrailingTrivia.Count > 0 AndAlso newTrailingTrivia.First.IsKind(SyntaxKind.WhitespaceTrivia) Then
                        newTrailingTrivia.RemoveAt(0)
                    End If
                    If newTrailingTrivia.Last.IsKind(SyntaxKind.WhitespaceTrivia) Then
                        newTrailingTrivia.RemoveAt(newTrailingTrivia.Count - 1)
                    End If
                    If newTrailingTrivia.ContainsCommentOrDirectiveTrivia Then
                        newLeadingTrivia.AddRange(newTrailingTrivia)
                    End If
                End If
                Dim newModifierList As List(Of SyntaxToken) = Param.Modifiers.ToList
                If Param.Modifiers.Count = 1 Then
                    If Param.Identifier.GetLeadingTrivia.ContainsCommentOrDirectiveTriviaOrContinuation Then
                        newLeadingTrivia.AddRange(Param.Identifier.GetLeadingTrivia)
                    End If
                    newParamsList.Add(Param.WithModifiers(Nothing).WithIdentifier(Param.Identifier.WithLeadingTrivia(newLeadingTrivia)))
                Else
                    If indexOfByVal < Param.Modifiers.Count Then
                        Stop
                    End If
                    newParamsList.Add(Param.WithModifiers(SyntaxFactory.TokenList(newModifierList)).NormalizeWhitespace)
                End If

            Next

            editor.ReplaceNode(parameterList, parameterList.WithParameters(SyntaxFactory.SeparatedList(newParamsList, SyntaxFactory.TokenList(parameterList.Parameters.GetSeparators))))
            Return editor.GetChangedDocument
        End Function

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim FirstDiagnostic As Diagnostic = context.Diagnostics.First()
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)

            Dim diagnosticSpan As TextSpan = FirstDiagnostic.Location.SourceSpan

            Dim textSpan As TextSpan = context.Span
            Dim token As SyntaxToken = root.FindToken(textSpan.Start)
            Dim param As ParameterSyntax = CType(token.Parent, ParameterSyntax)
            Dim paramList As ParameterListSyntax = CType(param.Parent, ParameterListSyntax)
            ' Register a code action that will invoke the fix.
            Dim CodeAction As CodeAction = CodeAction.Create(
                Title,
                createChangedDocument:=Function(c) RemoveByValOccuranceAsync(context.Document, paramList, c),
                equivalenceKey:=Title)
            context.RegisterCodeFix(CodeAction, FirstDiagnostic)
        End Function

    End Class

End Namespace
