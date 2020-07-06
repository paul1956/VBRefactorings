' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:="RemoveByValVB"), [Shared]>
    Partial Class ByValCodeRefactoringProvider
        Inherits CodeRefactoringProvider

        Private Async Function RemoveAllOccurancesAsync(document As Document, CancelToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
            Dim rewriter As ByValCodeRefactoringProviderRewriter = New ByValCodeRefactoringProviderRewriter(Function(current As SyntaxToken) True)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Private Async Function RemoveOccuranceAsync(document As Document, token As SyntaxToken, CancelToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
            Dim rewriter As ByValCodeRefactoringProviderRewriter = New ByValCodeRefactoringProviderRewriter(Function(t As SyntaxToken) t = token)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim CurrentDocument As Document = context.Document
            Dim CurrentTextSpan As TextSpan = context.Span
            Dim CancelToken As CancellationToken = context.CancellationToken

            Dim root As SyntaxNode = Await CurrentDocument.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
            Dim token As SyntaxToken = root.FindToken(CurrentTextSpan.Start)
            Dim parameterListNode As SyntaxNode = token.Parent

            If token.Kind = SyntaxKind.ByValKeyword AndAlso token.Span.IntersectsWith(CurrentTextSpan.Start) Then
                context.RegisterRefactoring(New ByValAnalyzerFixCodeFixProvider.ByValAnalyzerFixCodeFixProviderCodeAction("Remove unnecessary ByVal keyword",
                                                                  CType(Function(c As CancellationToken) RemoveOccuranceAsync(CurrentDocument, token, c), Func(Of Object, Task(Of Document)))))
                context.RegisterRefactoring(New ByValAnalyzerFixCodeFixProvider.ByValAnalyzerFixCodeFixProviderCodeAction("Remove all occurrences of unnecessary ByVal keywords",
                                                                  CType(Function(c As CancellationToken) RemoveAllOccurancesAsync(CurrentDocument, c), Func(Of Object, Task(Of Document)))))
            End If
        End Function

    End Class

End Namespace
