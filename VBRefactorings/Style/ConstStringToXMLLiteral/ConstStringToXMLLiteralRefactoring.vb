' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(ConstStringToXMLLiteralRefactoring)), [Shared]>
    Public Class ConstStringToXMLLiteralRefactoring
        Inherits CodeRefactoringProvider
        Const DoubleQuote As Char = """"c

        Private Async Function ConvertToXMLLiteralAsync(document As Document, localDeclaration As LocalDeclarationStatementSyntax, lVariableDeclarator As VariableDeclaratorSyntax, c As CancellationToken) As Task(Of Document)
            Dim newRoot As SyntaxNode

            Try
                Dim declaration As VariableDeclaratorSyntax = localDeclaration.Declarators.First
                Dim ConstModifier As SyntaxToken = localDeclaration.Modifiers.First()

                Dim dimModifier As SyntaxToken = SyntaxFactory.Token(SyntaxKind.DimKeyword).
                WithLeadingTrivia(dimModifier.LeadingTrivia).
                WithTrailingTrivia(dimModifier.TrailingTrivia).
                WithAdditionalAnnotations(Formatter.Annotation)

                Dim modifiers As SyntaxTokenList = localDeclaration.Modifiers.Replace(ConstModifier, dimModifier)

                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(c).ConfigureAwait(False)
                Dim value As String = lVariableDeclarator.Initializer.Value.ToString
                Dim Value1 As String = value.ToString.Substring(1, value.Length - 2).Replace(New String(DoubleQuote, 2), DoubleQuote)
                Dim SimpleMemberAccessExpressionSyntax As ExpressionSyntax = SyntaxFactory.ParseExpression($"<text>{Value1}</text>.Value")
                Dim NewVariableDeclarator As VariableDeclaratorSyntax = lVariableDeclarator.ReplaceNode(lVariableDeclarator.Initializer, lVariableDeclarator.Initializer.WithValue(SimpleMemberAccessExpressionSyntax))
                Dim lSeparatedSyntaxList As New SeparatedSyntaxList(Of VariableDeclaratorSyntax)
                lSeparatedSyntaxList = SyntaxFactory.SingletonSeparatedList(NewVariableDeclarator)

                Dim newLocalDeclaration As LocalDeclarationStatementSyntax = localDeclaration.
            WithDeclarators(lSeparatedSyntaxList).
            WithModifiers(modifiers).
            WithLeadingTrivia(localDeclaration.GetLeadingTrivia()).
            WithTrailingTrivia(localDeclaration.GetTrailingTrivia()).
            WithAdditionalAnnotations(Formatter.Annotation)
                newRoot = root.ReplaceNode(localDeclaration, newLocalDeclaration)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim span As TextSpan = context.Span
            If Not span.IsEmpty OrElse context.Span.Start = 0 Then
                Return
            End If
            Dim document As Document = context.Document
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim _VariableDeclarator As VariableDeclaratorSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of VariableDeclaratorSyntax)()
            If _VariableDeclarator Is Nothing Then
                Return
            End If
            If _VariableDeclarator.Names.Count > 1 Then
                Return
            End If
            Dim EqualsValue As EqualsValueSyntax = _VariableDeclarator.Initializer
            If EqualsValue Is Nothing Then
                Return
            End If

            If EqualsValue.Value.Kind <> SyntaxKind.StringLiteralExpression Then
                Return
            End If

            Dim syntaxToken1 As SyntaxToken = root.FindToken(context.Span.Start)
            If syntaxToken1 = Nothing Then
                Return
            End If
            Dim syntaxToken1Parent As SyntaxNode = syntaxToken1.Parent
            If syntaxToken1Parent Is Nothing Then
                Return
            End If
            ' only for a local declaration node
            Dim PossibleLocalDeclarationStatementSyntax As IEnumerable(Of LocalDeclarationStatementSyntax) = syntaxToken1Parent.AncestorsAndSelf().OfType(Of LocalDeclarationStatementSyntax)
            If PossibleLocalDeclarationStatementSyntax Is Nothing Then
                Return
            End If
            If PossibleLocalDeclarationStatementSyntax.Count = 0 Then
                Return
            End If
            Dim localDeclaration As LocalDeclarationStatementSyntax = PossibleLocalDeclarationStatementSyntax.First()
            If localDeclaration Is Nothing Then
                Return
            End If
            Dim ConstModifier As SyntaxToken = localDeclaration.Modifiers.First()
            If ConstModifier.Text <> "Const" Then
                Return
            End If

            context.RegisterRefactoring(CodeAction.Create($"Convert Const String to XML Literal", Function(c As CancellationToken) Me.ConvertToXMLLiteralAsync(document, localDeclaration, _VariableDeclarator, c)))
            Exit Function
        End Function

    End Class

End Namespace
