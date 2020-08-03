' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(RemoveAsClauseRefactoringProvider)), [Shared]>
    Public Class RemoveAsClauseRefactoringProvider
        Inherits CodeRefactoringProvider

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim span As TextSpan = context.Span
            If Not span.IsEmpty Then
                Return
            End If
            Dim document As Document = context.Document
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim SemanticModel As SemanticModel = Await document.GetSemanticModelAsync(Nothing)
            ' Don't allow Implicit of Option Explicit ON
            If SemanticModel.OptionExplicit = True Then Return
            ' Find the node at the selection.

            ' only for a local declaration node
            Dim VariableDeclarator As VariableDeclaratorSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of VariableDeclaratorSyntax)()
            If VariableDeclarator Is Nothing Then
                Return
            End If

            If VariableDeclarator.Names.Count > 1 Then
                Return
            End If
            If VariableDeclarator.AsClause Is Nothing Then
                Return
            End If
            If VariableDeclarator.AsClause.Kind = SyntaxKind.AsNewClause Then
                Return
            End If
            If VariableDeclarator.Initializer Is Nothing Then
                Return
            End If
            ' Should be also to do this but it is more complicated
            If VariableDeclarator.Initializer.Value.Kind = SyntaxKind.CollectionInitializer Then
                Return
            End If
            context.RegisterRefactoring(CodeAction.Create($"Remove As Clause", Function(c As CancellationToken) Me.MakeImplicit(document, VariableDeclarator, c)))
        End Function

        Private Async Function MakeImplicit(document As Document, VariableDeclarator As VariableDeclaratorSyntax, cancellationToken As CancellationToken) As Task(Of Document)
            Dim SemanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)

            Dim asClause As SimpleAsClauseSyntax = CType(VariableDeclarator.AsClause, SimpleAsClauseSyntax)
            Dim variableTypeName As TypeSyntax = asClause.Type

            Dim variableType As ITypeSymbol = CType(SemanticModel.GetSymbolInfo(variableTypeName).Symbol, ITypeSymbol)

            Dim initializerInfo As TypeInfo = SemanticModel.GetTypeInfo(VariableDeclarator.Initializer.Value)

            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken)
            If Equals(variableType, initializerInfo.Type) Then
                Dim newDeclator As VariableDeclaratorSyntax = VariableDeclarator.WithAsClause(Nothing)
                Dim newRoot As SyntaxNode = root.ReplaceNode(VariableDeclarator, newDeclator)
                Return document.WithSyntaxRoot(newRoot)
            End If
            Return document
        End Function

    End Class

End Namespace
