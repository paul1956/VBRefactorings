' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Composition
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(AddAsClauseRefactoringProvider)), [Shared]>
    Public Class AddAsClauseRefactoringProvider
        Inherits CodeRefactoringProvider

        Private Const Title As String = "Add As Clause refactoring"
        Private ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(
            AddAsClauseDiagnosticId,
            AddAsClauseForLambdaDiagnosticId,
            ERR_NameNotDeclared1DiagnosticId,
            ERR_StrictDisallowImplicitObjectDiagnosticId,
            ERR_StrictDisallowsImplicitProcDiagnosticId,
            ERR_EnumNotExpression1DiagnosticId,
            ERR_TypeNotExpression1DiagnosticId,
            ERR_ClassNotExpression1DiagnosticId,
            ERR_StructureNotExpression1DiagnosticId,
            ERR_InterfaceNotExpression1DiagnosticId,
            ERR_NamespaceNotExpression1DiagnosticId
            )

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim span As TextSpan = context.Span
            If Not span.IsEmpty Then
                Return
            End If
            Dim _Document As Document = context.Document
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim model As SemanticModel = Await _Document.GetSemanticModelAsync(Nothing)
            ' Don't offer this refactoring if Option Explicit because the CodeFix provider will handle it
            If model.OptionStrict = OptionStrict.On AndAlso model.OptionInfer = True Then Return
            ' Find the node at the selection.

            For Each DiagnosticEntry As Diagnostic In model.GetDiagnostics(context.Span)
                ' Does Diagnostic contain any of the errors I care about?
                If FixableDiagnosticIds.Any(Function(m As String) m.ToString = DiagnosticEntry.Id.ToString) Then
                    Continue For
                End If
                Return
            Next

            Dim _VariableDeclarator As VariableDeclaratorSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of VariableDeclaratorSyntax)()
            If _VariableDeclarator IsNot Nothing Then
                ' only for a local declaration node
                If _VariableDeclarator.Names.Count > 1 Then
                    Return
                End If
                If _VariableDeclarator.AsClause IsNot Nothing Then
                    Return
                End If
                If _VariableDeclarator.Initializer Is Nothing Then
                    Return
                End If
                context.RegisterRefactoring(CodeAction.Create(Title, Function(c As CancellationToken) AddAsClauseDocumentAsync(root, model, _Document, _VariableDeclarator, c)))
                Exit Function
            End If
            Dim ForStatement As ForStatementSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of ForStatementSyntax)()
            If ForStatement IsNot Nothing Then
                Dim ControlVariable As VisualBasicSyntaxNode = ForStatement.ControlVariable
                If ControlVariable Is Nothing Then Return

                Dim ControlVariableName As IdentifierNameSyntax = TryCast(ControlVariable, IdentifierNameSyntax)
                If ControlVariableName?.Kind = SyntaxKind.IdentifierName Then
                    context.RegisterRefactoring(CodeAction.Create(Title, Function(c As CancellationToken) AddAsClauseDocumentAsync(_Document, ForStatement, c)))
                End If
                Exit Function
            End If
            Dim ForEachStatement As ForEachStatementSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of ForEachStatementSyntax)()
            If ForEachStatement Is Nothing Then
                Return
            End If

            Dim ForEachControlVariable As VisualBasicSyntaxNode = ForEachStatement.ControlVariable
            If ForEachControlVariable Is Nothing Then Return
            Dim IdentifierName As IdentifierNameSyntax = TryCast(ForEachControlVariable, IdentifierNameSyntax)
            If IdentifierName?.Kind = SyntaxKind.IdentifierName Then
                context.RegisterRefactoring(CodeAction.Create(Title, Function(c As CancellationToken) AddAsClauseDocumentAsync(_Document, ForEachStatement, c)))
            End If
        End Function

    End Class

End Namespace
