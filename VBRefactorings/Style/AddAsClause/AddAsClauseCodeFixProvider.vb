' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable
Imports System.Composition
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(AddAsClauseCodeFixProvider)), [Shared]>
    Partial Public Class AddAsClauseCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(AddAsClauseDiagnosticId,
                                             AddAsClauseForLambdaDiagnosticId,
                                             ChangeAsObjectToMoreSpecificDiagnosticId,
                                             ERR_ClassNotExpression1DiagnosticId,
                                             ERR_EnumNotExpression1DiagnosticId,
                                             ERR_InterfaceNotExpression1DiagnosticId,
                                             ERR_NameNotDeclared1DiagnosticId,
                                             ERR_NamespaceNotExpression1DiagnosticId,
                                             ERR_StrictDisallowImplicitObjectDiagnosticId,
                                             ERR_StrictDisallowsImplicitProcDiagnosticId,
                                             ERR_StructureNotExpression1DiagnosticId,
                                             ERR_TypeNotExpression1DiagnosticId
                                             )
            End Get
        End Property

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides Async Function RegisterCodeFixesAsync(Context As CodeFixContext) As Task
            Try
                Dim Root As SyntaxNode = Await Context.Document.GetSyntaxRootAsync(Context.CancellationToken).ConfigureAwait(False)
                Dim diagnostic As Diagnostic = Context.Diagnostics.First()
                Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan
                Dim DiagnosticSpanStart As Integer = diagnostic.Location.SourceSpan.Start
                Dim VariableDeclaration As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(VariableDeclaratorSyntax))
                Dim Model As SemanticModel = Await Context.Document.GetSemanticModelAsync(Context.CancellationToken)
                If VariableDeclaration IsNot Nothing Then
                    For Each DiagnosticEntry As Diagnostic In Model.GetDiagnostics(VariableDeclaration.Span)
                        ' Does Diagnostic contain any of the errors I care about?
                        If FixableDiagnosticIds.Any(Function(m As String) m.ToString = DiagnosticEntry.Id.ToString) Then
                            Exit For
                        End If
                        Exit Function
                    Next
                    Dim PossibleVariableDeclaratorSyntax As VariableDeclaratorSyntax = CType(VariableDeclaration, VariableDeclaratorSyntax)
                    If PossibleVariableDeclaratorSyntax IsNot Nothing Then
                        If PossibleVariableDeclaratorSyntax.AsClause Is Nothing OrElse PossibleVariableDeclaratorSyntax.AsClause.GetText.ToString.Trim = "As Object" Then
                            Dim newVariable As SyntaxNode = AddAsClauseAsync(Root, Model, Context.Document, DirectCast(VariableDeclaration, VariableDeclaratorSyntax))
                            Dim action As AddAsClauseCodeFixProviderCodeAction = New AddAsClauseCodeFixProviderCodeAction("Add As Clause", Context.Document, VariableDeclaration, newVariable)
                            Context.RegisterCodeFix(action, Context.Diagnostics)
                            Exit Function
                        End If
                    End If
                End If
                Dim LambdaHeader As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(LambdaHeaderSyntax))
                If LambdaHeader IsNot Nothing Then
                    Dim newLambdaStatement As SyntaxNode = AddAsClauseAsync(Model, Context.Document, DirectCast(LambdaHeader, LambdaHeaderSyntax))
                    Context.RegisterCodeFix(New AddAsClauseCodeFixProviderCodeAction("Add As Clause", Context.Document, LambdaHeader, newLambdaStatement), Context.Diagnostics)
                    Exit Function
                End If
                Dim ForStatement As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(ForStatementSyntax))
                If ForStatement IsNot Nothing Then
                    Dim newForStatement As SyntaxNode = Await AddAsClauseAsync(Context.Document, DirectCast(ForStatement, ForStatementSyntax), Context.CancellationToken)
                    Context.RegisterCodeFix(New AddAsClauseCodeFixProviderCodeAction("Add As Clause", Context.Document, ForStatement, newForStatement), Context.Diagnostics)
                    Exit Function
                End If
                Dim ForEachStatement As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(ForEachStatementSyntax))
                If ForEachStatement IsNot Nothing Then
                    Dim newForEachStatement As SyntaxNode = Await AddAsClauseAsync(Context.Document, DirectCast(ForEachStatement, ForEachStatementSyntax), Context.CancellationToken)
                    Context.RegisterCodeFix(New AddAsClauseCodeFixProviderCodeAction("Add As Clause", Context.Document, ForEachStatement, newForEachStatement), Context.Diagnostics)
                End If
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Function

    End Class

End Namespace
