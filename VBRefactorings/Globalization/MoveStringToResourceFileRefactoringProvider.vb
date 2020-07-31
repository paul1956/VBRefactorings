' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Composition
Imports System.Diagnostics.Debug
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Globalization

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(MoveStringToResourceFileRefactoringProvider)), [Shared]>
    Partial Public Class MoveStringToResourceFileRefactoringProvider
        Inherits CodeRefactoringProvider
        Private _resourceClass As ResourceXClass

        Private Async Function MoveStringToResourceFileAsync(Statement As StatementSyntax, invocation As ExpressionSyntax, document As Document, cancellationToken As CancellationToken) As Task(Of Document)
            Assert(Not invocation Is Nothing)
            Try
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
                Dim ResourceText As String = ""

                If invocation.Kind = SyntaxKind.StringLiteralExpression Then
                    Dim InvicationRightToken As SyntaxToken = CType(invocation, LiteralExpressionSyntax).Token
                    ResourceText = InvicationRightToken.ValueText.Replace("""", """""")
                Else
                    Stop
                End If

                Dim identifierNameString As String = _resourceClass.GetUniqueResourceName(ResourceText)
                Dim ModifiedIdentifier As ModifiedIdentifierSyntax = SyntaxFactory.ModifiedIdentifier(identifierNameString)
                Dim ResourceRetriever As ExpressionSyntax = SyntaxFactory.ParseExpression($" ResourceRetriever.GetString(""{identifierNameString}"")")
                Dim NewEqualsValue As EqualsValueSyntax = SyntaxFactory.EqualsValue(ResourceRetriever)

                Dim newRoot As SyntaxNode = Nothing
                Select Case Statement.Kind
                    Case SyntaxKind.ExpressionStatement
                        newRoot = root.ReplaceNode(invocation, ResourceRetriever.WithoutTrivia)
                    Case SyntaxKind.FieldDeclaration
                        Dim FieldDeclaration As FieldDeclarationSyntax = CType(Statement, FieldDeclarationSyntax)
                        Dim VariableDefinationLine As New SeparatedSyntaxList(Of VariableDeclaratorSyntax)
                        VariableDefinationLine = SyntaxFactory.SingletonSeparatedList(FieldDeclaration.Declarators(0).WithInitializer(NewEqualsValue))
                        Dim NewFieldStatement As SyntaxNode = SyntaxFactory.FieldDeclaration(FieldDeclaration.AttributeLists, FieldDeclaration.Modifiers, VariableDefinationLine).WithTriviaFrom(Statement)
                        newRoot = root.ReplaceNode(Statement, NewFieldStatement)
                    Case SyntaxKind.LocalDeclarationStatement
                        Dim SimpleAssignmentStatement As LocalDeclarationStatementSyntax = CType(Statement, LocalDeclarationStatementSyntax)
                        Dim declarators As New SeparatedSyntaxList(Of VariableDeclaratorSyntax)()
                        declarators = declarators.Add(SimpleAssignmentStatement.Declarators(0).WithInitializer(NewEqualsValue))
                        Dim newLocalDeclaration As LocalDeclarationStatementSyntax = Microsoft.CodeAnalysis.SyntaxNodeExtensions.WithTriviaFrom(SyntaxFactory.LocalDeclarationStatement(SimpleAssignmentStatement.Modifiers, declarators), Statement)
                        newRoot = root.ReplaceNode(Statement, newLocalDeclaration)
                    Case SyntaxKind.PropertyStatement
                        Dim NewPropertyStatement As SyntaxNode = CType(Statement, PropertyStatementSyntax).WithInitializer(NewEqualsValue).WithTriviaFrom(Statement)
                        newRoot = root.ReplaceNode(Statement, NewPropertyStatement)
                    Case Else

                End Select
                If _resourceClass.AddToResourceFile(identifierNameString, ResourceText) Then
                    Return document.WithSyntaxRoot(newRoot)
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return Nothing
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim semanticModel As SemanticModel = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken)
            Dim invocation As ExpressionSyntax = root.FindNode(context.Span, getInnermostNodeForTie:=True)?.FirstAncestorOrSelf(Of LiteralExpressionSyntax)()
            If invocation Is Nothing Then
                Exit Function
            End If
            If _resourceClass Is Nothing Then
                If context.Document.Project.FilePath Is Nothing Then
                    _resourceClass = New ResourceXClass()
                Else
                    _resourceClass = New ResourceXClass(context.Document.Project.FilePath)
                End If
            End If
            If Me._resourceClass.Initialized = ResourceXClass.InitializedValues.NoResourceDirectory Then
                Exit Function
            End If
            Select Case invocation.Kind
                Case SyntaxKind.StringLiteralExpression
                    Dim Statement As StatementSyntax = invocation.FirstAncestorOfType(Of StatementSyntax)
                    If TypeOf Statement Is LocalDeclarationStatementSyntax AndAlso CType(Statement, LocalDeclarationStatementSyntax).Declarators.Count <> 1 Then
                        Exit Function
                    End If
                    If TypeOf Statement Is MethodStatementSyntax Then
                        Exit Function
                    End If
                    context.RegisterRefactoring(New MoveStringToResourceFileCodeAction("Move String to Resource File",
                                                Function(c As CancellationToken) Me.MoveStringToResourceFileAsync(Statement, invocation, context.Document, c)))
                Case SyntaxKind.NothingLiteralExpression
                    Exit Function
                Case SyntaxKind.CharacterLiteralExpression
                    Exit Function
                Case Else
                    Stop
                    Exit Function
            End Select
        End Function

    End Class

End Namespace
