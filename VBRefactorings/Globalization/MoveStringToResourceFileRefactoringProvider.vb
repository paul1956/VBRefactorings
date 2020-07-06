' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace InternationalizationResources

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(MovePropertyAssignmentStringToResourceFileCodeFixProvider)), [Shared]>
    Public Class MovePropertyAssignmentStringToResourceFileCodeFixProvider
        Inherits CodeFixProvider

        Private ReadOnly _resourceClasses As New Dictionary(Of String, ResourceXClass)
        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create("CA1303")

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Async Function MoveStringToResourceFileAsync(CurrentDocument As Document, diagnostic As Diagnostic, CancelToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await CurrentDocument.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)
            Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan
            Dim Expression As ExpressionSyntax = root.FindToken(diagnosticSpan.Start).Parent.AncestorsAndSelf().OfType(Of ExpressionSyntax).First()
            Dim ResourceClass As ResourceXClass = InitializeResourceClasses(_resourceClasses, CurrentDocument)
            If ResourceClass.Initialized = ResourceXClass.InitializedValues.NoResourceDirectory Then
                Return CurrentDocument
            End If
            If Expression IsNot Nothing Then
                Dim Rewriter As New VisualBasicConstantStringReplacement(ResourceClass)
                Dim NewRoot As SyntaxNode = root.ReplaceNode(Expression, Rewriter.Visit(Expression))
                Return CurrentDocument.WithSyntaxRoot(NewRoot)
            Else
                Stop
            End If

            Return CurrentDocument
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            context.RegisterCodeFix(
                CodeAction.Create(
                                  "Move string to resource file.",
                                  Function(c As CancellationToken) MoveStringToResourceFileAsync(context.Document, context.Diagnostics.First(), c),
                                  equivalenceKey:=NameOf(MoveStringToResourceFileAsync)
                                 ),
                context.Diagnostics.First())
            Return Task.FromResult(0)
        End Function

    End Class

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:=NameOf(MoveStringToResourceFileRefactoringProvider)), [Shared]>
    Public Class MoveStringToResourceFileRefactoringProvider
        Inherits CodeRefactoringProvider
        Private ReadOnly _resourceClasses As New Dictionary(Of String, ResourceXClass)

        Private Async Function MoveStringToResourceFileAsync(Expression As ExpressionSyntax, document As Document, CancelToken As CancellationToken) As Task(Of Document)
            Try
                Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(CancelToken).ConfigureAwait(False)

                Dim ResourceClass As ResourceXClass = _resourceClasses(document.Project.FilePath & "")
                Dim newRoot As SyntaxNode = MoveStringToResourceFile(ResourceClass, root, Expression)
                If newRoot IsNot Nothing Then
                    Return document.WithSyntaxRoot(newRoot)
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return Nothing
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim CurrentDocument As Document = context.Document
            Dim semanticModel As SemanticModel = Await CurrentDocument.GetSemanticModelAsync(context.CancellationToken)
            Dim root As SyntaxNode = Await CurrentDocument.GetSyntaxRootAsync(context.CancellationToken)
            Dim StringSyntaxNode As SyntaxNode = root.FindNode(context.Span, getInnermostNodeForTie:=True)
            Dim invocation As ExpressionSyntax = StringSyntaxNode?.FirstAncestorOrSelf(Of LiteralExpressionSyntax)()
            If invocation Is Nothing Then
                invocation = StringSyntaxNode?.FirstAncestorOrSelf(Of InterpolatedStringExpressionSyntax)()
                If invocation Is Nothing Then
                    Exit Function
                End If
            End If
            Dim ResourceClass As ResourceXClass = InitializeResourceClasses(_resourceClasses, CurrentDocument)
            If ResourceClass.Initialized = ResourceXClass.InitializedValues.NoResourceDirectory Then
                Exit Function
            End If
            Select Case invocation.Kind
                Case SyntaxKind.StringLiteralExpression, SyntaxKind.InterpolatedStringExpression
                    Dim Statement As StatementSyntax = invocation.FirstAncestorOfType(Of StatementSyntax)
                    If TypeOf Statement Is LocalDeclarationStatementSyntax AndAlso CType(Statement, LocalDeclarationStatementSyntax).Declarators.Count <> 1 Then
                        Exit Function
                    End If
                    If TypeOf Statement Is MethodStatementSyntax Then
                        Exit Function
                    End If
                    context.RegisterRefactoring(New MoveStringToResourceFileCodeAction("Move String to Resource File",
                                                Function(c As CancellationToken) MoveStringToResourceFileAsync(invocation, CurrentDocument, c)))
                Case SyntaxKind.FalseLiteralExpression, SyntaxKind.TrueLiteralExpression, SyntaxKind.CharacterLiteralExpression, SyntaxKind.NothingLiteralExpression
                    Exit Function
                Case Else
                    Stop
                    Exit Function
            End Select
        End Function

        Private Class MoveStringToResourceFileCodeAction
            Inherits CodeAction

            Private ReadOnly _title As String
            Private ReadOnly _generateDocument As Func(Of CancellationToken, Task(Of Document))

            Public Sub New(title As String, generateDocument As Func(Of CancellationToken, Task(Of Document)))
                _title = title
                _generateDocument = generateDocument
            End Sub

            Public Overrides ReadOnly Property Title As String
                Get
                    Return _title
                End Get
            End Property

            Protected Overrides Function GetChangedDocumentAsync(CancelToken As CancellationToken) As Task(Of Document)
                Return _generateDocument(CancelToken)
            End Function

        End Class

    End Class

End Namespace
