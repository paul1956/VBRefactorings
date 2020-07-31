Option Infer Off

Imports Microsoft.CodeAnalysis.Formatting

Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:="Convert 'If' to 'Select Case'")>
    Public Class ConvertIfStatementToSelectCaseStatementCodeRefactoringProvider
        Inherits CodeRefactoringProvider

        Private Shared ReadOnly validTypes() As SpecialType = {SpecialType.System_String, SpecialType.System_Boolean, SpecialType.System_Char, SpecialType.System_Byte, SpecialType.System_SByte, SpecialType.System_Int16, SpecialType.System_Int32, SpecialType.System_Int64, SpecialType.System_UInt16, SpecialType.System_UInt32, SpecialType.System_UInt64}

        Private Shared Function CollectCaseLabels(ByVal result As List(Of CaseClauseSyntax), ByVal context As SemanticModel, ByVal condition As ExpressionSyntax, ByVal switchExpr As ExpressionSyntax) As Boolean
            If TypeOf condition Is ParenthesizedExpressionSyntax Then
                Return CollectCaseLabels(result, context, CType(condition, ParenthesizedExpressionSyntax).Expression, switchExpr)
            End If

            Dim binaryOp As BinaryExpressionSyntax = TryCast(condition, BinaryExpressionSyntax)
            If binaryOp Is Nothing Then
                Return False
            End If

            If binaryOp.IsKind(SyntaxKind.OrExpression) OrElse binaryOp.IsKind(SyntaxKind.OrElseExpression) Then
                Return CollectCaseLabels(result, context, binaryOp.Left, switchExpr) AndAlso CollectCaseLabels(result, context, binaryOp.Right, switchExpr)
            End If

            If binaryOp.IsKind(SyntaxKind.EqualsExpression, SyntaxKind.IsExpression) Then
                If switchExpr.IsEquivalentTo(binaryOp.Left, True) Then
                    If IsConstantExpression(context, binaryOp.Right) Then
                        result.Add(SyntaxFactory.SimpleCaseClause(binaryOp.Right))
                        Return True
                    End If
                ElseIf switchExpr.IsEquivalentTo(binaryOp.Right, True) Then
                    If IsConstantExpression(context, binaryOp.Left) Then
                        result.Add(SyntaxFactory.SimpleCaseClause(binaryOp.Left))
                        Return True
                    End If
                End If
            End If

            Return False
        End Function

        Private Shared Function IsConstantExpression(ByVal context As SemanticModel, ByVal expr As ExpressionSyntax) As Boolean
            If TypeOf expr Is LiteralExpressionSyntax Then
                Return True
            End If
            Return context.GetConstantValue(expr).HasValue
        End Function

        Private Shared Function IsValidSwitchType(ByVal type As ITypeSymbol) As Boolean
            If type Is Nothing OrElse TypeOf type Is IErrorTypeSymbol Then
                Return False
            End If
            If type.TypeKind = TypeKind.Enum Then
                Return True
            End If

            If type.IsNullableType() Then
                type = type.GetNullableUnderlyingType()
                If type Is Nothing OrElse TypeOf type Is IErrorTypeSymbol Then
                    Return False
                End If
            End If
            Return Array.IndexOf(validTypes, type.SpecialType) <> -1
        End Function

        Friend Shared Function CollectCaseBlocks(ByVal result As List(Of CaseBlockSyntax), ByVal context As SemanticModel, ByVal ifBlock As MultiLineIfBlockSyntax, ByVal switchExpr As ExpressionSyntax) As Boolean
            ' if
            Dim labels As List(Of CaseClauseSyntax) = New List(Of CaseClauseSyntax)
            If Not CollectCaseLabels(labels, context, ifBlock.IfStatement.Condition, switchExpr) Then
                Return False
            End If

            result.Add(SyntaxFactory.CaseBlock(SyntaxFactory.CaseStatement(labels.ToArray()), ifBlock.Statements))

            For Each block As ElseIfBlockSyntax In ifBlock.ElseIfBlocks
                labels = New List(Of CaseClauseSyntax)()
                If Not CollectCaseLabels(labels, context, block.ElseIfStatement.Condition, switchExpr) Then
                    Return False
                End If

                result.Add(SyntaxFactory.CaseBlock(SyntaxFactory.CaseStatement(labels.ToArray()), block.Statements))
            Next block

            ' else
            If ifBlock.ElseBlock IsNot Nothing Then
                result.Add(SyntaxFactory.CaseElseBlock(SyntaxFactory.CaseElseStatement(SyntaxFactory.ElseCaseClause()), ifBlock.ElseBlock.Statements))
            End If
            Return True
        End Function

        Friend Shared Function GetSelectCaseExpression(ByVal context As SemanticModel, ByVal expr As ExpressionSyntax) As ExpressionSyntax
            Dim binaryOp As BinaryExpressionSyntax = TryCast(expr, BinaryExpressionSyntax)
            If binaryOp Is Nothing Then
                Return Nothing
            End If

            If binaryOp.OperatorToken.IsKind(SyntaxKind.OrElseKeyword, SyntaxKind.OrKeyword) Then
                Return GetSelectCaseExpression(context, binaryOp.Left)
            End If

            If binaryOp.OperatorToken.IsKind(SyntaxKind.EqualsToken) Then
                Dim switchExpr As ExpressionSyntax = Nothing
                If IsConstantExpression(context, binaryOp.Right) Then
                    switchExpr = binaryOp.Left
                End If
                If IsConstantExpression(context, binaryOp.Left) Then
                    switchExpr = binaryOp.Right
                End If
                If switchExpr IsNot Nothing AndAlso IsValidSwitchType(context.GetTypeInfo(switchExpr).Type) Then
                    Return switchExpr
                End If
            End If

            Return Nothing
        End Function

        Public Overrides Async Function ComputeRefactoringsAsync(ByVal context As CodeRefactoringContext) As Task
            Dim document As Document = context.Document
            If document.Project.Solution.Workspace.Kind = WorkspaceKind.MiscellaneousFiles Then
                Return
            End If
            Dim span As TextSpan = context.Span
            If Not span.IsEmpty Then
                Return
            End If
            Dim cancellationToken As CancellationToken = context.CancellationToken
            If cancellationToken.IsCancellationRequested Then
                Return
            End If
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim model As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken).ConfigureAwait(False)
            If model.IsFromGeneratedCode(cancellationToken) Then
                Return
            End If
            Dim node As IfStatementSyntax = TryCast(root.FindNode(span), IfStatementSyntax)

            If node Is Nothing OrElse Not (TypeOf node.Parent Is MultiLineIfBlockSyntax) Then
                Return
            End If

            Dim ifBlock As MultiLineIfBlockSyntax = TryCast(node.Parent, MultiLineIfBlockSyntax)

            Dim selectCaseExpression As ExpressionSyntax = GetSelectCaseExpression(model, node.Condition)
            If selectCaseExpression Is Nothing Then
                Return
            End If

            Dim caseBlocks As New List(Of CaseBlockSyntax)
            If Not CollectCaseBlocks(caseBlocks, model, ifBlock, selectCaseExpression) Then
                Return
            End If

            context.RegisterRefactoring(New DocumentChangeAction(
                                                                span,
                                                                DiagnosticSeverity.Info,
                                                                "To 'Select Case'",
                                                                Function(ct As CancellationToken)
                                                                    Dim selectCaseStatement As SelectBlockSyntax = SyntaxFactory.SelectBlock(SyntaxFactory.SelectStatement(selectCaseExpression).WithCaseKeyword(SyntaxFactory.Token(SyntaxKind.CaseKeyword)), (New SyntaxList(Of CaseBlockSyntax)()).AddRange(caseBlocks)).NormalizeWhitespace()
                                                                    Return Task.FromResult(document.WithSyntaxRoot(root.ReplaceNode(ifBlock, selectCaseStatement.WithLeadingTrivia(ifBlock.GetLeadingTrivia()).WithTrailingTrivia(ifBlock.GetTrailingTrivia()).WithAdditionalAnnotations(Formatter.Annotation))))
                                                                End Function)
                                                                )
        End Function
    End Class

End Namespace