' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Style

    Partial Friend Class ConcatenateExpressionToInterpolatedStringRefactoringProvider

        Private Class ConcatenateExpressionToInterpolatedStringRefactoringProviderRewriter
            Inherits VisualBasicSyntaxRewriter
            Private ReadOnly _expandedArguments As ImmutableArray(Of ExpressionSyntax)

            Private Sub New(expandedArguments As ImmutableArray(Of ExpressionSyntax))
                _expandedArguments = expandedArguments
            End Sub

            Public Overloads Shared Function Visit(interpolatedString As InterpolatedStringExpressionSyntax, expandedArguments As ImmutableArray(Of ExpressionSyntax)) As InterpolatedStringExpressionSyntax
                Return DirectCast(New ConcatenateExpressionToInterpolatedStringRefactoringProviderRewriter(expandedArguments).Visit(interpolatedString), InterpolatedStringExpressionSyntax)
            End Function

            Public Overrides Function VisitInterpolation(node As InterpolationSyntax) As SyntaxNode
                Dim literalExpression As LiteralExpressionSyntax = TryCast(node.Expression, LiteralExpressionSyntax)
                If literalExpression IsNot Nothing AndAlso literalExpression.IsKind(SyntaxKind.NumericLiteralExpression) Then
                    Dim index As Integer = CInt(literalExpression.Token.Value)
                    If index >= 0 AndAlso index < _expandedArguments.Length Then
                        Return node.WithExpression(_expandedArguments(index))
                    End If
                End If

                Return MyBase.VisitInterpolation(node)
            End Function

        End Class

    End Class

End Namespace
