' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Friend Module VBUtil

    ''' <summary>
    ''' When negating an expression this is required, otherwise you would end up with
    ''' a or b -> !a or b
    ''' </summary>
    Public Function AddParensForUnaryExpressionIfRequired(expression As ExpressionSyntax) As ExpressionSyntax
        If (TypeOf expression Is BinaryExpressionSyntax) OrElse (TypeOf expression Is CastExpressionSyntax) OrElse (TypeOf expression Is LambdaExpressionSyntax) Then
            Return SyntaxFactory.ParenthesizedExpression(expression)
        End If

        Return expression
    End Function

    Public Function GetExpressionOperatorTokenKind(op As SyntaxKind) As SyntaxKind
        Select Case op
            Case SyntaxKind.EqualsExpression
                Return SyntaxKind.EqualsToken
            Case SyntaxKind.NotEqualsExpression
                Return SyntaxKind.LessThanGreaterThanToken
            Case SyntaxKind.GreaterThanExpression
                Return SyntaxKind.GreaterThanToken
            Case SyntaxKind.GreaterThanOrEqualExpression
                Return SyntaxKind.GreaterThanEqualsToken
            Case SyntaxKind.LessThanExpression
                Return SyntaxKind.LessThanToken
            Case SyntaxKind.LessThanOrEqualExpression
                Return SyntaxKind.LessThanEqualsToken
            Case SyntaxKind.OrExpression
                Return SyntaxKind.OrKeyword
            Case SyntaxKind.OrElseExpression
                Return SyntaxKind.OrElseKeyword
            Case SyntaxKind.AndExpression
                Return SyntaxKind.AndKeyword
            Case SyntaxKind.AndAlsoExpression
                Return SyntaxKind.AndAlsoKeyword
            Case SyntaxKind.AddExpression
                Return SyntaxKind.PlusToken
            Case SyntaxKind.ConcatenateExpression
                Return SyntaxKind.AmpersandToken
            Case SyntaxKind.SubtractExpression
                Return SyntaxKind.MinusToken
            Case SyntaxKind.MultiplyExpression
                Return SyntaxKind.AsteriskToken
            Case SyntaxKind.DivideExpression
                Return SyntaxKind.SlashToken
            Case SyntaxKind.ModuloExpression
                Return SyntaxKind.ModKeyword
                ' assignments
            Case SyntaxKind.SimpleAssignmentStatement
                Return SyntaxKind.EqualsToken
            Case SyntaxKind.AddAssignmentStatement
                Return SyntaxKind.PlusEqualsToken
            Case SyntaxKind.SubtractAssignmentStatement
                Return SyntaxKind.MinusEqualsToken
                ' unary
            Case SyntaxKind.UnaryPlusExpression
                Return SyntaxKind.PlusToken
            Case SyntaxKind.UnaryMinusExpression
                Return SyntaxKind.MinusToken
            Case SyntaxKind.NotExpression
                Return SyntaxKind.NotKeyword
        End Select
        Throw New ArgumentOutOfRangeException(NameOf(op))
    End Function

    ''' <summary>
    ''' Inverts a boolean condition. Note: The condition object can be frozen (from AST) it's cloned internally.
    ''' </summary>
    ''' <param name="condition">The condition to invert.</param>
    Public Function InvertCondition(condition As ExpressionSyntax) As ExpressionSyntax
        If TypeOf condition Is ParenthesizedExpressionSyntax Then
            Return SyntaxFactory.ParenthesizedExpression(InvertCondition(CType(condition, ParenthesizedExpressionSyntax).Expression))
        End If

        If TypeOf condition Is UnaryExpressionSyntax Then
            Dim uOp As UnaryExpressionSyntax = CType(condition, UnaryExpressionSyntax)
            If uOp.IsKind(SyntaxKind.NotExpression) Then
                Return uOp.Operand
            End If
            Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, uOp.OperatorToken, uOp)
        End If

        If TypeOf condition Is BinaryExpressionSyntax Then
            Dim bOp As BinaryExpressionSyntax = CType(condition, BinaryExpressionSyntax)

            If bOp.IsKind(SyntaxKind.AndExpression) OrElse bOp.IsKind(SyntaxKind.AndAlsoExpression) OrElse bOp.IsKind(SyntaxKind.OrExpression) OrElse bOp.IsKind(SyntaxKind.OrElseExpression) Then
                Dim kind As SyntaxKind = NegateConditionOperator(bOp.Kind())
                Return SyntaxFactory.BinaryExpression(kind, InvertCondition(bOp.Left), SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), InvertCondition(bOp.Right))
            End If

            If bOp.IsKind(SyntaxKind.EqualsExpression) OrElse bOp.IsKind(SyntaxKind.NotEqualsExpression) OrElse bOp.IsKind(SyntaxKind.GreaterThanExpression) OrElse bOp.IsKind(SyntaxKind.GreaterThanOrEqualExpression) OrElse bOp.IsKind(SyntaxKind.LessThanExpression) OrElse bOp.IsKind(SyntaxKind.LessThanOrEqualExpression) Then
                Dim kind As SyntaxKind = NegateRelationalOperator(bOp.Kind())
                Return SyntaxFactory.BinaryExpression(kind, bOp.Left, SyntaxFactory.Token(GetExpressionOperatorTokenKind(kind)), bOp.Right)
            End If

            Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, SyntaxFactory.Token(SyntaxKind.NotKeyword), SyntaxFactory.ParenthesizedExpression(condition))
        End If

        If TypeOf condition Is TernaryConditionalExpressionSyntax Then
            Dim cEx As TernaryConditionalExpressionSyntax = TryCast(condition, TernaryConditionalExpressionSyntax)
            Return cEx.WithCondition(InvertCondition(cEx.Condition))
        End If

        If TypeOf condition Is LiteralExpressionSyntax Then
            If condition.Kind() = SyntaxKind.TrueLiteralExpression Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.FalseLiteralExpression, SyntaxFactory.Token(SyntaxKind.FalseKeyword))
            End If
            If condition.Kind() = SyntaxKind.FalseLiteralExpression Then
                Return SyntaxFactory.LiteralExpression(SyntaxKind.TrueLiteralExpression, SyntaxFactory.Token(SyntaxKind.TrueKeyword))
            End If
        End If

        Return SyntaxFactory.UnaryExpression(SyntaxKind.NotExpression, SyntaxFactory.Token(SyntaxKind.NotKeyword), AddParensForUnaryExpressionIfRequired(condition))
    End Function

    <Extension>
    Public Function IsKind(node As SyntaxNode, kind1 As SyntaxKind, kind2 As SyntaxKind) As Boolean
        If node Is Nothing Then
            Return False
        End If

        Dim vbKind As SyntaxKind = node.Kind()
        Return vbKind = kind1 OrElse vbKind = kind2
    End Function

    <Extension>
    Public Function IsKind(node As SyntaxNode, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind) As Boolean
        If node Is Nothing Then
            Return False
        End If

        Dim vbKind As SyntaxKind = node.Kind()
        Return vbKind = kind1 OrElse vbKind = kind2 OrElse vbKind = kind3
    End Function

    <Extension>
    Public Function IsKind(node As SyntaxNode, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind, kind4 As SyntaxKind) As Boolean
        If node Is Nothing Then
            Return False
        End If

        Dim vbKind As SyntaxKind = node.Kind()
        Return vbKind = kind1 OrElse vbKind = kind2 OrElse vbKind = kind3 OrElse vbKind = kind4
    End Function

    <Extension>
    Public Function IsKind(node As SyntaxNode, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind, kind4 As SyntaxKind, kind5 As SyntaxKind) As Boolean
        If node Is Nothing Then
            Return False
        End If

        Dim vbKind As SyntaxKind = node.Kind()
        Return vbKind = kind1 OrElse vbKind = kind2 OrElse vbKind = kind3 OrElse vbKind = kind4 OrElse vbKind = kind5
    End Function

    ''' <summary>
    ''' Get negation of the condition operator
    ''' </summary>
    ''' <returns>
    ''' negation of the specified condition operator, or BinaryOperatorType.Any if it's not a condition operator
    ''' </returns>
    Public Function NegateConditionOperator(op As SyntaxKind) As SyntaxKind
        Select Case op
            Case SyntaxKind.OrExpression
                Return SyntaxKind.AndExpression
            Case SyntaxKind.OrElseExpression
                Return SyntaxKind.AndAlsoExpression
            Case SyntaxKind.AndExpression
                Return SyntaxKind.OrExpression
            Case SyntaxKind.AndAlsoExpression
                Return SyntaxKind.OrElseExpression
        End Select
        Throw New ArgumentOutOfRangeException(NameOf(op))
    End Function

    ''' <summary>
    ''' Get negation of the specified relational operator
    ''' </summary>
    ''' <returns>
    ''' negation of the specified relational operator, or BinaryOperatorType.Any if it's not a relational operator
    ''' </returns>
    Public Function NegateRelationalOperator(op As SyntaxKind) As SyntaxKind
        Select Case op
            Case SyntaxKind.EqualsExpression
                Return SyntaxKind.NotEqualsExpression
            Case SyntaxKind.NotEqualsExpression
                Return SyntaxKind.EqualsExpression
            Case SyntaxKind.GreaterThanExpression
                Return SyntaxKind.LessThanOrEqualExpression
            Case SyntaxKind.GreaterThanOrEqualExpression
                Return SyntaxKind.LessThanExpression
            Case SyntaxKind.LessThanExpression
                Return SyntaxKind.GreaterThanOrEqualExpression
            Case SyntaxKind.LessThanOrEqualExpression
                Return SyntaxKind.GreaterThanExpression
            Case SyntaxKind.OrExpression
                Return SyntaxKind.AndExpression
            Case SyntaxKind.OrElseExpression
                Return SyntaxKind.AndAlsoExpression
            Case SyntaxKind.AndExpression
                Return SyntaxKind.OrExpression
            Case SyntaxKind.AndAlsoExpression
                Return SyntaxKind.OrElseExpression
        End Select
        Throw New ArgumentOutOfRangeException(NameOf(op))
    End Function

End Module
