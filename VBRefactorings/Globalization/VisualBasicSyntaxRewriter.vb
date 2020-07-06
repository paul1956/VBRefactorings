' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Public Class VisualBasicConstantStringReplacement
    Inherits VisualBasicSyntaxVisitor(Of SyntaxNode)

    Private ReadOnly ResourceClass As ResourceXClass

    Public Sub New(ResxClass As ResourceXClass)
        ResourceClass = ResxClass
    End Sub

    Private Function ConvertStringLineralToResource(node As ExpressionSyntax) As ExpressionSyntax
        Dim ResourceText As String
        ResourceText = GetResourceText(node)
        Dim identifierNameString As String
        identifierNameString = ResourceClass.GetResourceNameFromValue(ResourceText)
        If String.IsNullOrWhiteSpace(identifierNameString) Then
            identifierNameString = ResourceClass.GetResourceName(ResourceText)
            If Not ResourceClass.AddToResourceFile(identifierNameString, ResourceText) Then
                Return node
            End If
        End If
        Return SyntaxFactory.ParseExpression($"My.Resources.{ResourceClass.ResourceFileNameWithoutExtension}.{identifierNameString}").WithTriviaFrom(node)
    End Function

    Public Overrides Function VisitBinaryExpression(node As BinaryExpressionSyntax) As SyntaxNode
        If Not node.IsKind(SyntaxKind.ConcatenateExpression) Then
            Return MyBase.VisitBinaryExpression(node)
        End If
        Return SyntaxFactory.ConcatenateExpression(CType(node.Left.Accept(Me).WithTriviaFrom(node.Left), ExpressionSyntax), CType(node.Right.Accept(Me).WithTriviaFrom(node.Right), ExpressionSyntax))
    End Function

    Public Overrides Function VisitInterpolatedStringExpression(node As InterpolatedStringExpressionSyntax) As SyntaxNode
        Return ConvertStringLineralToResource(node)
    End Function

    Public Overrides Function VisitLiteralExpression(node As LiteralExpressionSyntax) As SyntaxNode
        Return ConvertStringLineralToResource(node)
    End Function

End Class