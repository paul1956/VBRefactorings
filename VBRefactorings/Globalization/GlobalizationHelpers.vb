' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Text

Public Module GlobalizationHelpers
    Public Function MoveStringToResourceFile(ResourceClass As ResourceXClass, root As SyntaxNode, Expression As ExpressionSyntax) As SyntaxNode
        Dim Rewriter As New VisualBasicConstantStringReplacement(ResourceClass)
        Return root.ReplaceNode(Expression, Rewriter.Visit(Expression))
    End Function

    Public Function GetResourceText(Expression As ExpressionSyntax) As String
        Select Case Expression.Kind
            Case SyntaxKind.StringLiteralExpression
                Dim InvicationRightToken As SyntaxToken = CType(Expression, LiteralExpressionSyntax).Token
                Return InvicationRightToken.ValueText.Replace("""", """""")
            Case SyntaxKind.InterpolatedStringExpression
                Return CType(Expression, InterpolatedStringExpressionSyntax).Contents.ToString
            Case Else
                Throw New ArgumentException("Expression is not StringLiteralExpression or InterpolatedStringExpression")
        End Select
    End Function

    Public Function InitializeResourceClasses(ResourceClasses As Dictionary(Of String, ResourceXClass), document As Document) As ResourceXClass
        Dim ResourceClass As ResourceXClass = Nothing
        Dim ResourceFileWithPath As String = document.Project.FilePath
        If String.IsNullOrWhiteSpace(ResourceFileWithPath) Then
            Dim NewResourceXClass As ResourceXClass = New ResourceXClass()
            ResourceClasses.Add(String.Empty, NewResourceXClass)
            Return NewResourceXClass
        End If
        If Not ResourceClasses.TryGetValue(ResourceFileWithPath, ResourceClass) Then
            Dim NewResourceXClass As ResourceXClass = New ResourceXClass(ResourceFileWithPath)
            If NewResourceXClass.Initialized = ResourceXClass.InitializedValues.NoResourceDirectory Then
                Return NewResourceXClass
            Else
                ResourceClasses.Add(ResourceFileWithPath, NewResourceXClass)
                ResourceClass = ResourceClasses(ResourceFileWithPath)
            End If
        End If

        Return ResourceClass
    End Function

End Module
