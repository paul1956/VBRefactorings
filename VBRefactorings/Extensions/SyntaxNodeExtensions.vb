' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic

Public Module SyntaxNodeExtensions

    <Extension>
    Public Function FirstAncestorOfType(Of T As SyntaxNode)(node As SyntaxNode) As T
        Dim currentNode As SyntaxNode = node
        Do
            Dim parent As SyntaxNode = currentNode.Parent
            If parent Is Nothing Then
                Exit Do
            End If
            Dim tParent As T = TryCast(parent, T)
            If tParent IsNot Nothing Then
                Return tParent
            End If
            currentNode = parent
        Loop
        Return Nothing
    End Function

    <Extension>
    Public Function FirstAncestorOfType(node As SyntaxNode, ParamArray types() As Type) As SyntaxNode
        Dim currentNode As SyntaxNode = node
        Do
            Dim parent As SyntaxNode = currentNode.Parent
            If parent Is Nothing Then
                Exit Do
            End If
            For Each Type As Type In types
                If parent.GetType() Is Type Then
                    Return parent
                End If
            Next Type
            currentNode = parent
        Loop
        Return Nothing
    End Function

    <Extension>
    Public Function FirstAncestorOrSelfOfType(Of T As SyntaxNode)(node As SyntaxNode) As T
        Return CType(node.FirstAncestorOrSelfOfType(GetType(T)), T)
    End Function

    <Extension>
    Public Function FirstAncestorOrSelfOfType(node As SyntaxNode, ParamArray types() As Type) As SyntaxNode
        Dim currentNode As SyntaxNode = node
        Do
            If currentNode Is Nothing Then
                Exit Do
            End If
            For Each Type As Type In types
                If currentNode.GetType() Is Type Then
                    Return currentNode
                End If
            Next Type
            currentNode = currentNode.Parent
        Loop
        Return Nothing
    End Function

    <Extension>
    Public Function GetAncestorOrThis(Of TNode As SyntaxNode)(node As SyntaxNode) As TNode
        If node Is Nothing Then
            Return Nothing
        End If

        Return node.GetAncestorsOrThis(Of TNode)().FirstOrDefault()
    End Function

    <Extension>
    Public Iterator Function GetAncestorsOrThis(Of TNode As SyntaxNode)(node As SyntaxNode) As IEnumerable(Of TNode)
        Dim current As SyntaxNode = node
        While current IsNot Nothing
            If TypeOf current Is TNode Then
                Yield DirectCast(current, TNode)
            End If

            current = If(TypeOf current Is IStructuredTriviaSyntax, DirectCast(current, IStructuredTriviaSyntax).ParentTrivia.Token.Parent, current.Parent)
        End While
    End Function

    <Extension>
    Public Function IsKind(token As SyntaxToken, kind1 As SyntaxKind, kind2 As SyntaxKind) As Boolean
        Return Kind(token) = kind1 OrElse Kind(token) = kind2
    End Function

    <Extension>
    Public Function IsParentKind(node As SyntaxNode, kind As SyntaxKind) As Boolean
        Return node IsNot Nothing AndAlso node.Parent.IsKind(kind)
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, ParamArray trivia As SyntaxTrivia()) As T
        If trivia.Length = 0 Then
            Return node
        End If

        Return node.WithAppendedTrailingTrivia(DirectCast(trivia, IEnumerable(Of SyntaxTrivia)))
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, trivia As SyntaxTriviaList) As T
        If trivia.Count = 0 Then
            Return node
        End If

        Return node.WithTrailingTrivia(node.GetTrailingTrivia().Concat(trivia))
    End Function

    <Extension>
    Public Function WithAppendedTrailingTrivia(Of T As SyntaxNode)(node As T, trivia As IEnumerable(Of SyntaxTrivia)) As T
        Return node.WithAppendedTrailingTrivia(trivia.ToSyntaxTriviaList())
    End Function

End Module
