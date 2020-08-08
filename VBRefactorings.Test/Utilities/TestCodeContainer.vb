﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

''' <summary> Helper class to bundle together information about a piece of analyzed test code. </summary>
Class TestCodeContainer

    Public Property Position As Integer

    Public Property Text As String

    Public Property SyntaxTree As SyntaxTree

    Public Property Token As SyntaxToken

    Public Property SyntaxNode As SyntaxNode

    Public Property Compilation As Compilation

    Public Property SemanticModel As SemanticModel

    Public Sub New(textWithMarker As String)
        Position = textWithMarker.IndexOf("$"c)
        If Position <> -1 Then
            textWithMarker = textWithMarker.Remove(Position, 1)
        End If

        Text = textWithMarker
        SyntaxTree = VisualBasic.SyntaxFactory.ParseSyntaxTree(Text)
        If Position <> -1 Then
            Token = SyntaxTree.GetRoot().FindToken(Position)
            SyntaxNode = Token.Parent
        End If

        ' Use the mscorlib from our current process
        Compilation = VisualBasicCompilation.Create(
            "test",
            syntaxTrees:={SyntaxTree},
            SharedReferences.VisualBasicReferences(""))

        SemanticModel = Compilation.GetSemanticModel(SyntaxTree)
    End Sub

    Public Sub GetStatementsBetweenMarkers(ByRef firstStatement As StatementSyntax, ByRef lastStatement As StatementSyntax)
        Dim span As TextSpan = Me.GetSpanBetweenMarkers()
        Dim statementsInside As IEnumerable(Of StatementSyntax) = SyntaxTree.GetRoot().DescendantNodes(span).OfType(Of StatementSyntax).Where(Function(s) span.Contains(s.Span))
        Dim first As StatementSyntax = statementsInside.First()
        firstStatement = first
        lastStatement = statementsInside.Where(Function(s) s.Parent Is first.Parent).Last()
    End Sub

    Public Function GetSpanBetweenMarkers() As TextSpan
        Dim startComment As SyntaxTrivia = SyntaxTree.GetRoot().DescendantTrivia().First(Function(t) t.ToString().Contains("start"))
        Dim endComment As SyntaxTrivia = SyntaxTree.GetRoot().DescendantTrivia().First(Function(t) t.ToString().Contains("end"))
        Dim span As TextSpan = TextSpan.FromBounds(startComment.FullSpan.End, endComment.FullSpan.Start)
        Return span
    End Function

End Class
