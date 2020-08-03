' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Public Module AnalyzerExtensions

    <Extension>
    Public Iterator Function AllBaseTypes(ByVal typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
        Do While typeSymbol.BaseType IsNot Nothing
            Yield typeSymbol.BaseType
            typeSymbol = typeSymbol.BaseType
        Loop
    End Function

    <Extension>
    Public Iterator Function AllBaseTypesAndSelf(ByVal typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
        Yield typeSymbol
        For Each b As INamedTypeSymbol In AllBaseTypes(typeSymbol)
            Yield b
        Next
    End Function

    <Extension>
    Public Function FirstAncestorOfType(Of T As SyntaxNode)(ByVal node As SyntaxNode) As T
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
    Public Function FirstAncestorOfType(ByVal node As SyntaxNode, ParamArray ByVal types() As Type) As SyntaxNode
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
    Public Function FirstAncestorOrSelfOfType(Of T As SyntaxNode)(ByVal node As SyntaxNode) As T
        Return CType(node.FirstAncestorOrSelfOfType(GetType(T)), T)
    End Function

    <Extension>
    Public Function FirstAncestorOrSelfOfType(ByVal node As SyntaxNode, ParamArray ByVal types() As Type) As SyntaxNode
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
    Public Function GetAllMethodsIncludingFromInnerTypes(typeSymbol As INamedTypeSymbol) As IList(Of IMethodSymbol)
        Dim methods As List(Of IMethodSymbol) = typeSymbol.GetMembers().OfType(Of IMethodSymbol)().ToList()
        Dim innerTypes As IEnumerable(Of INamedTypeSymbol) = typeSymbol.GetMembers().OfType(Of INamedTypeSymbol)()
        For Each innerType As INamedTypeSymbol In innerTypes
            methods.AddRange(innerType.GetAllMethodsIncludingFromInnerTypes())
        Next
        Return methods
    End Function

    <Extension>
    Public Function IsAnyKind(ByVal displayPart As SymbolDisplayPart, ParamArray ByVal kinds() As SymbolDisplayPartKind) As Boolean
        For Each Kind As SymbolDisplayPartKind In kinds
            If displayPart.Kind = Kind Then
                Return True
            End If
        Next Kind
        Return False
    End Function

End Module