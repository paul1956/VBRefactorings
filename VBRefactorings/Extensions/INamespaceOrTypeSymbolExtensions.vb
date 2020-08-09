' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis

Friend Module INamespaceOrTypeSymbolExtensions

    Private Sub GetNameParts(namespaceOrTypeSymbol As INamespaceOrTypeSymbol, result As List(Of String))
        If namespaceOrTypeSymbol Is Nothing OrElse (namespaceOrTypeSymbol.IsNamespace AndAlso DirectCast(namespaceOrTypeSymbol, INamespaceSymbol).IsGlobalNamespace) Then
            Return
        End If

        GetNameParts(namespaceOrTypeSymbol.ContainingNamespace, result)
        result.Add(namespaceOrTypeSymbol.Name)
    End Sub

    <Extension>
    Public Iterator Function AllBaseTypes(typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
        Do While typeSymbol.BaseType IsNot Nothing
            Yield typeSymbol.BaseType
            typeSymbol = typeSymbol.BaseType
        Loop
    End Function

    <Extension>
    Public Iterator Function AllBaseTypesAndSelf(typeSymbol As INamedTypeSymbol) As IEnumerable(Of INamedTypeSymbol)
        Yield typeSymbol
        For Each b As INamedTypeSymbol In AllBaseTypes(typeSymbol)
            Yield b
        Next
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

End Module
