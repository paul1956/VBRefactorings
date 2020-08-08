' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis

Friend Module INamespaceOrTypeSymbolExtensions

    Private Sub GetNameParts(namespaceOrTypeSymbol As INamespaceOrTypeSymbol, result As List(Of String))
        If namespaceOrTypeSymbol Is Nothing OrElse (namespaceOrTypeSymbol.IsNamespace AndAlso DirectCast(namespaceOrTypeSymbol, INamespaceSymbol).IsGlobalNamespace) Then
            Return
        End If

        GetNameParts(namespaceOrTypeSymbol.ContainingNamespace, result)
        result.Add(namespaceOrTypeSymbol.Name)
    End Sub

End Module
